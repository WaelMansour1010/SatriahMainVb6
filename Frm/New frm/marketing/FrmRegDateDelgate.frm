VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmRegDateDelgate 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   ‘«‘…  ”ÃÌ· „Ê«⁄Ìœ «·“Ì«—« "
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   Icon            =   "FrmRegDateDelgate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   17070
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   3435
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   480
      Width           =   17415
      Begin VB.Timer Timer1 
         Left            =   5280
         Top             =   1380
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·“Ì«—…"
         Height          =   2775
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   600
         Width           =   4935
         Begin VB.TextBox TxtRemark1 
            Alignment       =   1  'Right Justify
            Height          =   1485
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   104
            Top             =   1080
            Width           =   3825
         End
         Begin MSComCtl2.DTPicker DateVisit1 
            Height          =   315
            Left            =   2460
            TabIndex        =   92
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   213712897
            CurrentDate     =   41640
         End
         Begin MSDataListLib.DataCombo DcbFrom1 
            Bindings        =   "FrmRegDateDelgate.frx":038A
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   2460
            TabIndex        =   94
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DcbTO1 
            Bindings        =   "FrmRegDateDelgate.frx":039F
            Height          =   315
            Left            =   990
            TabIndex        =   95
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo DcbTypeVisit1 
            Height          =   315
            Left            =   0
            TabIndex        =   98
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„·«ÕŸ« "
            Height          =   285
            Index           =   10
            Left            =   3450
            TabIndex        =   105
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄Â«"
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   99
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰ «·”«⁄Â "
            Height          =   285
            Index           =   32
            Left            =   3810
            TabIndex        =   97
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   285
            Index           =   12
            Left            =   1860
            TabIndex        =   96
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·“Ì«—…"
            Height          =   285
            Index           =   23
            Left            =   3690
            TabIndex        =   93
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·⁄„Ì·"
         Height          =   2775
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   690
         Width           =   11685
         Begin VB.OptionButton optInvType 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄„Ì·"
            Height          =   285
            Index           =   0
            Left            =   9090
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   150
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton optInvType 
            Alignment       =   1  'Right Justify
            Caption         =   "„‘—Ê⁄"
            Height          =   285
            Index           =   1
            Left            =   7260
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdApi 
            Caption         =   "Load From Web"
            Height          =   600
            Left            =   10845
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   270
            Width           =   780
         End
         Begin VB.Frame frmProgress 
            Height          =   1935
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   2505
            Visible         =   0   'False
            Width           =   4335
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   615
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Ì „ Ã·» «·„Êﬁ⁄"
               Height          =   285
               Index           =   40
               Left            =   900
               TabIndex        =   141
               Top             =   165
               Width           =   1785
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Êﬁ  Ê„Êﬁ⁄ «‰ Â«¡ «·“Ì«—…"
            Height          =   1575
            Index           =   1
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   1020
            Width           =   5595
            Begin VB.CommandButton cmdStartVisit 
               Caption         =   "‰Â«Ì… «·“Ì«—…"
               Height          =   435
               Index           =   1
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   300
               Width           =   1335
            End
            Begin VB.TextBox txtGPS 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   133
               Top             =   780
               Width           =   4815
            End
            Begin VB.TextBox txtAddress 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   132
               Top             =   1200
               Width           =   4815
            End
            Begin MSComCtl2.DTPicker txtDateVis 
               Height          =   435
               Index           =   1
               Left            =   3900
               TabIndex        =   134
               Top             =   300
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   767
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   213712897
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker txtVisitTime 
               Height          =   420
               Index           =   1
               Left            =   1740
               TabIndex        =   135
               Top             =   300
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   213712899
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "GPS"
               Height          =   285
               Index           =   38
               Left            =   4380
               TabIndex        =   137
               Top             =   780
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   285
               Index           =   29
               Left            =   4500
               TabIndex        =   136
               Top             =   1200
               Width           =   945
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Êﬁ  Ê„Êﬁ⁄ »œ¡ «·“Ì«—…"
            Height          =   1575
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1080
            Width           =   5595
            Begin VB.CommandButton cmdStartVisit 
               Caption         =   "»œ¡ «·“Ì«—…"
               Height          =   435
               Index           =   0
               Left            =   195
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   300
               Width           =   1335
            End
            Begin VB.TextBox txtAddress 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   128
               Top             =   1200
               Width           =   4815
            End
            Begin VB.TextBox txtGPS 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   126
               Top             =   780
               Width           =   4815
            End
            Begin MSComCtl2.DTPicker txtDateVis 
               Height          =   435
               Index           =   0
               Left            =   3900
               TabIndex        =   129
               Top             =   285
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   767
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   213712897
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker txtVisitTime 
               Height          =   420
               Index           =   0
               Left            =   1740
               TabIndex        =   130
               Top             =   300
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   213712899
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   285
               Index           =   39
               Left            =   4500
               TabIndex        =   127
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "GPS"
               Height          =   285
               Index           =   31
               Left            =   4380
               TabIndex        =   125
               Top             =   780
               Width           =   945
            End
         End
         Begin VB.TextBox DcbJobID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            TabIndex        =   109
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷ «·„Êﬁ⁄"
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   2700
            Width           =   975
         End
         Begin VB.TextBox TxtMap 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   102
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtAdres 
            Alignment       =   1  'Right Justify
            Height          =   555
            Left            =   -360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   101
            Top             =   2580
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox TxtEnter 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2280
            TabIndex        =   89
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TxtEmail 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox TxtMobi 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtTel 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            TabIndex        =   70
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox TxtPersonCont 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6240
            TabIndex        =   69
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox TxtCustomer 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8640
            TabIndex        =   67
            Top             =   480
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DcbJobID1 
            Height          =   315
            Left            =   3480
            TabIndex        =   78
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   255
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCustomer 
            Height          =   315
            Left            =   6240
            TabIndex        =   110
            Top             =   450
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«Œ·Ì"
            Height          =   285
            Index           =   26
            Left            =   3000
            TabIndex        =   90
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊŸÌ›…"
            Height          =   285
            Index           =   21
            Left            =   5280
            TabIndex        =   77
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
            Height          =   285
            Index           =   20
            Left            =   2160
            TabIndex        =   76
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«·"
            Height          =   285
            Index           =   19
            Left            =   1800
            TabIndex        =   75
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·›Ê‰"
            Height          =   285
            Index           =   18
            Left            =   5640
            TabIndex        =   74
            Top             =   240
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘Œ’ «·„”∆Ê·"
            Height          =   285
            Index           =   9
            Left            =   9600
            TabIndex        =   73
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   0
            Left            =   9600
            TabIndex        =   68
            Top             =   480
            Width           =   1125
         End
      End
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "FrmRegDateDelgate.frx":03B4
         Left            =   18840
         List            =   "FrmRegDateDelgate.frx":03BE
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   14610
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   150
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmRegDateDelgate.frx":03CC
         Height          =   315
         Left            =   6480
         TabIndex        =   48
         Top             =   150
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   390
         TabIndex        =   65
         Top             =   150
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   12120
         TabIndex        =   86
         Top             =   150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   213647361
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   -1320
         TabIndex        =   55
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«œŒ«·"
         Height          =   285
         Index           =   1
         Left            =   13470
         TabIndex        =   52
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ«∆„ »«·“Ì«—…"
         Height          =   285
         Index           =   3
         Left            =   5310
         TabIndex        =   51
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·“Ì«—…"
         Height          =   285
         Index           =   4
         Left            =   16110
         TabIndex        =   50
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13410
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14190
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   19470
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   14190
      TabIndex        =   38
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
      Width           =   17085
      _cx             =   30136
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
      Caption         =   "   ‘«‘…  ”ÃÌ· „Ê«⁄Ìœ «·“Ì«—«   "
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
         ButtonImage     =   "FrmRegDateDelgate.frx":03E1
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
         ButtonImage     =   "FrmRegDateDelgate.frx":077B
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
         ButtonImage     =   "FrmRegDateDelgate.frx":0B15
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
         ButtonImage     =   "FrmRegDateDelgate.frx":0EAF
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
         Picture         =   "FrmRegDateDelgate.frx":1249
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
      Left            =   5070
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9900
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
      Left            =   11100
      TabIndex        =   13
      Top             =   9480
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   5475
      Left            =   0
      TabIndex        =   23
      Top             =   4020
      Width           =   17040
      _cx             =   30057
      _cy             =   9657
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
      BackColor       =   14871017
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "»Ì«‰«  «·“Ì«—« |11|„Êﬁ› «·›Ê« Ì—|Õ«·Â «·«⁄ „«œ|«·÷„«‰"
      Align           =   0
      CurrTab         =   4
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
      Picture(0)      =   "FrmRegDateDelgate.frx":4EB1
      Flags(1)        =   2
      Flags(3)        =   2
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   5010
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   173
         Top             =   45
         Width           =   16950
         Begin VB.CommandButton Command5 
            BackColor       =   &H000000FF&
            Caption         =   "X"
            Height          =   375
            Left            =   16440
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   120
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid GrdWa 
            Height          =   4095
            Left            =   450
            TabIndex        =   176
            Top             =   720
            Width           =   16140
            _cx             =   28469
            _cy             =   7223
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
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegDateDelgate.frx":524B
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·“Ì«—«  «·Œ«’Â »«·⁄„Ì·"
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
            Height          =   255
            Index           =   43
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   240
            Width           =   7365
         End
      End
      Begin VB.TextBox StrCusID 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   12840
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1995
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5010
         Index           =   0
         Left            =   -18195
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   16950
         _cx             =   29898
         _cy             =   8837
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3630
            Left            =   120
            TabIndex        =   25
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   6403
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
            FormatString    =   $"FrmRegDateDelgate.frx":5551
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5010
         Index           =   15
         Left            =   -18495
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   16950
         _cx             =   29898
         _cy             =   8837
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   12
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6000
            Index           =   16
            Left            =   15
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   15
            Width           =   17520
            _cx             =   30903
            _cy             =   10583
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
            Appearance      =   5
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «· ⁄«ﬁœ"
               Height          =   1065
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   2520
               Width           =   10575
               Begin VB.TextBox TxtBillNo 
                  Alignment       =   1  'Right Justify
                  Height          =   555
                  Left            =   4680
                  TabIndex        =   122
                  Top             =   240
                  Width           =   4335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·›« Ê—…"
                  Height          =   405
                  Index           =   35
                  Left            =   9120
                  TabIndex        =   123
                  Top             =   360
                  Width           =   1365
               End
            End
            Begin VB.CommandButton Command2 
               Caption         =   " ›«’Ì· «·“Ì«—« "
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   1905
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   3240
               Width           =   16935
               Begin VB.CommandButton bClose 
                  BackColor       =   &H000000FF&
                  Caption         =   "X"
                  Height          =   375
                  Left            =   16440
                  Style           =   1  'Graphical
                  TabIndex        =   119
                  Top             =   120
                  Width           =   375
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   1275
                  Left            =   0
                  TabIndex        =   117
                  Top             =   420
                  Width           =   16875
                  _cx             =   29766
                  _cy             =   2249
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
                  Rows            =   50
                  Cols            =   18
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmRegDateDelgate.frx":569D
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·“Ì«—«  «·Œ«’Â »«·⁄„Ì·"
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
                  Height          =   255
                  Index           =   37
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   240
                  Width           =   7365
               End
            End
            Begin VB.TextBox txtnotAccept 
               Alignment       =   1  'Right Justify
               Height          =   945
               Left            =   6360
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   114
               Top             =   2640
               Width           =   8625
            End
            Begin VB.TextBox TxtLongTime 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   2895
               TabIndex        =   79
               Top             =   5865
               Visible         =   0   'False
               Width           =   3645
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„«»⁄œ «·“Ì«—…"
               Height          =   1995
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   3030
               Width           =   16905
               Begin VB.TextBox TxtRemark2 
                  Alignment       =   1  'Right Justify
                  Height          =   945
                  Left            =   6600
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   113
                  Top             =   840
                  Width           =   8625
               End
               Begin MSDataListLib.DataCombo DcbTypeVisit2 
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   57
                  Top             =   360
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker XpDtbVisit 
                  Height          =   315
                  Left            =   6600
                  TabIndex        =   60
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   184680449
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbFrom2 
                  Bindings        =   "FrmRegDateDelgate.frx":5949
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   81
                  Top             =   360
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  ListField       =   "account_name"
                  BoundColumn     =   "code"
                  Text            =   ""
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
               Begin MSDataListLib.DataCombo DcbTO2 
                  Bindings        =   "FrmRegDateDelgate.frx":595E
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   82
                  Top             =   360
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  ListField       =   "account_name"
                  BoundColumn     =   "code"
                  Text            =   ""
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
               Begin MSDataListLib.DataCombo DcbSpecialAs 
                  Height          =   315
                  Left            =   12480
                  TabIndex        =   106
                  Top             =   360
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬁÌÌ„ «·⁄„Ì·"
                  Height          =   285
                  Index           =   24
                  Left            =   15720
                  TabIndex        =   107
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   285
                  Index           =   17
                  Left            =   2400
                  TabIndex        =   64
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   16
                  Left            =   4080
                  TabIndex        =   63
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   15
                  Left            =   7470
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”«⁄…"
                  Height          =   285
                  Index           =   14
                  Left            =   4440
                  TabIndex        =   61
                  Top             =   360
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„·«ÕŸ« "
                  Height          =   285
                  Index           =   13
                  Left            =   15600
                  TabIndex        =   59
                  Top             =   1080
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ŒÿÊ… «· «·Ì…"
                  Height          =   285
                  Index           =   5
                  Left            =   11400
                  TabIndex        =   58
                  Top             =   360
                  Width           =   1005
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   645
               Left            =   300
               TabIndex        =   36
               Top             =   6000
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   1138
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   1635
               Left            =   8760
               TabIndex        =   54
               Top             =   360
               Width           =   8115
               _cx             =   14314
               _cy             =   2884
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
               Rows            =   50
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmRegDateDelgate.frx":5973
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
            Begin VSFlex8UCtl.VSFlexGrid Fg2 
               Height          =   1635
               Left            =   120
               TabIndex        =   83
               Top             =   390
               Width           =   8475
               _cx             =   14949
               _cy             =   2884
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
               Rows            =   50
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmRegDateDelgate.frx":5AA9
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   21
               Left            =   9600
               TabIndex        =   84
               Top             =   2070
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmRegDateDelgate.frx":5BDA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   8
               Left            =   0
               TabIndex        =   85
               Top             =   2100
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmRegDateDelgate.frx":6174
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin XtremeSuiteControls.CheckBox ChekAccept 
               Height          =   375
               Left            =   15840
               TabIndex        =   108
               Top             =   2040
               Width           =   975
               _Version        =   786432
               _ExtentX        =   1720
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " „ «·“Ì«—…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox CHekNotAccept 
               Height          =   375
               Left            =   14040
               TabIndex        =   111
               Top             =   2040
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " „ ≈·€«¡ «·“Ì«—…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChekContracted 
               Height          =   375
               Left            =   12240
               TabIndex        =   112
               Top             =   2040
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " „ «· ⁄«ﬁœ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”»» «·€«¡ «·“Ì«—…"
               Height          =   525
               Index           =   36
               Left            =   15120
               TabIndex        =   115
               Top             =   2640
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„ ÿ·»« "
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
               Height          =   255
               Index           =   34
               Left            =   5010
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   120
               Width           =   3645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ÿ«ﬁ„ «·„ÿ·Ê»"
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
               Height          =   255
               Index           =   33
               Left            =   12810
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   120
               Width           =   3645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì·"
               Height          =   450
               Index           =   22
               Left            =   6075
               TabIndex        =   80
               Top             =   5550
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3465
               Index           =   62
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1605
               Width           =   675
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6000
            Index           =   9
            Left            =   15
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   15
            Width           =   17640
            _cx             =   31115
            _cy             =   10583
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   4500
               Left            =   4560
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   1290
               Width           =   1155
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3105
               Left            =   5925
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   1605
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3105
               Index           =   67
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1605
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3000
               Index           =   68
               Left            =   5715
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   2040
               Width           =   30
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3615
               Index           =   69
               Left            =   4185
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1605
               Width           =   375
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5010
         Index           =   1
         Left            =   -17895
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   45
         Width           =   16950
         _cx             =   29898
         _cy             =   8837
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
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ê—œÌ‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„·«¡"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   4020
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   30
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ﬂ"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„” √Ã—Ì‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   -240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "⁄—÷ «·„Êﬁ›"
            Height          =   495
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   -60
            Width           =   5715
         End
         Begin VSFlex8Ctl.VSFlexGrid grdAging 
            Height          =   2010
            Left            =   60
            TabIndex        =   144
            Top             =   420
            Width           =   16770
            _cx             =   29580
            _cy             =   3545
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegDateDelgate.frx":670E
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   1560
            Left            =   2520
            TabIndex        =   145
            Top             =   2640
            Width           =   12945
            _cx             =   22834
            _cy             =   2752
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
            Rows            =   50
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRegDateDelgate.frx":6A8E
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
         Begin VSFlex8Ctl.VSFlexGrid grdAging2 
            Height          =   1470
            Left            =   2880
            TabIndex        =   148
            Top             =   -60
            Visible         =   0   'False
            Width           =   10050
            _cx             =   17727
            _cy             =   2593
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegDateDelgate.frx":6D97
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
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   345
            Left            =   0
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   185270275
            CurrentDate     =   37140
         End
         Begin XtremeSuiteControls.CheckBox ChekCustomer 
            Height          =   375
            Left            =   1080
            TabIndex        =   155
            Top             =   0
            Visible         =   0   'False
            Width           =   3075
            _Version        =   786432
            _ExtentX        =   5424
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "⁄„Ì·/„Ê—œ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckAllCustomer 
            Height          =   375
            Left            =   0
            TabIndex        =   156
            Top             =   360
            Visible         =   0   'False
            Width           =   4155
            _Version        =   786432
            _ExtentX        =   7329
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ ⁄„Ì· /„Ê—œ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«‰Ã«“« "
            Height          =   195
            Index           =   41
            Left            =   15585
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   2655
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5010
         Index           =   2
         Left            =   -17595
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   45
         Width           =   16950
         _cx             =   29898
         _cy             =   8837
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
         Begin VB.CommandButton Command4 
            Caption         =   "⁄—÷ «·„Êﬁ›"
            Height          =   495
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   -60
            Width           =   5715
         End
         Begin VB.OptionButton Option7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„” √Ã—Ì‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   -240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ﬂ"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„·«¡"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   4020
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   30
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ê—œÌ‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   2010
            Left            =   60
            TabIndex        =   166
            Top             =   420
            Width           =   16770
            _cx             =   29580
            _cy             =   3545
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegDateDelgate.frx":709F
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
            Height          =   1560
            Left            =   2520
            TabIndex        =   167
            Top             =   2640
            Width           =   12945
            _cx             =   22834
            _cy             =   2752
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
            Rows            =   50
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRegDateDelgate.frx":741F
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
            Height          =   1470
            Left            =   2880
            TabIndex        =   168
            Top             =   -60
            Visible         =   0   'False
            Width           =   10050
            _cx             =   17727
            _cy             =   2593
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRegDateDelgate.frx":7728
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   0
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   185270275
            CurrentDate     =   37140
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   375
            Left            =   1080
            TabIndex        =   170
            Top             =   0
            Visible         =   0   'False
            Width           =   3075
            _Version        =   786432
            _ExtentX        =   5424
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "⁄„Ì·/„Ê—œ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckBox2 
            Height          =   375
            Left            =   0
            TabIndex        =   171
            Top             =   360
            Visible         =   0   'False
            Width           =   4155
            _Version        =   786432
            _ExtentX        =   7329
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ ⁄„Ì· /„Ê—œ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«‰Ã«“« "
            Height          =   195
            Index           =   42
            Left            =   15585
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   2655
            Width           =   1080
         End
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13470
      TabIndex        =   43
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
      TabIndex        =   44
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
      TabIndex        =   100
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
      TabIndex        =   45
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
      Left            =   13845
      TabIndex        =   18
      Top             =   9555
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
      Top             =   9630
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
      Top             =   9630
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   15
      Top             =   9180
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4260
      TabIndex        =   14
      Top             =   9660
      Width           =   615
   End
End
Attribute VB_Name = "FrmRegDateDelgate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public bol As Boolean
Public novalue As Boolean

Private Sub cmdApi_Click()
    Dim Req  As New WinHttp.WinHttpRequest
    '   "customercode": "1121J1015",
    '        "employeeid": 4,
    '        "customername": "Al Faisal University -Riyadh J1015",
    '        "employeename": "\nTalal Issa Muhammad Arafat",
    '        "date": "21/02/2024",
    '        "singintime": "21/02/2024 00:43",
    '        "signouttime": "21/02/2024 00:43",
    '        "notes": "??????? ??????",
    '        "signinlocation": "30.0580864,31.342592",
    '        "signoutlocation": "30.0580864,31.342592",
    '        "rowId": "d15f9e26-d988-4984-a9b4-318a29c1d0d7"
    Dim intX As Long, Num As Long
    Dim AllDes
    Dim EmpID As Integer
    Dim Row   As Integer
    Dim strFilterText, strFilterText1
    Dim NooFRows, StrSQL
    Dim RsDetails As ADODB.Recordset
    Dim rsDummy   As ADODB.Recordset
    
    Dim moption
    moption = Req.Option(WinHttpRequestOption_SslErrorIgnoreFlags)
    moption = moption Or SslErrorFlag_Ignore_All
    Req.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = moption

    '    //SslErrorFlag_Ignore_All                                  0x3300
    '    //Unknown certification authority (CA) or untrusted root   0x0100
    '    //Wrong usage                                              0x0200
    '    //Invalid common name (CN)                                 0x1000
    '    //Invalid date or certificate expired                      0x2000
    
    Req.Open "get", APIURL & "/api/empdata/getvisit", async:=False
    Req.setRequestHeader "Content-Type", "application/hal+json"
    Req.setRequestHeader "Accept", "text/*, application/hal+json, application/json"
    Req.send
    
    Dim p As Object
    Dim i
    Set p = JSON.parse(Req.responseText)
    Dim S As String
    If Not (p Is Nothing) Then
        '        If JSON.GetParserErrors <> "" Then
        '            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
        '        Else
        If p.count > 0 Then
            
            frmEmpVacList.mIndex = 1
            frmEmpVacList.Fg2.Visible = True
            frmEmpVacList.Fg2.rows = 1
            Dim Rs As ADODB.Recordset
            For i = 1 To p.count
                Dim itemDic As Dictionary
                Set itemDic = p(i)
                Dim customercode
                Dim EmployeeID
                Dim customername
                Dim customerid
                Dim employeename
                Dim mDate
                Dim singintime
                Dim signouttime
                Dim notes
                Dim signinlocation
                Dim signoutlocation
                Dim RowID
                Dim visitname
                Dim vdate
                Dim time1
                Dim time2
                Dim mCanceled, mVisitDone, mContracted, mType
'                 "name": "??????? ??????",
'        "visitdate": "06/03/2024",
'        "time1": "23:35",
'        "time2": "23:33"
        
        visitname = itemDic("name")
         vdate = itemDic("visitdate")
          time1 = itemDic("time1")
           time2 = itemDic("time2")
        
                customercode = itemDic("customercode")
                EmployeeID = itemDic("employeeid")
                customername = itemDic("customername")
                employeename = itemDic("employeename")
                customerid = itemDic("customerid")
                mDate = itemDic("visitdate")
                singintime = itemDic("singintime")
                signouttime = itemDic("signouttime")
                notes = itemDic("notes")
                signinlocation = itemDic("signinlocation")
                signoutlocation = itemDic("signoutlocation")
                RowID = itemDic("rowId")
                
                mCanceled = itemDic("canceled")
                mVisitDone = itemDic("visitDone")
                mContracted = itemDic("contracted")
                mType = itemDic("type")




                DcbCustomer.BoundText = customerid
                retInfoCustomer
                S = ""
                S = S & "SELECT * "
                S = S & "FROM TblRegDateDelgate "
                S = S & "WHERE RowId = '" & RowID & "';"
                Set Rs = Nothing
                Set Rs = New ADODB.Recordset
                Rs.Open S, Cn, adOpenKeyset, adLockOptimistic
                If Rs.EOF Then
                    Rs.AddNew
                    Rs!ID = CStr(new_id("TblRegDateDelgate", "ID", "", True))
                End If
                Dim visitdate, timevisit
                Rs!customerid = customerid
                Rs!DelgID = EmployeeID
                Rs("VisitID") = val(mType)

                If mCanceled Then
                    Rs("Accept").value = 3
                ElseIf mVisitDone Then
                    Rs("Accept").value = 1
                ElseIf mContracted Then
                    Rs("Accept").value = 2
                
                End If
'
'                If Me.ChekAccept.value = vbChecked Then
'    rs("Accept").value = 1
'      End If
'If Me.ChekContracted.value = vbChecked Then
'    rs("Accept").value = 2
'      End If
'If Me.CHekNotAccept.value = vbChecked Then
'    rs("Accept").value = 3
'      End If
      
                
         
                Rs!RecordDate = mDate
                XPDtbTrans.value = Date 'Format(mDate)
                
                Dim st() As String
                If singintime & "" <> "" Then
                    st = Split(singintime, " ")
                    Rs!DateVis0 = st(0)
                    Rs!VisitTime0 = st(1)
                End If
                If signouttime & "" <> "" Then
                    st = Split(signouttime, " ")
                    Rs!DateVis1 = st(0)
                    Rs!VisitTime1 = st(1)
                End If
                Rs!Remark = notes
               
                Rs!GPS0 = signinlocation
                Rs!GPS1 = signoutlocation
                Rs!BranchID = Current_branch
                Rs("CustomerName") = Me.TxtCustomer
              '  rs("FromTime1") = time1
              '  rs("ToTime1") = time2
                Rs!VisitDate1 = vdate
                
                


                Rs("JobID").value = IIf(Me.DcbJobID.Text = "", "", (Me.DcbJobID.Text))
                Rs!RowID = "{" & RowID & "}"
                Rs.update
                Rs.Close

            Next
            MsgBox "update All  Done "
        Else
            MsgBox "No Data"
        End If
    End If
    On Error Resume Next
   Rs.Resync adAffectAll
   Rs.Requery
End Sub

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(Rs("Posted")) Then
        Rs("Posted") = user_id
        Rs("PostedDate") = Time
    Else
        Rs("Posted") = Null
       Rs("PostedDate") = Time
    End If
   
    Rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(Me.XPTxtID.Text))
End Sub

Private Sub bClose_Click()
Frame6.Visible = False
If Me.ChekAccept.value = xtpChecked Then
Frame2.Visible = True
End If
If Me.ChekContracted.value = xtpChecked Then
Frame5.Visible = True
End If
End Sub

Private Sub ChekAccept_Click()
If Me.ChekAccept.value = vbChecked Then
Me.CHekNotAccept.value = vbUnchecked
Me.ChekContracted.value = vbUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
Me.Frame2.Visible = True
Me.Frame5.Visible = False
Else
Me.Frame2.Visible = False
End If
End Sub
Private Sub RemoveGridRow()

    With Me.FG

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Private Sub RemoveGridRow2()

    With Me.Fg2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub ChekContracted_Click()
If Me.ChekContracted.value = xtpChecked Then
Me.CHekNotAccept.value = xtpUnchecked
Me.ChekAccept.value = xtpUnchecked
lbl(36).Visible = False
Me.txtnotAccept.Visible = False
Me.Frame2.Visible = False
Frame5.Visible = True
Else
Me.Frame5.Visible = False
End If

End Sub

Private Sub CHekNotAccept_Click()
If Me.CHekNotAccept.value = vbChecked Then
Me.Frame2.Visible = False
Me.Frame5.Visible = False
lbl(36).Visible = True
Me.txtnotAccept.Visible = True
Me.ChekAccept.value = vbUnchecked
Me.ChekContracted.value = vbUnchecked
Else
'Me.Frame2.Visible = True
lbl(36).Visible = False
Me.txtnotAccept.Visible = False
End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index
            'Case 8
            'MsgBox Format(TimeFrom1.value, "hh:mm AM/PM")
        Case 8
            RemoveGridRow2
        Case 21
            RemoveGridRow
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            ' Me.DcbTO1.BoundText = 0
            '   Me.DcbTO2.BoundText = 0
            clear_all Me
            ChekAccept.value = xtpUnchecked
            Me.ChekContracted.value = xtpUnchecked
            Me.CHekNotAccept.value = xtpUnchecked
            '  lbl(20).Caption = "0"
            '    lbl(21).Caption = "0"
            ' lbl(31).Caption = "0"
            'lbl(23).Caption = "0"
            
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
    
            Fg2.Clear flexClearScrollable, flexClearEverything
            Fg2.rows = 2

            Me.DCboUserName.BoundText = user_id
            '    TxtPaymentCounts.text = 1
            dcBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
            Else
                Accredit.Caption = " send to Approval   "
            End If
             
        Case 1
            FG.rows = FG.rows + 1
            FG.Enabled = True
            Fg2.rows = Fg2.rows + 1
            Fg2.Enabled = True
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
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

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
        
            Load FrmRegDateDelgSearch
            FrmRegDateDelgSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
            
        Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
        
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
MySQL = "SELECT     TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Namee, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Name1, "
MySQL = MySQL & "                      TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4, TblEmployee_2.Nationality, TblEmployee_2.Emp_Namee1,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee4, dbo.TblRegDateDelgate.Id, dbo.TblRegDateDelgateDails.EmpID,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_ID, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.BranchID, dbo.TblBranchesData.branch_name,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD, TblEmployee_1.Emp_Name AS Emp_NameD,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Name1 AS Emp_Name1D, TblEmployee_1.Emp_Name2 AS Emp_Name2D, TblEmployee_1.Emp_Name3 AS Emp_Name3D,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Name4 AS Emp_Name4D, TblEmployee_1.Nationality AS NationalityD, TblEmployee_1.Emp_Namee AS Emp_NameeD,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee1 AS Emp_Namee1D, TblEmployee_1.Emp_Namee2 AS Emp_Namee2D, TblEmployee_1.Emp_Namee3 AS Emp_Namee3D,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee4 AS Emp_Namee4D, TblEmployee_1.Fullcode AS FullcodeD, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID,"
 MySQL = MySQL & "                     TblTypeVisit_1.name, TblTypeVisit_1.namee, dbo.TblRegDateDelgate.VisitID2, dbo.TblRegDateDelgate.SpAsID, dbo.TblSpeciaAsement.name AS nameSp,"
 MySQL = MySQL & "                     dbo.TblSpeciaAsement.namee AS nameeSp, dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2,"
 MySQL = MySQL & "                     dbo.TblRegDateDelgate.TimeFrom1, dbo.TblRegDateDelgate.TimeFrom2, dbo.TblRegDateDelgate.TimeTo1, dbo.TblRegDateDelgate.TimeTo2,"
 MySQL = MySQL & "                     dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID,"
 MySQL = MySQL & "                     dbo.TblRegDateDelgateDails.remark AS remarkD, dbo.TblRegDateDelgate.LongTime, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgateDails.quantity,"
 MySQL = MySQL & "                     dbo.TblRegDateDelgateDails.Type, dbo.TblCompo.name AS Nmecompo, dbo.TblCompo.namee AS Nmeecompo, TblTypeVisit_2.name AS name2,"
 MySQL = MySQL & "                     TblTypeVisit_2.namee AS namee2, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map, dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept,"
MySQL = MySQL & "                      dbo.TblRegDateDelgate.BillNo, dbo.TblRegDateDelgate.CustomerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.FromTime1,"
 MySQL = MySQL & "                     TblRegTimeDelgate_1.name AS FromTime11, dbo.TblRegDateDelgate.FromTime2, TblRegTimeDelgate_2.name AS FromTime22, dbo.TblRegDateDelgate.ToTime1,"
MySQL = MySQL & "                      TblRegTimeDelgate_3.name AS ToTime11, dbo.TblRegDateDelgate.ToTime2, dbo.TblRegTimeDelgate.name AS ToTime22"
MySQL = MySQL & " FROM         dbo.TblSpeciaAsement RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblTypeVisit TblTypeVisit_2 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblRegTimeDelgate TblRegTimeDelgate_1 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblRegTimeDelgate TblRegTimeDelgate_2 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblRegTimeDelgate RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblRegDateDelgate ON dbo.TblRegTimeDelgate.Id = dbo.TblRegDateDelgate.ToTime2 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.ToTime1 = TblRegTimeDelgate_3.Id ON"
 MySQL = MySQL & "                     TblRegTimeDelgate_2.Id = dbo.TblRegDateDelgate.FromTime2 ON TblRegTimeDelgate_1.Id = dbo.TblRegDateDelgate.FromTime1 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID ON TblTypeVisit_2.Id = dbo.TblRegDateDelgate.VisitID2 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID = TblTypeVisit_1.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblRegDateDelgateDails ON dbo.TblRegDateDelgate.Id = dbo.TblRegDateDelgateDails.DelgID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id ON TblEmployee_2.Emp_ID = dbo.TblRegDateDelgateDails.EmpID ON"
MySQL = MySQL & "                      dbo.TblSpeciaAsement.Id = dbo.TblRegDateDelgate.SpAsID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblRegDateDelgate.BranchID = dbo.TblBranchesData.branch_id"

MySQL = MySQL & "  Where (dbo.TblRegDateDelgate.id =" & val(Me.XPTxtID.Text) & " )"



  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRegDateDelgate.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRegDateDelgate.rpt"
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
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

Private Sub cmdStartVisit_Click(Index As Integer)
frmProgress.Visible = True
txtDateVis(Index).value = Date
txtVisitTime(Index).value = Time

Dim i As Long
For i = 1 To 1000000
    DoEvents
    Label1.Tag = i
    If i = 1 Then
        Label1.Caption = "10%"
    ElseIf i = 100000 Then
        Label1.Caption = "20%"
    ElseIf i = 200000 Then
        Label1.Caption = "30%"
    ElseIf i = 300000 Then
        Label1.Caption = "40%"
         DoEvents
    ElseIf i = 400000 Then
        Label1.Caption = "50%"
    ElseIf i = 500000 Then
        Label1.Caption = "60%"
    ElseIf i = 600000 Then
        Label1.Caption = "70%"
        DoEvents
    ElseIf i = 700000 Then
        Label1.Caption = "80%"
    ElseIf i = 800000 Then
        Label1.Caption = "90%"
    ElseIf i = 900000 Then
        Label1.Caption = "100%"
    ElseIf i = 1000000 Then
        Label1.Caption = " „ Ã·» «·„Êﬁ⁄"
    End If
    
    DoEvents
Next
frmProgress.Visible = False
txtGPS(Index).Text = "24.62799,46.81171"
TxtAddress(Index).Text = "7668 Abdullah Al Wahaibi ,An Noor,Riyadh 1432"
End Sub

Private Sub Command2_Click()
FillGridDetails
Frame6.Visible = True
Frame5.Visible = False
Frame2.Visible = False
End Sub
'Sub FillGridWa()
'Dim StrSQL As String
'Dim i As Integer
'Dim RsDetails As ADODB.Recordset
'Set RsDetails = New ADODB.Recordset
'
'StrSQL = " SELECT TBLRegularMaint.id, TBLRegularMaint.DateOfRegularMaint, TBLRegularMaint.GranteeStartDate, "
'StrSQL = StrSQL & "        TBLRegularMaint.GranteeEndDate, TBLRegularMaint.Done, TBLRegularMaint.DoneDate, "
'StrSQL = StrSQL & "        TblMaintenanceType.name, TblMaintenanceType.namee, TblMaintenanceType.Valuee, "
'StrSQL = StrSQL & "        TblMaintenanceType.REMARKS, TblWarrantyOffer.ProjectID "
'StrSQL = StrSQL & " FROM TBLRegularMaint "
'StrSQL = StrSQL & " INNER JOIN TblWarrantyOffer ON TblWarrantyOffer.id = TBLRegularMaint.WarntID "
'StrSQL = StrSQL & " LEFT OUTER JOIN TblMaintenanceType ON TblMaintenanceType.id = TBLRegularMaint.MaintenanceIDS "
'StrSQL = StrSQL & " WHERE TblWarrantyOffer.ProjectID = " & val(DcbCustomer.BoundText)
'
'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'GrdWa.Clear flexClearScrollable, flexClearEverything
'GrdWa.rows = GrdWa.FixedRows
'
'With GrdWa
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        .rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .rows - 1
'            On Error Resume Next
'            .TextMatrix(i, .ColIndex("Serial")) = i
'            On Error GoTo 0
'
'            .TextMatrix(i, .ColIndex("MainID")) = IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value)
'            .TextMatrix(i, .ColIndex("MaDate")) = IIf(IsNull(RsDetails("DateOfRegularMaint").value), "", RsDetails("DateOfRegularMaint").value)
'
'            If SystemOptions.UserInterface = EnglishInterface Then
'                .TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
'            Else
'                .TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
'            End If
'
'            .TextMatrix(i, .ColIndex("Interval")) = IIf(IsNull(RsDetails("Valuee").value), "", RsDetails("Valuee").value)
'            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("REMARKS").value), "", RsDetails("REMARKS").value)
'
'            If IsNull(RsDetails("Done").value) Then
'                .TextMatrix(i, .ColIndex("StatusVisit")) = ""
'            ElseIf val(RsDetails("Done").value) = 0 Then
'                .TextMatrix(i, .ColIndex("StatusVisit")) = "·„   „"
'            ElseIf val(RsDetails("Done").value) = 1 Then
'                .TextMatrix(i, .ColIndex("StatusVisit")) = " „ "
'            Else
'                .TextMatrix(i, .ColIndex("StatusVisit")) = RsDetails("Done").value
'            End If
'
'            RsDetails.MoveNext
'        Next i
'    End If
'End With
'
'RsDetails.Close
'Set RsDetails = Nothing
'End Sub
Public Sub FillGridWa()

    Dim StrSQL As String
    Dim i As Integer
    Dim RsDetails As ADODB.Recordset

    Set RsDetails = New ADODB.Recordset

    StrSQL = " SELECT TBLRegularMaint.id, TBLRegularMaint.DateOfRegularMaint, TBLRegularMaint.GranteeStartDate, "
    StrSQL = StrSQL & "        TBLRegularMaint.GranteeEndDate, TBLRegularMaint.Done, TBLRegularMaint.DoneDate, "
    StrSQL = StrSQL & "        TBLRegularMaint.MaintenanceIDS, TBLRegularMaint.StatusVisit, "
    StrSQL = StrSQL & "        TblMaintenanceType.name, TblMaintenanceType.namee, TblMaintenanceType.Valuee, "
    StrSQL = StrSQL & "        TblMaintenanceType.REMARKS, TblWarrantyOffer.ProjectID, TblWarrantyOffer.id AS WarntID "
    StrSQL = StrSQL & " FROM TBLRegularMaint "
    StrSQL = StrSQL & " INNER JOIN TblWarrantyOffer ON TblWarrantyOffer.id = TBLRegularMaint.WarntID "
    StrSQL = StrSQL & " LEFT OUTER JOIN TblMaintenanceType ON TblMaintenanceType.id = TBLRegularMaint.MaintenanceIDS "
    StrSQL = StrSQL & " WHERE TblWarrantyOffer.ProjectID = " & val(DcbCustomer.BoundText)
    StrSQL = StrSQL & " ORDER BY TBLRegularMaint.DateOfRegularMaint, TBLRegularMaint.id "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    GrdWa.Clear flexClearScrollable, flexClearEverything
    GrdWa.rows = GrdWa.FixedRows

    With GrdWa
        If Not (RsDetails.BOF Or RsDetails.EOF) Then
            RsDetails.MoveFirst
            .rows = .FixedRows + RsDetails.RecordCount

            For i = .FixedRows To .rows - 1

                .TextMatrix(i, 0) = CStr(i - .FixedRows + 1)

                .TextMatrix(i, .ColIndex("MainID")) = IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value)
                .TextMatrix(i, .ColIndex("MaDate")) = IIf(IsNull(RsDetails("DateOfRegularMaint").value), "", RsDetails("DateOfRegularMaint").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
                Else
                    .TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
                End If

                .TextMatrix(i, .ColIndex("Interval")) = IIf(IsNull(RsDetails("Valuee").value), "", RsDetails("Valuee").value)
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("REMARKS").value), "", RsDetails("REMARKS").value)

                .TextMatrix(i, .ColIndex("StatusVisit")) = GetStatusVisitText(GetStatusVisitFromDone(RsDetails("StatusVisit").value))

                RsDetails.MoveNext
            Next i
        End If
    End With

    RsDetails.Close
    Set RsDetails = Nothing

End Sub


Public Sub SaveGridWa()

    Dim i As Integer
    Dim StrSQL As String

    On Error GoTo ErrHandler

    '??? ??????
    Cn.Execute "DELETE FROM TblRegDateDelgateDailsGrantee WHERE DelgID = " & val(Me.XPTxtID.Text)

    With GrdWa
        For i = .FixedRows To .rows - 1

            If Trim(.TextMatrix(i, .ColIndex("MainID")) & "") <> "" Then

                StrSQL = ""
                StrSQL = StrSQL & " INSERT INTO TblRegDateDelgateDailsGrantee ( "
                StrSQL = StrSQL & " DelgID, Transaction_ID, Transaction_Type, WarntID, ProjectID, "
                StrSQL = StrSQL & " Serial, MainID, MaDate, MainName, Interval, Remarks, StatusVisit, "
                StrSQL = StrSQL & " DateOfRegularMaint, MaintenanceIDS, Done "
                StrSQL = StrSQL & " ) VALUES ( "

                StrSQL = StrSQL & val(Me.XPTxtID.Text) & ", "                      'DelgID
                StrSQL = StrSQL & SqlLong(GetFormTransactionID()) & ", "          'Transaction_ID
                StrSQL = StrSQL & SqlLong(GetFormTransactionType()) & ", "        'Transaction_Type
                StrSQL = StrSQL & SqlLong(GetCurrentWarntID()) & ", "             'WarntID
                StrSQL = StrSQL & val(DcbCustomer.BoundText) & ", "               'ProjectID

                StrSQL = StrSQL & SqlLong(.TextMatrix(i, 0)) & ", "
                StrSQL = StrSQL & SqlLong(.TextMatrix(i, .ColIndex("MainID"))) & ", "
                StrSQL = StrSQL & SqlDateOrNull(.TextMatrix(i, .ColIndex("MaDate"))) & ", "
                StrSQL = StrSQL & SqlText(.TextMatrix(i, .ColIndex("MainName"))) & ", "
                StrSQL = StrSQL & SqlLong(.TextMatrix(i, .ColIndex("Interval"))) & ", "
                StrSQL = StrSQL & SqlText(.TextMatrix(i, .ColIndex("Remarks"))) & ", "
                StrSQL = StrSQL & SqlLong(GetStatusVisitValue(.TextMatrix(i, .ColIndex("StatusVisit")))) & ", "

                StrSQL = StrSQL & SqlDateOrNull(.TextMatrix(i, .ColIndex("MaDate"))) & ", " 'DateOfRegularMaint
                StrSQL = StrSQL & SqlLong(.TextMatrix(i, .ColIndex("MainID"))) & ", "       'MaintenanceIDS ?? MainID ??? ????????
                StrSQL = StrSQL & SqlLong(GetStatusVisitValue(.TextMatrix(i, .ColIndex("StatusVisit"))))

                StrSQL = StrSQL & " ) "

                Cn.Execute StrSQL
            End If
        Next i
    End With
    
    UpdateRegularMaintFromGridWa

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight, "SaveGridWa"
End Sub

Private Function GetDoneValueFromStatusVisit(ByVal V As Variant) As Integer
    Select Case val(V)
        Case 1
            GetDoneValueFromStatusVisit = 0   '·„   „
        Case 2
            GetDoneValueFromStatusVisit = 1   ' „ 
        Case 3
            GetDoneValueFromStatusVisit = 2   '√·€Ì 
        Case Else
            GetDoneValueFromStatusVisit = 0
    End Select
End Function

Public Sub UpdateRegularMaintFromGridWa()

    Dim i As Integer
    Dim MainID As Long
    Dim DoneValue As Integer
    Dim StrSQL As String

    On Error GoTo ErrHandler

    With GrdWa
        For i = .FixedRows To .rows - 1

            MainID = val(.TextMatrix(i, .ColIndex("MainID")))

            If MainID > 0 Then

                DoneValue = GetDoneValueFromStatusVisit(.TextMatrix(i, .ColIndex("StatusVisit")))

                StrSQL = ""
                StrSQL = StrSQL & " UPDATE TBLRegularMaint SET "
                StrSQL = StrSQL & " StatusVisit = " & DoneValue

                If DoneValue = 1 Then
                    StrSQL = StrSQL & ", DoneDate = " & SqlDateOrNull(Now)
                Else
                    StrSQL = StrSQL & ", DoneDate = NULL"
                End If

                StrSQL = StrSQL & " WHERE id = " & MainID

                Cn.Execute StrSQL
            End If
        Next i
    End With

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical + vbMsgBoxRight, "UpdateRegularMaintFromGridWa"
End Sub

Sub FillGridDetails()
Dim StrSQL As String
Dim i As Integer
Dim RsDetails As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblRegDateDelgate.Id, TblEmployee_1.Emp_ID, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD, "
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name AS Emp_NameD, TblEmployee_1.Nationality AS NationalityD, TblEmployee_1.Fullcode AS FullcodeD,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID, TblTypeVisit_1.name, TblTypeVisit_1.namee,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitID2, dbo.TblRegDateDelgate.SpAsID, dbo.TblSpeciaAsement.name AS nameSp, dbo.TblSpeciaAsement.namee AS nameeSp,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2, dbo.TblRegDateDelgate.PersonConc,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.LongTime,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitDate1, TblTypeVisit_2.name AS name2, TblTypeVisit_2.namee AS namee2, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, TblEmployee_1.Emp_Namee, dbo.TblRegDateDelgate.CustomerID,"
StrSQL = StrSQL & "                         dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.ToTime1, dbo.TblRegTimeDelgate.name AS ToTime11,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.FromTime1, TblRegTimeDelgate_2.name AS FromTime11, dbo.TblRegDateDelgate.FromTime2, TblRegTimeDelgate_3.name AS FromTime22,"
StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.ToTime2, TblRegTimeDelgate_1.name AS ToTime22"
StrSQL = StrSQL & "    FROM         dbo.TblRegTimeDelgate RIGHT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_1 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                        dbo.TblRegDateDelgate ON TblRegTimeDelgate_1.Id = dbo.TblRegDateDelgate.ToTime2 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                        dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.FromTime2 = TblRegTimeDelgate_3.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_2 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_2.Id ON"
StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate.Id = dbo.TblRegDateDelgate.ToTime1 LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_2 ON dbo.TblRegDateDelgate.VisitID2 = TblTypeVisit_2.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID = TblTypeVisit_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblSpeciaAsement ON dbo.TblRegDateDelgate.SpAsID = dbo.TblSpeciaAsement.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID"
StrSQL = StrSQL & "    Where (dbo.TblRegDateDelgate.customerid =" & val(Me.DcbCustomer.BoundText) & ")"
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = VSFlexGrid1.FixedRows
With VSFlexGrid1
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("PersonConc")) = IIf(IsNull(RsDetails("PersonConc").value), "", RsDetails("PersonConc").value) ' RsDetails("remark").value
           ' .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CustomerName").value), "", RsDetails("CustomerName").value) 'RsDetails("fullcode").value
            If SystemOptions.UserInterface = EnglishInterface Then
           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
           Else
           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_NameD").value), "", RsDetails("Emp_NameD").value) ' RsDetails("emp_name").value
            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
           End If
            .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(RsDetails("Mobile").value), "", RsDetails("Mobile").value)
             .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDetails("JobID").value), "", RsDetails("JobID").value)
              .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(RsDetails("Tel").value), "", RsDetails("Tel").value)
               .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(RsDetails("Email").value), "", RsDetails("Email").value)
                .TextMatrix(i, .ColIndex("FromTim11")) = IIf(IsNull(RsDetails("FromTime11").value), "", RsDetails("FromTime11").value)
                 .TextMatrix(i, .ColIndex("ToTime11")) = IIf(IsNull(RsDetails("ToTime11").value), "", RsDetails("ToTime11").value)
                  .TextMatrix(i, .ColIndex("Adress")) = IIf(IsNull(RsDetails("Adress").value), "", RsDetails("Adress").value)
                  .TextMatrix(i, .ColIndex("VisitDate1")) = IIf(IsNull(RsDetails("VisitDate1").value), "", RsDetails("VisitDate1").value)
                  DcbTypeVisit1.BoundText = val(IIf(IsNull(RsDetails("VisitID").value), "", RsDetails("VisitID").value))
                  .TextMatrix(i, .ColIndex("VisitID")) = DcbTypeVisit1.Text
                If RsDetails("Accept").value = 0 Then
                .TextMatrix(i, .ColIndex("Accept")) = ""
                End If
                 If RsDetails("Accept").value = 1 Then
                .TextMatrix(i, .ColIndex("Accept")) = " „ «·“Ì«—…"
                End If
                 If RsDetails("Accept").value = 2 Then
                .TextMatrix(i, .ColIndex("Accept")) = " „ «· ⁄«ﬁœ"
                End If
                 If RsDetails("Accept").value = 3 Then
                .TextMatrix(i, .ColIndex("Accept")) = "≈·€«¡ «·“Ì«—…"
                End If
                 .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
                
                
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
End Sub

Private Sub Command3_Click()
print_report66
End Sub

Private Sub DateVisit1_KeyUp(KeyCode As Integer, Shift As Integer)
If TxtModFlg.Text <> "R" Then
If val(Me.DcboEmpName.BoundText) = 0 Then
MsgBox "ÌÃ»  ÕœÌœ «”„ «·„‰œÊ» «Ê·«"
Exit Sub
Else
fileFgtim val(Me.DcboEmpName.BoundText), 0
refiltimdetails val(Me.DcboEmpName.BoundText), 0
End If
End If
End Sub



Private Sub DcbCustomer_Change()
If Me.TxtModFlg.Text <> "R" Then
Me.TxtCustomer.Text = ""
If optInvType(0) Then
    retInfoCustomer
    
End If
FillGridWa
End If
End Sub

Private Sub DcbCustomer_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 9878
        FrmCustemerSearch.show vbModal

    End If
 

End Sub

Private Sub DcbFrom1_Change()
If TxtModFlg.Text <> "R" Then
If val(Me.DcboEmpName.BoundText) = 0 Then
MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
Exit Sub
End If
If Me.DcbFrom1.Text <> "" And Me.DcbTO1.Text <> "" Then
If val(Me.DcbFrom1.Text) >= val(Me.DcbTO1.Text) Then
MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
'DcbTO2.SetFocus
Exit Sub

Else
chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
fileFgtim val(Me.DcboEmpName.BoundText), 0
refiltimdetails val(Me.DcboEmpName.BoundText), 0
End If
End If
End If
End Sub

Private Sub DcboEmpName_Change()
If TxtModFlg.Text <> "R" Then
fileFgtim val(Me.DcboEmpName.BoundText), 0
refiltimdetails val(Me.DcboEmpName.BoundText), 0
End If
End Sub

'Private Sub DcboEmpName_Change()
'DcboEmpName_Click (0)

'End Sub






 




 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 8
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
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

 Sub GetEmployeeSalaryAccordingToComponentAll(Emp_id As Integer)
                                                    
  Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim Rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i As Integer
    Mofradd = ""
  
    sql = "SELECT     dbo.EmpSalaryComponent.[Value],dbo.mofrdat.mofrad_name,dbo.mofrdat.mofrad_type "
    sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
    sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
    sql = sql & " WHERE   (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
 
      Rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs.RecordCount > 0 Then
  With Me.FG
  .rows = Rs.RecordCount + 1
      For i = 1 To Rs.RecordCount
       .TextMatrix(i, .ColIndex("Serial")) = i
      .TextMatrix(i, .ColIndex("mofrdID")) = IIf(IsNull(Rs("mofrad_type").value), 0, Rs("mofrad_type").value)
       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(Rs("mofrad_name").value), "", Rs("mofrad_name").value)
 .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs("value").value), 0, Rs("value").value)

 
 Rs.MoveNext
      Next i
 End With
     End If
     
     

    Rs.Close
    
End Sub






Private Sub DcbTO1_Change()
If TxtModFlg.Text <> "R" Then
If val(Me.DcboEmpName.BoundText) = 0 Then
MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
Exit Sub
End If
If Me.DcbFrom1.Text <> "" And Me.DcbTO1.Text <> "" Then
If val(Me.DcbFrom1.Text) >= val(Me.DcbTO1.Text) Then
MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
'DcbTO2.SetFocus
Exit Sub

Else
chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
fileFgtim val(Me.DcboEmpName.BoundText), 0
refiltimdetails val(Me.DcboEmpName.BoundText), 0
End If
End If
End If
End Sub






















Private Sub DcbTO2_Change()
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
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim Rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
Dim bol As Boolean
Dim Tye As Integer
    With FG
               
    

        Select Case .ColKey(Col)
 
            Case "empname"
                                       
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
                ChekRepeat val(StrAccountCode), Row, bol
                If bol = False Then
               ' If StrAccountCode <> "" Then
                .TextMatrix(Row, .ColIndex("empid")) = val(StrAccountCode)
                StrSQL = " select Fullcode from  TblEmployee where Emp_ID=" & StrAccountCode & ""
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("code")) = Rs("Fullcode").value
                End If
                Else
                MsgBox "·«Ì„ﬂ‰ «Œ Ì«— «·„‰œÊ» Ê·«Ì„ﬂ‰ «· ﬂ—«—"
                .TextMatrix(Row, .ColIndex("empname")) = ""
                .TextMatrix(Row, .ColIndex("code")) = ""
                Exit Sub
                End If
              If TxtModFlg.Text <> "R" Then
fileFgtim val(StrAccountCode), Tye
refiltimdetails val(StrAccountCode), Tye
If Tye = 1 Then
MsgBox .TextMatrix(Row, .ColIndex("empname")) & "·«Ì„ﬂ‰ «Œ Ì«—"
.TextMatrix(Row, .ColIndex("empname")) = ""
                .TextMatrix(Row, .ColIndex("code")) = ""
              '  Exit Sub
End If
End If
          Case "code"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
               If StrAccountCode <> "" Then
                .TextMatrix(Row, .ColIndex("empid")) = StrAccountCode
                 StrSQL = " select * from  TblEmployee where Emp_ID=" & StrAccountCode & ""
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("empname")) = Rs("Emp_Name").value
                Else
                .TextMatrix(Row, .ColIndex("empname")) = Rs("Emp_Namee").value
                End If
                End If
                   End Select
   
        If Row = .rows - 1 Then
    
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
 Case "code"
FG.ComboList = ""
End Select

End With
FG.ComboList = ""
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String


    With FG

        Select Case .ColKey(Col)

            Case "empname"
       
 StrSQL = " select * from  TblEmployee "
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(Rs, "emp_name", "Emp_ID")
                Else
                    StrComboList = FG.BuildComboList(Rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "code"
              StrSQL = " select * from  TblEmployee "
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(Rs, "FullCode", "Emp_ID")
                Else
                    StrComboList = FG.BuildComboList(Rs, "FullCode", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub FG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim Rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim qunt, qunt1, qunt2, total As Long
Dim StrComboList As String
Dim bol As Boolean
    With Fg2
               
    

        Select Case .ColKey(Col)
 
            Case "name"
                                       
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
                'If StrAccountCode <> "" Then
                ChekRepeat val(StrAccountCode), Row, bol
                If bol = False Then
                .TextMatrix(Row, .ColIndex("empid")) = StrAccountCode
                StrAccountCode = val(.TextMatrix(Row, .ColIndex("empid")))
                StrSQL = " select quantity from  TblCompo where id=" & StrAccountCode & ""
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs.RecordCount > 0 Then
                qunt = IIf(IsNull(Rs("quantity").value), 0, Rs("quantity").value)
 End If
               ' .TextMatrix(Row, .ColIndex("code")) = rs("quantity").value
               Set Rs1 = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblRegDateDelgateDails.quantity AS sm, dbo.TblRegDateDelgate.Id, dbo.TblRegDateDelgate.Accept"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRegDateDelgate ON dbo.TblRegDateDelgateDails.DelgID = dbo.TblRegDateDelgate.Id"
StrSQL = StrSQL & " WHERE     (dbo.TblRegDateDelgateDails.Type = 1) AND (dbo.TblRegDateDelgateDails.EmpID = " & StrAccountCode & ") AND (dbo.TblRegDateDelgateDails.DelgID <> " & val(XPTxtID.Text) & ") AND"
StrSQL = StrSQL & "                      (dbo.TblRegDateDelgate.Accept = 0)"
Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
qunt1 = Rs1("sm").value
qunt1 = qunt - val(Rs1("sm").value)
If qunt1 > 0 Then
.TextMatrix(Row, .ColIndex("code")) = qunt1
End If
Else
.TextMatrix(Row, .ColIndex("code")) = qunt
End If
              '  End If
            Else
            MsgBox "·«Ì„ﬂ‰ «· ﬂ—«—"
            .TextMatrix(Row, .ColIndex("name")) = ""
            Exit Sub
            End If
  Case "code"
    'StrAccountCode = .ComboData
    '            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
          
               ' .TextMatrix(Row, .ColIndex("empid")) = StrAccountCode
                StrAccountCode = val(.TextMatrix(Row, .ColIndex("empid")))
                StrSQL = " select quantity from  TblCompo where id=" & StrAccountCode & ""
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Rs.RecordCount > 0 Then
                qunt = Rs("quantity").value
                Else
                qunt = 0
                End If
 qunt2 = val(.TextMatrix(Row, .ColIndex("code")))
 
 
 '''//////////////////
                 Set Rs1 = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblRegDateDelgateDails.quantity AS sm, dbo.TblRegDateDelgate.Id, dbo.TblRegDateDelgate.Accept"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRegDateDelgate ON dbo.TblRegDateDelgateDails.DelgID = dbo.TblRegDateDelgate.Id"
StrSQL = StrSQL & " WHERE     (dbo.TblRegDateDelgateDails.Type = 1) AND (dbo.TblRegDateDelgateDails.EmpID = " & StrAccountCode & ") AND (dbo.TblRegDateDelgateDails.DelgID <> " & val(XPTxtID.Text) & ") AND"
StrSQL = StrSQL & "                      (dbo.TblRegDateDelgate.Accept = 0)"

Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
qunt1 = Rs1("sm").value
Else
qunt1 = 0
End If
total = qunt - qunt1 - qunt2
'qunt1 = qunt - val(rs1("sm").value)
If total >= 0 Then
.TextMatrix(Row, .ColIndex("code")) = qunt2

Else
MsgBox "«·ﬂ„ÌÂ «ﬂ»— „‰ «·ﬂ„ÌÂ «·„ Ê›—Â "

.TextMatrix(Row, .ColIndex("code")) = ""
Exit Sub
End If
'End If

                   End Select
   
        If Row = .rows - 1 Then
    
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub FG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg2
Select Case .ColKey(Col)
 Case "code"
Fg2.ComboList = ""
End Select

End With
Fg2.ComboList = ""
End Sub

Private Sub FG2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String


    With Fg2

        Select Case .ColKey(Col)

            Case "name"
       
 StrSQL = " select * from  TblCompo "
                Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg2.BuildComboList(Rs, "name", "Id")
                Else
                    StrComboList = Fg2.BuildComboList(Rs, "namee", "Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
             Case "code"
             If .TextMatrix(Row, .ColIndex("name")) = "" Then
             MsgBox "ÌÃ» «Œ Ì«— «·«”„ «Ê·«"
            Exit Sub
             End If
             ' StrSQL = " select * from  TblEmployee "
             '   rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrComboList = Fg.BuildComboList(rs, "FullCode", "Emp_ID")
'                Else
'                    StrComboList = Fg.BuildComboList(rs, "FullCode", "Emp_ID")
'                End If
       
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If

'                .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub



 

 

Private Sub menue_Click(Index As Integer)
If Index = 2 Then
 Load FrmCustemers
            FrmCustemers.show
            End If
End Sub

 
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
    TxtNoteSerial1.Text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

Frame6.Visible = False
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
    Set Dcombos = New ClsDataCombos
    Dcombos.GetTypeVisit Me.DcbTypeVisit1
    Dcombos.GetTypeVisit Me.DcbTypeVisit2
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetDelegate Me.DcboEmpName
    Dcombos.GetBranches Me.dcBranch
Dcombos.GetTimeDelegate Me.DcbFrom1
Dcombos.GetTimeDelegate Me.DcbFrom2
Dcombos.GetTimeDelegate Me.DcbTO1
Dcombos.GetTimeDelegate Me.DcbTO2
  '  Dcombos.GetEmpLocations Me.DcProject
    Dcombos.GetSpeciaAsement Me.DcbSpecialAs
  ' Dcombos.GetEmployeesNationlity Me.DcNational
    Dcombos.GetEmpJobsTypes Me.DcbJobID1
    Dcombos.GetFileCustomer Me.DcbCustomer
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If
    DTP_Date.value = Date
Frame2.Visible = False
Frame5.Visible = False

'GrdWa.ColComboList(GrdWa.ColIndex("StatusVisit")) = "#1; ·„   „|#2;  „  |#3; «·€Ì |"
PrepareGrdWa
Me.txtnotAccept.Visible = False
lbl(36).Visible = False
    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set Rs = New ADODB.Recordset
    StrSQL = "select * From TblRegDateDelgate     Order By ID"
    Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
            If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Retrive
   
'Me.menue(2).Enabled = True
'Me.TimeFrom1.value = Time
'Me.TimeFrom2.value = Time
'Me.TimeTo1.value = Time
'Me.TimeTo2.value = Time
 

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
cmdApi_Click
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
  lbl(26).Caption = "Internal"
  lbl(35).Caption = "Bill No"
  Command2.Caption = "Visit Details"
  lbl(36).Caption = "Reason cancellation visit"
  Frame5.Caption = "Data of  Contracted"
  lbl(34).Caption = "Requirements"
  lbl(33).Caption = "Team Required"
  Command1.Caption = "Show"
  lbl(29).Caption = "Address"
  lbl(31).Caption = "Maps"
  Frame4.Caption = "Data of Visit"
  lbl(23).Caption = "Date Visit"
ChekAccept.Caption = "Visited"
ChekAccept.RightToLeft = False
    Cmd(21).Caption = "Remove"
    Cmd(8).Caption = "Remove"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
XPTab301.Caption = "Data"
lbl(37).Caption = "Data Visits the Client's"
    Me.Caption = "Screen Recording Dates Delegates"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
   lbl(32).Caption = "Time"
   lbl(11).Caption = "From"
   lbl(12).Caption = "To"
   
    lbl(2).Caption = "Type Visit"
    lbl(3).Caption = "Delegate"
    Frame3.Caption = "Data of Customer"
     lbl(0).Caption = "Customer"
    lbl(9).Caption = "Admin"
 'lbDW.Caption = "Data of Employee"
   lbl(21).Caption = "Job"
  lbl(20).Caption = "Email"
    lbl(18).Caption = "Phone"
   lbl(10).Caption = "Remarks"
    lbl(13).Caption = "Remarks"
  lbl(19).Caption = "Mobile"
  lbl(24).Caption = "Valuation"
  Frame2.Caption = "After the Visit"
  lbl(5).Caption = "Next Step"
  lbl(15).Caption = "Date"
 lbl(14).Caption = "Time"
   lbl(16).Caption = "From"
   lbl(17).Caption = "To"
  
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
   Me.CHekNotAccept.Caption = "Cancel Visit"
   Me.CHekNotAccept.RightToLeft = False
   Me.ChekContracted.RightToLeft = False
   Me.ChekContracted.Caption = "Been contracted"
    lbl(8).Caption = "By"
    With Me.FG
       .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("empname")) = "Employee"
.TextMatrix(0, .ColIndex("remarks")) = "Remarks"
    End With

    With Me.Fg2
       .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("code")) = "Quantity"
        .TextMatrix(0, .ColIndex("name")) = "name"
.TextMatrix(0, .ColIndex("remarks")) = "Remarks"
    End With
  With Me.VSFlexGrid1
       .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("Emp_NameD")) = "Delegate"
        .TextMatrix(0, .ColIndex("CustomerName")) = "CustomerName"
.TextMatrix(0, .ColIndex("PersonConc")) = "PersonConc"
       .TextMatrix(0, .ColIndex("JobID")) = "JobName"
        .TextMatrix(0, .ColIndex("Tel")) = "Phone"
        .TextMatrix(0, .ColIndex("Mobile")) = "Mobile"
.TextMatrix(0, .ColIndex("Email")) = "Email"
   .TextMatrix(0, .ColIndex("Adress")) = "Address"
        .TextMatrix(0, .ColIndex("VisitDate1")) = "Date"
        .TextMatrix(0, .ColIndex("VisitID")) = "TypeVisit"
.TextMatrix(0, .ColIndex("FromTim11")) = "From Hour"
.TextMatrix(0, .ColIndex("ToTime11")) = "To Hour"
    .TextMatrix(0, .ColIndex("Accept")) = "Status"
        .TextMatrix(0, .ColIndex("Remark")) = "Remark"
       

    End With
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

    If Rs.State = adStateOpen Then
        If Not (Rs.EOF Or Rs.BOF) Then
            If Rs.EditMode <> adEditNone Then
                Rs.CancelUpdate
            End If
        End If

        Rs.Close
        Set Rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtAdvanceValue_LostFocus()
 
   
End Sub

Public Sub retInfoCustomer()
 Dim EmpID As Integer
Dim Name As String
Dim Mobile As String
Dim phone As String
Dim boxmail As String
Dim fax As String
Dim mail As String
Dim adress As String
Dim ZipCode As String
Dim DigCus As String
    Dim fullcode As String
    Dim map As String
Dim entry As String
Dim ResponsibleContact As String
    Dim jobname As String
        GetCustomerIDFromCode Me.TxtCustomer.Text, EmpID, , fullcode, Me.DcbCustomer.Text, Name, Mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus, jobname, entry, map, ResponsibleContact
         Me.TxtCustomer = fullcode
         Me.TxtPersonCont.Text = ResponsibleContact
        Me.DcbCustomer.BoundText = EmpID
        Me.TxtMobi.Text = Mobile
        Me.txtTel.Text = phone
       Me.TxtMap.Text = map
        Me.TxtEnter.Text = entry
        Me.DcbJobID.Text = jobname
        Me.TxtEmail.Text = mail
        Me.TxtAdres.Text = adress
        'Me.txtboxzip.text = ZipCode
        
        'Me.TxtTypeCustomer.text = val(DigCus) + 1
       ' DcboEmpName.BoundText = EmpID
    
End Sub

Private Sub optInvType_Click(Index As Integer)
    On Error GoTo ErrH
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
    ' ›—Ì€ ﬁ»· ≈⁄«œ… «· ⁄»∆…
    DcbCustomer.Text = ""
    DcbCustomer.BoundText = ""

    Select Case Index
        Case 0  '⁄„Ì·
            lbl(0).Caption = "⁄„Ì·"
            ' ⁄»∆… «·⁄„·«¡
            Dcombos.GetFileCustomer Me.DcbCustomer

        Case 1  '„‘—Ê⁄
            lbl(0).Caption = "„‘—Ê⁄"
            ' ⁄»∆… «·„‘«—Ì⁄
            Dim My_SQL As String
            My_SQL = "select id, project_name from Projects where isnull(project_name,'') <> ''"
            
            fill_combo Me.DcbCustomer, My_SQL
    End Select

    Exit Sub

ErrH:
    MsgBox Err.Description, vbExclamation, "optInvType_Click"
End Sub


Private Sub TxtCustomer_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 retInfoCustomer
End If
End Sub

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

            If Rs.RecordCount < 1 Then
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
            XPDtbTrans.value = Date

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

            If Not (Rs.EOF Or Rs.BOF) Then
                Rs.MovePrevious

                If Rs.BOF Then Rs.MoveFirst
            End If

        Case 1

            If Not (Rs.EOF Or Rs.BOF) Then
                Rs.MoveFirst
            End If

        Case 2

            If Not (Rs.EOF Or Rs.BOF) Then
                Rs.MoveLast
            End If

        Case 3

            If Not (Rs.EOF Or Rs.BOF) Then
                Rs.MoveNext

                If Rs.EOF Then Rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub
Sub chektime(Optional EmpID As Integer, Optional timfrom As Integer, Optional timto As Integer, Optional TypeViste As Integer, Optional ByRef Find As Boolean)
Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
    
StrSQL = " SELECT     dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.VisitID, dbo.TblTypeVisit.name, dbo.TblTypeVisit.namee, TblRegTimeDelgate_1.name AS namef1, "
 StrSQL = StrSQL & "                     dbo.TblRegTimeDelgate.name AS Namet1, dbo.TblRegDateDelgate.DelgID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblRegDateDelgate.Id, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.ToTime1,"
StrSQL = StrSQL & "                      dbo.TblRegDateDelgate.FromTime1 , dbo.TblRegDateDelgate.ToTime2, dbo.TblRegDateDelgate.FromTime2, dbo.TblRegDateDelgate.VisitDate"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgate LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblRegTimeDelgate ON dbo.TblRegDateDelgate.ToTime1 = dbo.TblRegTimeDelgate.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgateDails ON dbo.TblRegDateDelgate.Id = dbo.TblRegDateDelgateDails.DelgID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblRegDateDelgate.DelgID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblRegTimeDelgate TblRegTimeDelgate_1 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTypeVisit ON dbo.TblRegDateDelgate.VisitID = dbo.TblTypeVisit.Id"
StrSQL = StrSQL & " WHERE    ( (dbo.TblRegDateDelgate.DelgID = " & EmpID & ") or   (dbo.TblRegDateDelgateDails.EmpID = " & EmpID & ")) AND"
If TypeViste = 1 Then
StrSQL = StrSQL & "   (dbo.TblRegDateDelgate.ToTime2 <= " & timto & ") AND (dbo.TblRegDateDelgate.FromTime2 >=  " & timfrom & ") "
 StrSQL = StrSQL & "                  and   (dbo.TblRegDateDelgate.VisitDate =' " & SQLDate(Me.XpDtbVisit.value) & " ') "
Else
StrSQL = StrSQL & "   (dbo.TblRegDateDelgate.ToTime1 <= " & timto & ") AND (dbo.TblRegDateDelgate.FromTime1 >=  " & timfrom & ") "
 StrSQL = StrSQL & "                  and   (dbo.TblRegDateDelgate.VisitDate1 =' " & SQLDate(Me.DateVisit1.value) & " ') "
End If
'StrSQL = StrSQL & " WHERE     (dbo.TblRegDateDelgate.DelgID =" & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblRegDateDelgate.VisitDate1 =' " & SQLDate(Me.DateVisit1.value) & " ')"
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
Find = False
If RsDetails.RecordCount > 0 Then
Find = True
MsgBox "·«Ì„ﬂ‰ «·ÕÃ“ ›Ì Â–« «·„Ê⁄œ ·«‰Â „ÕÃÊ“ „”»ﬁ«"
End If
End Sub
Sub refiltimdetails(Optional ByRef EmpID As Integer, Optional ByRef Tye As Integer = 0)
Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
    
StrSQL = " SELECT     dbo.TblRegDateDelgate.VisitID, dbo.TblTypeVisit.name, dbo.TblTypeVisit.namee, TblRegTimeDelgate_1.name AS namef1, dbo.TblRegTimeDelgate.name AS Namet1, "
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgate.DelgID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, dbo.TblRegDateDelgate.Id,"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.ToTime1, dbo.TblRegDateDelgate.FromTime1,"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgateDails.Type,"
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map,"
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, dbo.TblRegDateDelgate.CustomerID, dbo.TblCustemers.CusName,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.CusNamee , dbo.TblRegDateDelgate.Remark"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgate LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblRegTimeDelgate ON dbo.TblRegDateDelgate.ToTime1 = dbo.TblRegTimeDelgate.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgateDails ON dbo.TblRegDateDelgate.Id = dbo.TblRegDateDelgateDails.DelgID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblRegDateDelgate.DelgID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_1 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_1.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblTypeVisit ON dbo.TblRegDateDelgate.VisitID = dbo.TblTypeVisit.Id"
StrSQL = StrSQL & " WHERE     (dbo.TblRegDateDelgateDails.EmpID =" & val(EmpID) & ") AND (dbo.TblRegDateDelgate.VisitDate1 =' " & SQLDate(Me.DateVisit1.value) & " ' ) and  (dbo.TblRegDateDelgateDails.Type = 0) and (dbo.TblRegDateDelgate.Accept = 0) "

RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

If RsDetails.RecordCount > 0 Then
Tye = 1
 FrmViewDelegate.Fg2.Clear flexClearScrollable, flexClearEverything
 With FrmViewDelegate.Fg2
    .rows = .FixedRows
FrmViewDelegate.DtpDateFrom.value = Me.DateVisit1.value
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
        .TextMatrix(i, .ColIndex("Ser")) = i
            If SystemOptions.UserInterface = EnglishInterface Then
                      .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
           Else
          
            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
           End If
        .TextMatrix(i, .ColIndex("fromtime")) = val(IIf(IsNull(RsDetails("namef1").value), "", RsDetails("namef1").value)) ' RsDetails("namef1").value
            .TextMatrix(i, .ColIndex("totime")) = val(IIf(IsNull(RsDetails("Namet1").value), "", RsDetails("Namet1").value)) 'RsDetails("Namet1").value
         .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("Emp_Name").value), "", RsDetails("Emp_Name").value)
        '  .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CustomerName").value), "", RsDetails("CustomerName").value)
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
End With
 Load FrmViewDelegate
 FrmViewDelegate.show
End If
End Sub
Sub fileFgtim(Optional ByRef EmpID As Integer, Optional ByRef Tye As Integer = 0)
Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
    
StrSQL = "SELECT     dbo.TblRegDateDelgate.VisitID, dbo.TblTypeVisit.name, dbo.TblTypeVisit.namee, TblRegTimeDelgate_1.name AS namef1, dbo.TblRegTimeDelgate.name AS Namet1, "
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgate.Id, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.ToTime1, dbo.TblRegDateDelgate.FromTime1,"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.Entry,"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgate.Map, dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, dbo.TblRegDateDelgate.JobID,"
 StrSQL = StrSQL & "                     dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.CustomerID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusNamee , dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.DelgID"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgate LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblRegTimeDelgate ON dbo.TblRegDateDelgate.ToTime1 = dbo.TblRegTimeDelgate.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_1 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTypeVisit ON dbo.TblRegDateDelgate.VisitID = dbo.TblTypeVisit.Id"

StrSQL = StrSQL & " WHERE   (dbo.TblRegDateDelgate.Accept = 0) And  (dbo.TblRegDateDelgate.DelgID =" & val(Me.DcboEmpName.BoundText) & ") AND (dbo.TblRegDateDelgate.VisitDate1 =' " & SQLDate(Me.DateVisit1.value) & " ')"
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

If RsDetails.RecordCount > 0 Then
Tye = 1
 FrmViewDelegate.FG.Clear flexClearScrollable, flexClearEverything
 With FrmViewDelegate.FG
    .rows = .FixedRows
FrmViewDelegate.DtpDateFrom.value = Me.DateVisit1.value
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .rows - 1
        
        .TextMatrix(i, .ColIndex("Ser")) = i
        .RightToLeft = True
        If SystemOptions.UserInterface = EnglishInterface Then
                      .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
           Else
          
            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
           End If
        .TextMatrix(i, .ColIndex("fromtime")) = val(IIf(IsNull(RsDetails("namef1").value), "", RsDetails("namef1").value)) ' RsDetails("namef1").value
            .TextMatrix(i, .ColIndex("totime")) = val(IIf(IsNull(RsDetails("Namet1").value), "", RsDetails("Namet1").value)) 'RsDetails("Namet1").value
        '  .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CustomerName").value), "", RsDetails("CustomerName").value)
           .TextMatrix(i, .ColIndex("PersonConc")) = IIf(IsNull(RsDetails("PersonConc").value), "", RsDetails("PersonConc").value)
            .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(RsDetails("Tel").value), "", RsDetails("Tel").value)
             .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(RsDetails("Mobile").value), "", RsDetails("Mobile").value)
              .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(RsDetails("Email").value), "", RsDetails("Email").value)
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
End With
 Load FrmViewDelegate
 FrmViewDelegate.show
End If
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
     Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
          FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
    'On Error GoTo ErrTrap
    If Rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Rs.EOF Or Rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            Rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If Rs.EOF Or Rs.BOF Then
                Exit Sub
            End If
        End If
    End If


    XPTxtID.Text = IIf(IsNull(Rs("Id").value), "", val(Rs("Id").value))
    Me.TxtAdres.Text = IIf(IsNull(Rs("Adress").value), "", Rs("Adress").value)
    Me.TxtBillNo.Text = IIf(IsNull(Rs("BillNo").value), "", Rs("BillNo").value)
    Me.txtnotAccept.Text = IIf(IsNull(Rs("NotAcept").value), "", Rs("NotAcept").value)
    
    
Me.TxtEnter.Text = IIf(IsNull(Rs("Entry").value), "", Rs("Entry").value)
Me.TxtMap.Text = IIf(IsNull(Rs("Map").value), "", Rs("Map").value)
    XPDtbTrans.value = IIf(IsNull(Rs("RecordDate").value), Date, Rs("RecordDate").value)
      DateVisit1.value = IIf(IsNull(Rs("VisitDate1").value), Date, Rs("VisitDate1").value)
 XpDtbVisit.value = IIf(IsNull(Rs("VisitDate").value), Date, Rs("VisitDate").value)
DcbJobID.Text = IIf(IsNull(Rs("JobID").value), "", Rs("JobID").value)
    dcBranch.BoundText = IIf(IsNull(Rs("BranchID").value), "", Rs("BranchID").value)
     Me.DcboEmpName.BoundText = IIf(IsNull(Rs("DelgID").value), "", Rs("DelgID").value)
    Me.DcbTypeVisit1.BoundText = IIf(IsNull(Rs("VisitID").value), "", Rs("VisitID").value)
    Me.DcbTypeVisit2.BoundText = IIf(IsNull(Rs("VisitID2").value), "", Rs("VisitID2").value)
    DCboUserName.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
    
    
      Me.TxtCustomer.Text = IIf(IsNull(Rs("CustomerName").value), "", Rs("CustomerName").value)
    Me.TxtRemark1.Text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
    Me.TxtRemark2.Text = IIf(IsNull(Rs("Remark2").value), "", Rs("Remark2").value)
    Me.TxtLongTime.Text = IIf(IsNull(Rs("LongTime").value), "", Rs("LongTime").value)
   If Rs("Accept").value = 0 Then
         Me.ChekContracted.value = xtpUnchecked
         Me.ChekAccept.value = xtpUnchecked
         Me.CHekNotAccept.value = xtpUnchecked
End If
   If Rs("Accept").value = 1 Then
         Me.ChekAccept.value = xtpChecked
         Me.CHekNotAccept.value = xtpUnchecked
         Me.ChekContracted.value = xtpUnchecked
End If

 If val(Rs("InvType").value & "") = 0 Then
 
        optInvType(0).value = True
        optInvType_Click 0
        retInfoCustomer
 ElseIf val(Rs("InvType").value & "") = 1 Then
    optInvType(1).value = True
    optInvType_Click 1
 End If
 
Me.DcbCustomer.BoundText = IIf(IsNull(Rs("customerid").value), "", Rs("customerid").value)
   If Rs("Accept").value = 2 Then
         Me.ChekContracted.value = xtpChecked
         Me.CHekNotAccept.value = xtpUnchecked
End If
   If Rs("Accept").value = 3 Then
         Me.ChekContracted.value = xtpUnchecked
         Me.ChekAccept.value = xtpUnchecked
         Me.CHekNotAccept.value = xtpChecked
End If
   Me.TxtPersonCont.Text = IIf(IsNull(Rs("PersonConc").value), "", Rs("PersonConc").value)
Me.txtTel.Text = IIf(IsNull(Rs("Tel").value), "", Rs("Tel").value)
   Me.TxtMobi.Text = IIf(IsNull(Rs("Mobile").value), "", Rs("Mobile").value)
Me.TxtEmail.Text = IIf(IsNull(Rs("Email").value), "", Rs("Email").value)

Me.DcbFrom1.BoundText = IIf(IsNull(Rs("FromTime1").value), "", Rs("FromTime1").value)
    Me.DcbFrom2.BoundText = IIf(IsNull(Rs("FromTime2").value), "", Rs("FromTime2").value)
    DcbTO1.BoundText = IIf(IsNull(Rs("ToTime1").value), "", Rs("ToTime1").value)
    Me.DcbTO2.BoundText = IIf(IsNull(Rs("ToTime2").value), "", Rs("ToTime2").value)


    txtDateVis(0).value = IIf(IsNull(Rs("DateVis0").value), Date, Rs("DateVis0").value)
    txtDateVis(1).value = IIf(IsNull(Rs("DateVis1").value), Date, Rs("DateVis1").value)
    
    txtVisitTime(0).value = IIf(IsNull(Rs("VisitTime0").value), Time, Rs("VisitTime0").value)
    txtVisitTime(1).value = IIf(IsNull(Rs("VisitTime1").value), Time, Rs("VisitTime1").value)
    Me.txtGPS(0).Text = IIf(IsNull(Rs("GPS0").value), "", Rs("GPS0").value)
    Me.txtGPS(1).Text = IIf(IsNull(Rs("GPS1").value), "", Rs("GPS1").value)

    Me.TxtAddress(0).Text = IIf(IsNull(Rs("Address0").value), "", Rs("Address0").value)
    Me.TxtAddress(1).Text = IIf(IsNull(Rs("Address1").value), "", Rs("Address1").value)
        


'Me.TimeFrom1.value = Format(CDate(rs("TimeFrom1").value), "hh:mm AM/PM")  'IIf(IsNull(rs("TimeFrom1").value), Time, rs("TimeFrom1").value)
' Me.TimeTo1.value = Format(CDate(rs("TimeTo1").value), "hh:mm AM/PM") 'IIf(IsNull(rs("TimeTo1").value), Time, rs("TimeTo1").value)
' Me.TimeFrom2.value = Format(CDate(rs("TimeFrom2").value), "hh:mm AM/PM") 'IIf(IsNull(rs("TimeFrom2").value), Time, rs("TimeFrom2").value)
' Me.TimeTo2.value = Format(CDate(rs("TimeTo2").value), "hh:mm AM/PM") 'IIf(IsNull(rs("TimeTo2").value), Time, rs("TimeTo2").value)
  
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
    Set RsDetails = New ADODB.Recordset
 StrSQL = " SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblEmployee.Emp_Code, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, dbo.TblRegDateDelgateDails.remark,"
StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 0) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.Text) & " )"
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"


 RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = FG.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        FG.rows = FG.FixedRows + RsDetails.RecordCount

        For i = Me.FG.FixedRows To FG.rows - 1
        FG.TextMatrix(i, FG.ColIndex("Serial")) = i
        FG.TextMatrix(i, FG.ColIndex("remarks")) = IIf(IsNull(RsDetails("remark").value), "", RsDetails("remark").value) ' RsDetails("remark").value
            FG.TextMatrix(i, FG.ColIndex("code")) = IIf(IsNull(RsDetails("fullcode").value), "", RsDetails("fullcode").value) 'RsDetails("fullcode").value
            If SystemOptions.UserInterface = EnglishInterface Then
           FG.TextMatrix(i, FG.ColIndex("empname")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
           Else
           FG.TextMatrix(i, FG.ColIndex("empname")) = IIf(IsNull(RsDetails("emp_name").value), "", RsDetails("emp_name").value) ' RsDetails("emp_name").value
           End If
            FG.TextMatrix(i, FG.ColIndex("empid")) = RsDetails("EmpID").value
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    
'    StrSQL = "Select * from TblRegDateDelgateDailsGrantee Where DelgID = " & val(Me.XPTxtID.Text)
'    loadgrid StrSQL, GrdWa, True, False

LoadSavedGridWa
    
    Set RsDetails = Nothing
   '''''''''''''///////////////////////
   Set RsDetails1 = New ADODB.Recordset
 StrSQL = "SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgateDails.remark, "
StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type , dbo.TblCompo.name, dbo.TblCompo.namee, dbo.TblRegDateDelgateDails.Quantity"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id"

StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 1) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.Text) & " )"



 RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    Fg2.Clear flexClearScrollable, flexClearEverything
    Fg2.rows = Fg2.FixedRows

    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
        RsDetails1.MoveFirst
        Fg2.rows = Fg2.FixedRows + RsDetails1.RecordCount

        For i = Me.Fg2.FixedRows To Fg2.rows - 1
        Fg2.TextMatrix(i, Fg2.ColIndex("Serial")) = i
        Fg2.TextMatrix(i, Fg2.ColIndex("remarks")) = IIf(IsNull(RsDetails1("remark").value), "", RsDetails1("remark").value) ' RsDetails1("remark").value
            Fg2.TextMatrix(i, Fg2.ColIndex("code")) = IIf(IsNull(RsDetails1("quantity").value), "", RsDetails1("quantity").value) 'RsDetails1("fullcode").value
            If SystemOptions.UserInterface = EnglishInterface Then
           Fg2.TextMatrix(i, Fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value) 'RsDetails1("Emp_Namee").value
           Else
           Fg2.TextMatrix(i, Fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value) ' RsDetails1("emp_name").value
           End If
            Fg2.TextMatrix(i, Fg2.ColIndex("empid")) = RsDetails1("EmpID").value
            RsDetails1.MoveNext
        Next i

    End If

    RsDetails1.Close
    Set RsDetails1 = Nothing
   
   
   
   
   
 '  fillapprovData
    
    XPTxtCurrent.Caption = Rs.AbsolutePosition
    XPTxtCount.Caption = Rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


Private Function GetStatusVisitFromDone(ByVal V As Variant) As Integer
    If IsNull(V) Or Trim(V & "") = "" Then
        GetStatusVisitFromDone = 1
    ElseIf val(V) = 0 Then
        GetStatusVisitFromDone = 1   '·„   „
    ElseIf val(V) = 1 Then
        GetStatusVisitFromDone = 2   ' „ 
    ElseIf val(V) = 2 Then
        GetStatusVisitFromDone = 3   '√·€Ì 
    Else
        GetStatusVisitFromDone = 1
    End If
End Function
Private Function GetStatusVisitText(ByVal V As Variant) As String
    Select Case val(V)
        Case 1
            GetStatusVisitText = "·„   „"
        Case 2
            GetStatusVisitText = " „ "
        Case 3
            GetStatusVisitText = "√·€Ì "
        Case Else
            GetStatusVisitText = ""
    End Select
End Function
Private Function GetStatusVisitValue(ByVal V As Variant) As Integer
    Dim S As String
    S = Trim(V & "")

    If S = "" Then
        GetStatusVisitValue = 0
    ElseIf val(S) > 0 Then
        GetStatusVisitValue = val(S)
    Else
        Select Case S
            Case "·„   „"
                GetStatusVisitValue = 1
            Case " „ "
                GetStatusVisitValue = 2
            Case "√·€Ì ", "«·€Ì "
                GetStatusVisitValue = 3
            Case Else
                GetStatusVisitValue = 0
        End Select
    End If
End Function

'-------

Private Function SqlDateOrNull(ByVal V As Variant) As String
    If IsDate(V) Then
        SqlDateOrNull = "'" & Format$(CDate(V), "yyyy-mm-dd hh:nn:ss") & "'"
    Else
        SqlDateOrNull = "NULL"
    End If
End Function
Private Function SqlText(ByVal V As Variant) As String
    SqlText = "'" & Replace(Trim(V & ""), "'", "''") & "'"
End Function
Private Function SqlLong(ByVal V As Variant) As String
    If Trim(V & "") = "" Then
        SqlLong = "NULL"
    Else
        SqlLong = CStr(val(V))
    End If
End Function
Public Sub PrepareGrdWa()

    On Error Resume Next

    GrdWa.ColComboList(GrdWa.ColIndex("StatusVisit")) = "#1; ·„   „|#2;  „ |#3; √·€Ì |"

    On Error GoTo 0

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
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„‰œÊ»..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If

Dim Find As Boolean
  If TxtModFlg.Text = "N" Then
            chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2, Find
            If Find = True Then
            MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ  »”»»  ⁄«—÷ «·Êﬁ "
            Exit Sub
            End If
End If
       If Me.TxtModFlg.Text <> "R" Then
        If TxtCustomer.Text = "" Then
'            Msg = "ÌÃ»  ÕœÌœ «”„ «·⁄Ì„·..!! "
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            TxtCustomer.SetFocus
        '    SendKeys "{F4}"
     '       Exit Sub
        End If
     Dim RsTest As New ADODB.Recordset

        Cn.BeginTrans
        BeginTrans = True
        
              If TxtModFlg.Text = "N" Then


        '”·› ”«»ﬁ…
   


            XPTxtID.Text = CStr(new_id("TblRegDateDelgate", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            Rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblRegDateDelgateDails Where DelgID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
           Rs("ID").value = val(XPTxtID.Text)
           Rs("Adress").value = IIf(Me.TxtAdres.Text = "", "", Me.TxtAdres.Text)
           Rs("Map").value = IIf(Me.TxtMap.Text = "", "", Me.TxtMap.Text)
           Rs("Entry").value = IIf(Me.TxtEnter.Text = "", "", Me.TxtEnter.Text)
            Rs("NotAcept").value = IIf(Me.txtnotAccept.Text = "", "", Me.txtnotAccept.Text)
             Rs("BillNo").value = IIf(Me.TxtBillNo.Text = "", "", Me.TxtBillNo.Text)
         Rs("RecordDate").value = XPDtbTrans.value
        Rs("VisitDate").value = XpDtbVisit.value
        
        Rs("DateVis0").value = txtDateVis(0).value
        Rs("DateVis1").value = txtDateVis(1).value
        
        Rs("VisitTime0").value = txtVisitTime(0).value
        Rs("VisitTime1").value = txtVisitTime(1).value
        Rs("GPS0").value = IIf(Me.txtGPS(0).Text = "", "", Me.txtGPS(0).Text)
        Rs("GPS1").value = IIf(Me.txtGPS(1).Text = "", "", Me.txtGPS(1).Text)
        
        Rs("Address1").value = IIf(Me.TxtAddress(1).Text = "", "", Me.TxtAddress(1).Text)
        Rs("Address0").value = IIf(Me.TxtAddress(0).Text = "", "", Me.TxtAddress(0).Text)
        
        
        If optInvType(0).value Then
            Rs("InvType").value = 0
      ElseIf optInvType(1).value Then
             Rs("InvType").value = 1
        End If

        
        
        Rs("VisitDate1").value = Me.DateVisit1.value
         Rs("Remark").value = IIf(Me.TxtRemark1.Text = "", "", Me.TxtRemark1.Text)
         Rs("Remark2").value = IIf(TxtRemark2.Text = "", "", Me.TxtRemark2.Text)
         Rs("CustomerName").value = IIf(TxtCustomer.Text = "", "", Me.TxtCustomer.Text)
        Rs("DelgID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
       Rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
         Rs("CustomerID").value = IIf(Me.DcbCustomer.BoundText = "", Null, Me.DcbCustomer.BoundText)
       Rs("JobID").value = IIf(Me.DcbJobID.Text = "", "", (Me.DcbJobID.Text))
         Rs("VisitID").value = val(IIf(Me.DcbTypeVisit1.BoundText = "", 0, Me.DcbTypeVisit1.BoundText))
         Rs("VisitID2").value = val(IIf(Me.DcbTypeVisit2.BoundText = "", 0, Me.DcbTypeVisit2.BoundText))
  Rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText) 'Me.DCboUserName.BoundText
 Rs("SpAsID").value = IIf(Me.DcbSpecialAs.BoundText = "", Null, Me.DcbSpecialAs.BoundText)
  Rs("PersonConc").value = IIf(Me.TxtPersonCont.Text = "", "", Me.TxtPersonCont.Text)
    Rs("Tel").value = IIf(Me.txtTel.Text = "", "", Me.txtTel.Text)

     Rs("LongTime").value = IIf(Me.TxtLongTime.Text = "", "", Me.TxtLongTime.Text)
    
      Rs("Mobile").value = IIf(Me.TxtMobi.Text = "", "", Me.TxtMobi.Text)
        Rs("Email").value = IIf(Me.TxtEmail.Text = "", "", Me.TxtEmail.Text)
        
        
        Rs("FromTime1").value = IIf(Me.DcbFrom1.BoundText = "", Null, Me.DcbFrom1.BoundText)
         Rs("FromTime2").value = IIf(Me.DcbFrom2.BoundText = "", Null, Me.DcbFrom2.BoundText)
          Rs("ToTime1").value = IIf(Me.DcbTO1.BoundText = "", Null, Me.DcbTO1.BoundText)
           Rs("ToTime2").value = IIf(Me.DcbTO2.BoundText = "", Null, Me.DcbTO2.BoundText)
        
        
     ' rs("TimeFrom1").value = Format(TimeFrom1.value, "hh:mm AM/PM")                          '    Me.TimeFrom1.value ' Format(Me.TimeFrom1.value, TimeFrom1.CustomFormat)
     ' rs("TimeTo1").value = Format(TimeTo1.value, "hh:mm AM/PM")     'Me.TimeTo1.value
     ' rs("TimeFrom2").value = Format(TimeFrom2.value, "hh:mm AM/PM")  'Me.TimeFrom2.value
     ' rs("TimeTo2").value = Format(TimeTo2.value, "hh:mm AM/PM") 'Me.TimeFrom2.value
      Rs("Accept").value = 0
   If Me.ChekAccept.value = vbChecked Then
    Rs("Accept").value = 1
      End If
If Me.ChekContracted.value = vbChecked Then
    Rs("Accept").value = 2
      End If
If Me.CHekNotAccept.value = vbChecked Then
    Rs("Accept").value = 3
      End If

        Rs.update
           Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
        For i = Me.FG.FixedRows To FG.rows - 1
       If val(FG.TextMatrix(i, FG.ColIndex("EmpID"))) <> 0 Then
            RsDetails.AddNew
            RsDetails("DelgID").value = val(XPTxtID.Text)
            RsDetails("Type").value = 0
           RsDetails("remark").value = FG.TextMatrix(i, FG.ColIndex("remarks"))
            RsDetails("EmpID").value = val(FG.TextMatrix(i, FG.ColIndex("empid")))
   
            RsDetails.update
        End If
        Next i
  ''///////////'''''''''''''''''''''''''''''''
        Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
        For i = Me.Fg2.FixedRows To Fg2.rows - 1
       If val(Fg2.TextMatrix(i, Fg2.ColIndex("EmpID"))) <> 0 Then
            RsDetails1.AddNew
            RsDetails1("DelgID").value = val(XPTxtID.Text)
            RsDetails1("Type").value = 1
           RsDetails1("remark").value = Fg2.TextMatrix(i, Fg2.ColIndex("remarks"))
            RsDetails1("EmpID").value = val(Fg2.TextMatrix(i, Fg2.ColIndex("empid")))
    RsDetails1("quantity").value = val(Fg2.TextMatrix(i, Fg2.ColIndex("code")))
            RsDetails1.update
        End If
        Next i
        
        
'               StrSQL = "delete from TblRegDateDelgateDailsGrantee  where DelgID = " & val(XPTxtID.Text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        Dim S As String
'        S = "Select * from TblRegDateDelgateDailsGrantee Where DelgID = " & val(Me.XPTxtID.Text)
'
'        saveGrid S, GrdWa, "MaDate", "", "DelgID", val(Me.XPTxtID.Text)
        
   SaveGridWa
        
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
        BeginTrans = False
        RsDetails.Close
        
        Set RsDetails = Nothing
        RsDetails1.Close
        Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = Rs.AbsolutePosition
        XPTxtCount.Caption = Rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.Text = "R"
    End If
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

Private Function GetFormTransactionID() As Variant
    On Error Resume Next
    GetFormTransactionID = val(Me.XPTxtID.Text)
End Function

Private Function GetFormTransactionType() As Variant
    On Error Resume Next
    GetFormTransactionType = 0
End Function

Private Function GetCurrentWarntID() As Variant
    On Error Resume Next
    GetCurrentWarntID = Null
End Function
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            Rs.Find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If Rs.EOF Or Rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
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

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not Rs.RecordCount < 1 Then
                Rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Rs.MoveFirst
                StrSQL1 = "Delete From TblDefinDetails Where IDDef=" & val(Me.XPTxtID.Text)
 Cn.Execute StrSQL1, , adExecuteNoRecords
                If Rs.RecordCount < 1 Then
                    clear_all Me
                     FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
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
    Rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.Text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.Text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close

End Function
Private Sub ChekRepeat(Optional Ind As Integer, Optional Row As Long, Optional ByRef bo As Boolean)
    Dim i As Integer


    With Fg2
 bo = False
        For i = .FixedRows To .rows - 1
If i <> Row Then
            If val(.TextMatrix(i, .ColIndex("empid"))) = val(Ind) Then
             bo = True
   End If
            End If
            Next i
            End With
        With FG
 bo = False
        For i = .FixedRows To .rows - 1
If i <> Row Then
            If val(.TextMatrix(i, .ColIndex("empid"))) = val(Ind) Then
             bo = True
             End If
             Else
             
            If val(Ind) = val(Me.DcboEmpName.BoundText) Then
              bo = True
              End If
   End If
            
            Next i
            End With
        End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
       
    IntCounter = 0

    With FG

        For i = .FixedRows To .rows - 1

            If val(.TextMatrix(i, .ColIndex("empid"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Serial")) = IntCounter
   
            End If

        Next i
 
    End With
    IntCounter = 0
      With Fg2

        For i = .FixedRows To .rows - 1

            If val(.TextMatrix(i, .ColIndex("empid"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Serial")) = IntCounter
   
            End If

        Next i
 
    End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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
        .Create Me.hWnd, "     ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ”ÃÌ· „Ê«⁄Ìœ «·„‰«œÌ»  ", 1, 15204351, -2147483630
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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

Private Sub XPTab301_Click()

End Sub




Function print_report66(Optional Ind As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim mCusType As String
    Dim StrFileName As String
    Dim Msg As String
       Dim mSql1 As String
    Dim mSql2 As String

    Dim X As Integer
    MySQL = "Select * from TblAging  "
    
  
   
'    RsData.Open MySQL, Cn, adOpenKeyset, adLockReadOnly
'    If Not RsData.EOF Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'                        X = MsgBox("ÌÊÃœ ⁄„— œÌ‰  „ ⁄„·Â „”»ﬁ« » «—ÌŒ " & DTP_Date.value & "" & " Â·  Êœ ⁄—÷Â ‰⁄„/·«", vbInformation + vbYesNo)
'                    Else
'                        X = MsgBox("No Contract For This Employee Create Contarct y / n", vbInformation + vbYesNo)
'                    End If
'
'                    If X = vbYes Then
'                        loadgrid MySQL, grdAging, True, False
'                        PrintAging Ind
'                        Exit Function
'                    End If
'    End If
'
'    RsData.Close
'
    

    
    grdAging2.rows = 1
    grdAging.rows = 1
    
Dim mWhereCus As String


'
'
'   MySQL = ""
'  MySQL = MySQL & " update"
' MySQL = MySQL & " Accounts"
' MySQL = MySQL & " SET BalanceAging ="
'
'MySQL = MySQL & " (SELECT SUM(XB.TransNet)"
'
'MySQL = MySQL & " FROM   ("
'MySQL = MySQL & "            SELECT dev.Account_Code,"
'MySQL = MySQL & "                   dev.Credit_Or_Debit,"
'MySQL = MySQL & "                   dev.branch_id,"
'MySQL = MySQL & "                   dev.Notes_ID,"
'MySQL = MySQL & "                   NotesTypeName,"
'MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
'MySQL = MySQL & "                   Notes.NoteDate,"
'MySQL = MySQL & "                   Notes.NoteType,"
'MySQL = MySQL & "                   Notes.NoteSerial,"
'MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
'MySQL = MySQL & "                   a.Account_Name          AS CusName,"
'MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "
'
'MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
'MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"
'
'
'MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
'MySQL = MySQL & "                   INNER JOIN Notes"
'MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
'MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
'MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
'MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
'MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
'MySQL = MySQL & "            Where (dev.Posted Is Null)"
'MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
'MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 1"
'MySQL = MySQL & "            Union all"
'MySQL = MySQL & "            SELECT dev.Account_Code,"
'MySQL = MySQL & "                   dev.Credit_Or_Debit,"
'MySQL = MySQL & "                   dev.branch_id,"
'MySQL = MySQL & "                   dev.Notes_ID,"
'MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
'MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
'MySQL = MySQL & "                   Notes1.NoteDate,"
'MySQL = MySQL & "                   Notes1.NoteType,"
'MySQL = MySQL & "                   Notes1.NoteSerial,"
'MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
'MySQL = MySQL & "                   a.Account_Name          AS CusName,"
'MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "
'
'MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
'MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"
'
'
'MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
'MySQL = MySQL & "                    INNER JOIN Notes1"
'MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
'MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
'MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
'MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
'MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
'MySQL = MySQL & "           Where (dev.Posted Is Null)"
'MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
'MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 1"
'MySQL = MySQL & "       ) XB"
'MySQL = MySQL & "       LEFT OUTER JOIN dbo.Ageng_type"
'MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
'MySQL = MySQL & " Where 1 = 1"
''
'If Not IsNull(DTP_Date.value) Then
'    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
'End If
'    If Option1(2).value = True Then
'        mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
'
'    ElseIf Option2.value = True Then
'        mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
'
'    End If
'
'
'     If ChekCustomer.value = vbChecked Then
'        If val(DcbCustomer.BoundText) <> 0 Then
'            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
'            mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
'        End If
'    Else
'         If val(DcbCustomer.BoundText) <> 0 Then
'            mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
'
'        End If
'    End If
'    MySQL = MySQL & mWhereCus
'
'
'
'
'   ' End If
'    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & ")"
'
'        If CheckEmp.value = vbChecked Then
'        If val(DcbEmployee.BoundText) <> 0 Then
'            MySQL = MySQL & " and Account_Code  "
'            MySQL = MySQL & " In (Select Account_Code from TblCustemers Where  TblCustemers.EmpID = " & val(DcbEmployee.BoundText) & ")"
'
'
'        End If
'    End If
'
'
'    MySQL = MySQL & " AND Credit_Or_Debit = 1"
'
'
'
'MySQL = MySQL & "       AND ISNULL(TransNet, 0) <> 0  AND ACCOUNTS.Account_Code = xb.Account_Code"
'MySQL = MySQL & " GROUP BY XB.Account_Code,CusName"
'
'MySQL = MySQL & " ) WHERE 1 = 1 " & mWhereCus
''MySQL = MySQL & " ,AgeID"
'
'
'
'   mSql2 = MySQL
'
'Cn.Execute mSql2

'-------------------------------
   Dim mCusTypeStr As String
   mCusType = 1
   StrCusID.Text = DcbCustomer.BoundText
   
MySQL = ""
MySQL = MySQL & " SELECT "

If mCusType = 2 Then
MySQL = MySQL & " SUM (TransNet) AS TransNet,XB.Account_Code"

Else
MySQL = MySQL & "        XB.Account_Code,"
MySQL = MySQL & " Xb.NotesTypeName ,"
MySQL = MySQL & "        XB.DueDate,"
MySQL = MySQL & "        DiffDate,"
MySQL = MySQL & "        XB.NoteDate,"
MySQL = MySQL & "        XB.NoteType,"
MySQL = MySQL & "        XB.Note_Value TransNet,"
MySQL = MySQL & "        XB.CusName,"
',BalanceAging,"
       '--XB.CusID,
MySQL = MySQL & "        xb.AgeID,"
MySQL = MySQL & "        XB.NoteSerial,"
MySQL = MySQL & "        dbo.Ageng_type.Name,"
MySQL = MySQL & "        dbo.Ageng_type.[From],"
MySQL = MySQL & "        dbo.Ageng_type.[To],"
MySQL = MySQL & "        dbo.Ageng_type.Color,"
MySQL = MySQL & "        dbo.Ageng_type.NameE,"

'       --    BranchId,


If SystemOptions.UserInterface = ArabicInterface Then
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
Else
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'Opening entry') AS TransactionTypeName"
End If
End If

MySQL = MySQL & " FROM   ("
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName,"
Else
    MySQL = MySQL & "                   NotesTypeNamee as NotesTypeName,"
End If

MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes.NoteDate,"
MySQL = MySQL & "                   Notes.NoteType,"
MySQL = MySQL & "                   Notes.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"

If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


'a.BalanceAging,"
MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"


MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
MySQL = MySQL & "                   INNER JOIN Notes"
MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "            Where (dev.Posted Is Null)"
MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 0"
MySQL = MySQL & "            Union all"
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"

If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
Else
    MySQL = MySQL & "                   NotesTypeName = 'Opening entry',"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes1.NoteDate,"
MySQL = MySQL & "                   Notes1.NoteType,"
MySQL = MySQL & "                   Notes1.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"


If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


'a.BalanceAging,"
MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"

    
MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
MySQL = MySQL & "                    INNER JOIN Notes1"
MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "           Where (dev.Posted Is Null)"
MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 0"
MySQL = MySQL & "       ) XB"
MySQL = MySQL & "       Right OUTER JOIN dbo.Ageng_type"
MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
MySQL = MySQL & " Where 1 = 1"
'
If Not IsNull(DTP_Date.value) Then
    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
End If
Option1(2).value = True
    If Option1(2).value = True Then
        mCusType = 1
        mCusTypeStr = "(1)"
        'MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        
    ElseIf Option2.value = True Then
        mCusType = 2
        mCusTypeStr = "(2) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        ElseIf Option3.value = True Then
        mCusType = 57
        mCusTypeStr = "(57) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        ElseIf Option4.value = True Then
        mCusType = 56
        mCusTypeStr = "(56) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
    End If
    ChekCustomer.value = vbChecked
    If ChekCustomer.value = vbChecked Then
        If val(DcbCustomer.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " and Type In " & mCusTypeStr & ")"
            'MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusType & ")"
        Else
            
        End If

    Else
         If val(DcbCustomer.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
        
        End If
        
    End If
    
     
  
  
   ' End If
    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & ")"
   
    
           
If mCusType = 1 Or mCusType = 56 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code   "
End If
   
    
    
    mSql1 = MySQL
    
  
  
   
   MySQL = ""
   
If mCusType = 2 Or mCusType = 57 Then
    MySQL = MySQL & " SELECT SUM(XB.TransNet) AS TransNet ,SUM(XB.TransNet) AS TransNet22,Account_Code,CusName"


Else
      MySQL = MySQL & " Select  SUM (TransNet) AS TransNet,Account_Code,CusName"

End If
   
   


MySQL = ""
MySQL = MySQL & " SELECT "

If mCusType = 1 Or mCusType = 56 Then
MySQL = MySQL & "   SUM (TransNet) AS TransNet,Account_Code,CusName"

Else
MySQL = MySQL & "        XB.Account_Code,"
MySQL = MySQL & " Xb.NotesTypeName ,"
MySQL = MySQL & "        XB.DueDate,"
MySQL = MySQL & "        DiffDate,"
MySQL = MySQL & "        XB.NoteDate,"
MySQL = MySQL & "        XB.NoteType,"
MySQL = MySQL & "        XB.Note_Value TransNet,"
MySQL = MySQL & "        XB.CusName,"
',BalanceAging,"
       '--XB.CusID,
MySQL = MySQL & "        xb.AgeID,"
MySQL = MySQL & "        XB.NoteSerial,"
MySQL = MySQL & "        dbo.Ageng_type.Name,"
MySQL = MySQL & "        dbo.Ageng_type.[From],"
MySQL = MySQL & "        dbo.Ageng_type.[To],"
MySQL = MySQL & "        dbo.Ageng_type.Color,"
MySQL = MySQL & "        dbo.Ageng_type.NameE,"

'       --    BranchId,

If SystemOptions.UserInterface = ArabicInterface Then
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
Else
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'Opening entry') AS TransactionTypeName"
End If

End If

',AgeID"




MySQL = MySQL & " FROM   ("
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName,"
Else
    MySQL = MySQL & "                   NotesTypeNamee as NotesTypeName,"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes.NoteDate,"
MySQL = MySQL & "                   Notes.NoteType,"
MySQL = MySQL & "                   Notes.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"


MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
MySQL = MySQL & "                   INNER JOIN Notes"
MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "            Where (dev.Posted Is Null)"
MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 1"
MySQL = MySQL & "            Union all"
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
Else
    MySQL = MySQL & "                   NotesTypeName = 'Opening entry',"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes1.NoteDate,"
MySQL = MySQL & "                   Notes1.NoteType,"
MySQL = MySQL & "                   Notes1.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"



If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"

    
MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
MySQL = MySQL & "                    INNER JOIN Notes1"
MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "           Where (dev.Posted Is Null)"
MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 1"
MySQL = MySQL & "       ) XB"
MySQL = MySQL & "       Right OUTER JOIN dbo.Ageng_type"
MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
MySQL = MySQL & " Where 1 = 1"
'
If Not IsNull(DTP_Date.value) Then
    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
End If
'    If Option1(2).value = True Then
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
'    ElseIf Option2.value = True Then
'        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
'    End If
    

     If ChekCustomer.value = vbChecked Then
        If val(DcbCustomer.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where   Type In " & mCusTypeStr & " and TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
        
        End If
    Else
         If val(DcbCustomer.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where   Type In " & mCusTypeStr & " and TblCustemers.CusId = " & val(DcbCustomer.BoundText) & " )"
        
        End If
    End If
    
            
       If CheckAllCustomer.value = vbChecked Then
         If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and  TblCustemers.CusID in (" & (StrCusID.Text) & ") )"
        End If
    Else
        If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and TblCustemers.CusID in (" & (StrCusID.Text) & ") )"
        End If
    End If
         
     
   ' End If
    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DcbCustomer.BoundText) & ")"
   
      

    MySQL = MySQL & " AND Credit_Or_Debit = 1"
    


MySQL = MySQL & "       AND ISNULL(TransNet, 0) <> 0"


If mCusType = 2 Or mCusType = 57 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code  ,CusName"
End If
   


'MySQL = MySQL & " ,AgeID"

   
   
   mSql2 = MySQL

    If Option1(2).value = True Or Option4.value Then
        loadgrid mSql1, grdAging, True, False
        loadgrid mSql2, grdAging2, False, False
    Else
        loadgrid mSql2, grdAging, True, False
        loadgrid mSql1, grdAging2, False, False
    End If
    
    
'
'
'    MySQL = MySQL & "                                          AND (DATEDIFF(DAY, '31-Aug-2020', RptLedger_Sub2.RecordDate) < 0)"
'    MySQL = MySQL & "                            ) XB"
'
'    MySQL = MySQL & "                               LEFT OUTER JOIN dbo.Ageng_type"
'    MySQL = MySQL & "                                    ON  XB.ID = dbo.Ageng_type.id"
'
'    MySQL = MySQL & "                        WHERE  XB.DueDate >= '01-Aug-2020'"
'    MySQL = MySQL & "                               AND XB.DueDate <= '31-Aug-2020'"
'
'    MySQL = MySQL & "                        Order By "
'    MySQL = MySQL & "                               Xb.ID , DueDate "
'
'
   

   Dim i As Long
   Dim mValue As Double
   Dim mCusId As Long
    Dim j As Long
    Dim mValue2 As Double
   Dim mCusId2 As Long
    Dim mPayedValue As Double

Dim mAccount_Code As String
Dim mAccount_Code2 As String
Dim Balance As String

'If grdAging.Rows > 1 Then
'    mAccount_Code = Trim(grdAging.TextMatrix(1, grdAging.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'    grdAging.TextMatrix(1, grdAging.ColIndex("Balance")) = Balance
'End If

'If grdAging2.Rows > 1 Then
'    Balance = ""
'    mAccount_Code = Trim(grdAging2.TextMatrix(1, grdAging2.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 1, , , , , FromDate1.value, 1
'    grdAging2.TextMatrix(1, grdAging2.ColIndex("Balance")) = Balance
'End If

Dim mJ As Long
mJ = 1
   For i = 1 To grdAging.rows - 1
   
     
'     If I = 1 Then
'        mAccount_Code = Trim(grdAging.TextMatrix(I, grdAging.ColIndex("Account_Code")))
'        WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'        grdAging.TextMatrix(I, grdAging.ColIndex("Balance")) = Balance
'     End If
      mValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet")))
      mAccount_Code = Trim(grdAging.TextMatrix(i, grdAging.ColIndex("Account_Code")))
      
        If val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) <> mValue Then
        
        
        mJ = grdAging2.FindRow(mAccount_Code, grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
        'mJ = grdAging2.FindRow("dsfdsf", grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
       ' For j = mJ To grdAging2.Rows - 1
'
       j = mJ
            If mJ <> -1 Then
                mValue2 = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")))
                mAccount_Code2 = Trim(grdAging2.TextMatrix(j, grdAging2.ColIndex("Account_Code")))
                mPayedValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
    
    
               If mValue2 <> 0 And mAccount_Code2 = mAccount_Code And mValue <> mPayedValue Then
    
                    If mValue - mPayedValue = mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue > mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) + mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue < mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mPayedValue + mValue - mPayedValue
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = mValue2 - (mValue - mPayedValue)
                        grdAging.TextMatrix(i, grdAging.ColIndex("TransNetGrid2")) = mValue2 - (mValue - mPayedValue)
                        'mValue - grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) + mValue2
                    End If
               End If
               grdAging2.TextMatrix(j, grdAging2.ColIndex("StillAmount")) = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet"))) - val(grdAging2.TextMatrix(j, grdAging2.ColIndex("PayedValue")))
            End If
'
'
'            If mAccount_Code2 <> mAccount_Code Then
'                GoTo ExitFor
'            End If
'        Next
        
      End If
      'mJ = j + 1
ExitFor:
      grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet"))) - val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
      If val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount"))) = 0 Then
        grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = ""
        grdAging.RowHidden(i) = True
      End If
 '     txtTotalStill = val(txtTotalStill) + val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")))
   Next
   Dim S As String
    S = "Delete TblAging "
    Cn.Execute S
    
    
    

    
    S = "Select * from TblAging  "
    
    
    
    saveGrid S, grdAging, "StillAmount", "", "Credit_Or_Debit", 0



    Dim rsDummyT As New ADODB.Recordset
    Dim rsDummyT2 As New ADODB.Recordset
    Dim Rs As New ADODB.Recordset
    S = " Select Account_Code,AGEID,CusName from TblAging"
    S = S & " GROUP BY Account_Code,AGEID,CusName"
    Set rsDummyT = New ADODB.Recordset
    rsDummyT.Open S, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rsDummyT.EOF
        S = "Select * from Ageng_type where Id Not In (Select  AGEID from TblAging Where  Account_Code = N'" & Trim(rsDummyT!Account_code & "") & "' )"
        Set rsDummyT2 = New ADODB.Recordset
        rsDummyT2.Open S, Cn, adOpenStatic, adLockReadOnly
        Do While Not rsDummyT2.EOF
            S = "Select * from TblAging "
            Rs.Open S, Cn, adOpenKeyset, adLockOptimistic
            Rs.AddNew
            Rs!ageid = rsDummyT2!ID
            Rs!Account_code = rsDummyT!Account_code & ""
            Rs!CusName = rsDummyT!CusName & ""
            Rs!To = rsDummyT2!To & ""
            Rs!From = rsDummyT2!From & ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Rs!Name = rsDummyT2!Name & ""
            Else
                Rs!Name = rsDummyT2!Name & ""
            End If
            Rs.update
            Rs.Close
            rsDummyT2.MoveNext
        Loop
        's = "Select Account_Code,AGEID from TblAging Where  Account_Code = " & Trim(rsDummyT!Account_code & "")
        
        rsDummyT.MoveNext
    Loop
    
 
'    s = "Select * from TblAging "
''    saveGrid s, grdAging2, "CusId", "Id", "Credit_Or_Debit", 1
'
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        Else
'            Msg = "No data"
'        End If
'        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    'End If
StrCusID = ""
    RsData.Close
    Set RsData = Nothing
   ' PrintAging Ind
    Screen.MousePointer = vbDefault
End Function



Public Sub LoadSavedGridWa()

    Dim StrSQL As String
    Dim Rs As ADODB.Recordset
    Dim i As Integer

    Set Rs = New ADODB.Recordset

    StrSQL = " SELECT * FROM TblRegDateDelgateDailsGrantee "
    StrSQL = StrSQL & " WHERE DelgID = " & val(Me.XPTxtID.Text)
    StrSQL = StrSQL & " ORDER BY Serial, ID "

    Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    GrdWa.Clear flexClearScrollable, flexClearEverything
    GrdWa.rows = GrdWa.FixedRows

    With GrdWa
        If Not (Rs.BOF Or Rs.EOF) Then
            Rs.MoveFirst
            .rows = .FixedRows + Rs.RecordCount

            For i = .FixedRows To .rows - 1

                .TextMatrix(i, 0) = IIf(IsNull(Rs("Serial").value), "", Rs("Serial").value)
                .TextMatrix(i, .ColIndex("MainID")) = IIf(IsNull(Rs("MainID").value), "", Rs("MainID").value)
                .TextMatrix(i, .ColIndex("MaDate")) = IIf(IsNull(Rs("MaDate").value), "", Rs("MaDate").value)
                .TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(Rs("MainName").value), "", Rs("MainName").value)
                .TextMatrix(i, .ColIndex("Interval")) = IIf(IsNull(Rs("Interval").value), "", Rs("Interval").value)
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs("Remarks").value), "", Rs("Remarks").value)
                .TextMatrix(i, .ColIndex("StatusVisit")) = GetStatusVisitText(Rs("StatusVisit").value)

                Rs.MoveNext
            Next i
        End If
    End With

    Rs.Close
    Set Rs = Nothing

End Sub
