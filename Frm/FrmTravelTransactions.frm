VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmTravelTransactions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· »Ì«‰«  «·—Õ·« "
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18510
   HelpContextID   =   280
   Icon            =   "FrmTravelTransactions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   18510
   Begin VB.CheckBox vbcheck 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   285
      Top             =   9240
      Width           =   135
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   274
      Top             =   8730
      Width           =   1785
   End
   Begin VB.TextBox txtgooglemap 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   13560
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   136
      Top             =   8880
      Width           =   4065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "⁄—÷"
      Height          =   315
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   135
      Top             =   8880
      Width           =   825
   End
   Begin VB.TextBox XPTxtValView 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12720
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   8610
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   4335
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   720
      Width           =   18615
      Begin VB.TextBox TxtManualNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11280
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   194
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   16080
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   273
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox TxtBasedNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   158
         Top             =   120
         Width           =   1425
      End
      Begin VB.ComboBox DcbBasedOn 
         Height          =   315
         Left            =   5760
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   10920
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   148
         Top             =   3870
         Width           =   3315
      End
      Begin VB.Frame Frame5 
         Caption         =   "»Ì«‰«  «·—Õ·…  ÊﬁÌ « "
         Height          =   1815
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   2520
         Width           =   5535
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000002&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtKMCounterAtEnd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtKMCounterBeforeStart 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   960
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DtpStartDate 
            Height          =   315
            Left            =   2640
            TabIndex        =   113
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   141950977
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DtpEndDate 
            Height          =   315
            Left            =   2640
            TabIndex        =   115
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   141950977
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker StartTime 
            Height          =   285
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   141950979
            UpDown          =   -1  'True
            CurrentDate     =   40909
         End
         Begin MSComCtl2.DTPicker EndTime 
            Height          =   285
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   141950979
            UpDown          =   -1  'True
            CurrentDate     =   40909
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ê“‰"
            Height          =   285
            Index           =   46
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ê“‰"
            Height          =   285
            Index           =   45
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   960
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁ—«¡… «·⁄œ«œ ⁄‰œ «·Ê’Ê·"
            Height          =   405
            Index           =   40
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁ—«¡… «·⁄œ«œ ﬁ»· «·–Â«»"
            Height          =   405
            Index           =   39
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Êﬁ  «·Ê’Ê·"
            Height          =   285
            Index           =   38
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Êﬁ  «·–Â«»"
            Height          =   285
            Index           =   37
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·Ê’Ê·"
            Height          =   285
            Index           =   36
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·–Â«»"
            Height          =   285
            Index           =   35
            Left            =   4380
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   255
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "»Ì«‰«  «·—Õ·… «·„«·Ì…"
         Height          =   2055
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtNoR 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   1320
            Width           =   1185
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   1680
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox TxtComm 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   1680
            Width           =   1425
         End
         Begin VB.TextBox txtDriverEra 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   1320
            Width           =   1425
         End
         Begin VB.TextBox TxtDriverPercentage 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   960
            Width           =   1425
         End
         Begin VB.TextBox txtTotalExpenses 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   960
            Width           =   1185
         End
         Begin VB.TextBox txtDriverValue 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox TXTTravelPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox TxtKmPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox TxtDistance 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·—œÊœ"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   88
            Left            =   1260
            RightToLeft     =   -1  'True
            TabIndex        =   287
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   47
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄„Ê·…"
            Height          =   285
            Index           =   44
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·”«∆ﬁ %"
            Height          =   285
            Index           =   43
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„’—Ê›« "
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   34
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄Âœ…«·”«∆ﬁ"
            Height          =   285
            Index           =   33
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·œÌ“·"
            Height          =   285
            Index           =   32
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„”«›… ﬂ„"
            Height          =   285
            Index           =   31
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—œ «·—Õ·…"
            Height          =   285
            Index           =   30
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—œ «·”«∆ﬁ"
            Height          =   285
            Index           =   29
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "»Ì«‰«  «·—Õ·…"
         Height          =   3135
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   480
         Width           =   8655
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2880
            TabIndex        =   176
            Top             =   1680
            Width           =   570
         End
         Begin VB.TextBox TxtIDNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   2760
            Width           =   1395
         End
         Begin VB.TextBox TxtTypGoods 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   2760
            Width           =   1515
         End
         Begin VB.TextBox TxtOrderNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   2760
            Width           =   1275
         End
         Begin VB.TextBox TxtLeaderName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox TxtLocation 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   960
            Width           =   3195
         End
         Begin MSDataListLib.DataCombo DcCityFromId 
            Height          =   315
            Left            =   4440
            TabIndex        =   93
            Top             =   600
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCCar 
            Height          =   315
            Left            =   120
            TabIndex        =   95
            Top             =   1320
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCEmp 
            Height          =   315
            Left            =   4440
            TabIndex        =   97
            Top             =   2400
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCityToId 
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   600
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo VehicleType 
            Height          =   315
            Left            =   4440
            TabIndex        =   146
            Top             =   1320
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbShip 
            Height          =   315
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSupplem 
            Height          =   315
            Left            =   4440
            TabIndex        =   163
            Top             =   1680
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCar2 
            Height          =   315
            Left            =   4440
            TabIndex        =   165
            Top             =   2040
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSupplem2 
            Height          =   315
            Left            =   120
            TabIndex        =   167
            Top             =   2040
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName2 
            Height          =   315
            Left            =   120
            TabIndex        =   175
            Top             =   1680
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbHarbor 
            Height          =   315
            Left            =   4440
            TabIndex        =   186
            Top             =   240
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton ChCarType 
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   201
            Top             =   960
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton ChCarType 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   202
            Top             =   960
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "„„·Êﬂ… ··€Ì—"
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„—ﬂ»…"
            Height          =   285
            Index           =   81
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   272
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ì‰«¡"
            Height          =   285
            Index           =   55
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «„— «· Õ„Ì·"
            Height          =   285
            Index           =   62
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÂÊÌ…"
            Height          =   285
            Index           =   60
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·»÷«⁄…"
            Height          =   285
            Index           =   61
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ê—œ"
            Height          =   285
            Index           =   64
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”«∆ﬁ Œ«—ÃÌ"
            Height          =   285
            Index           =   59
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„·Õﬁ"
            Height          =   285
            Index           =   58
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„„·ÊﬂÂ ··€Ì—"
            Height          =   285
            Index           =   57
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„·Õﬁ"
            Height          =   285
            Index           =   56
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·”›Ì‰…"
            Height          =   285
            Index           =   54
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„—ﬂ»…"
            Height          =   285
            Index           =   49
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   285
            Index           =   42
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃÂ… «·Ê’Ê·"
            Height          =   285
            Index           =   41
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·”«∆ﬁ"
            Height          =   285
            Index           =   28
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
            Height          =   285
            Index           =   26
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·—Õ·… „‰ "
            Height          =   285
            Index           =   25
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkDestribute 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ê“⁄"
         Enabled         =   0   'False
         Height          =   195
         Left            =   18960
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CBoBasedON 
         Height          =   315
         Left            =   18720
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1110
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   19320
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2790
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   990
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   14520
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   3285
         Left            =   14400
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   960
         Width           =   4155
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   2520
            Width           =   1005
         End
         Begin VB.TextBox TxtPrice1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   2520
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TxtPartPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   2880
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtTotal1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   2880
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   840
            Width           =   2715
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   1980
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Format          =   142409729
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   120
            TabIndex        =   90
            Top             =   1200
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   0
            Left            =   1965
            TabIndex        =   150
            Top             =   2280
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ„"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   1
            Left            =   885
            TabIndex        =   151
            Top             =   2280
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·—œ"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   2
            Left            =   45
            TabIndex        =   152
            Top             =   2280
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·Ê“‰"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   120
            TabIndex        =   209
            Top             =   1560
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
            Caption         =   "ÃÂ… «·’—›"
            Height          =   285
            Index           =   79
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ‹ ··„·Õﬁ"
            Height          =   285
            Index           =   67
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   2880
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   66
            Left            =   420
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   2880
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ‹ ··„—ﬂ»…"
            Height          =   285
            Index           =   65
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   2520
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÊÕœ… «·ﬁÌ«”"
            Height          =   285
            Index           =   50
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   2280
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   285
            Index           =   48
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   24
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1200
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   285
            Index           =   16
            Left            =   2790
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   2790
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1980
            Width           =   1275
         End
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   16080
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   5760
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   3870
         Width           =   4995
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   20160
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3150
         Width           =   2655
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14760
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   13860
         TabIndex        =   66
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143523841
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   2310
         TabIndex        =   67
         Top             =   10710
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
         Left            =   19560
         TabIndex        =   68
         Top             =   1110
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmTravelTransactions.frx":038A
         Height          =   315
         Left            =   18960
         TabIndex        =   69
         Top             =   630
         Width           =   3375
         _ExtentX        =   5953
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
         Bindings        =   "FrmTravelTransactions.frx":039F
         Height          =   315
         Left            =   8040
         TabIndex        =   80
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.RadioButton ChTripType 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   161
         Top             =   120
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "œ«Œ·Ì"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ChTripType 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   162
         Top             =   120
         Width           =   870
         _Version        =   786432
         _ExtentX        =   1535
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Œ«—ÃÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
         Height          =   285
         Index           =   82
         Left            =   12420
         RightToLeft     =   -1  'True
         TabIndex        =   219
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·—Õ·…"
         Height          =   195
         Index           =   63
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   171
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«„—"
         Height          =   195
         Index           =   53
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   195
         Index           =   52
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   285
         Index           =   0
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   195
         Index           =   22
         Left            =   18780
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   2790
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·—Õ·…"
         Height          =   285
         Index           =   4
         Left            =   17100
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   18840
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   15360
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   135
         Width           =   555
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   0
         Picture         =   "FrmTravelTransactions.frx":03B4
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
         Left            =   18600
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   195
         Index           =   15
         Left            =   17220
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   19080
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« /‰Ê⁄ «·Õ„Ê·…"
         Height          =   285
         Index           =   20
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   3630
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   405
         Index           =   21
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   3120
         Width           =   1275
      End
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   10
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
      Left            =   2580
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   10740
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
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12720
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8310
      Width           =   1905
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   0
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   18495
      _cx             =   32623
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
      Picture         =   "FrmTravelTransactions.frx":093E
      Caption         =   " ”ÃÌ· »Ì«‰«  «·—Õ·«   "
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
      Begin VB.TextBox oldTxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   12
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmTravelTransactions.frx":1618
         ColorButton     =   16777215
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
         TabIndex        =   13
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmTravelTransactions.frx":19B2
         ColorButton     =   16777215
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
         TabIndex        =   14
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmTravelTransactions.frx":1D4C
         ColorButton     =   16777215
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
         TabIndex        =   15
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmTravelTransactions.frx":20E6
         ColorButton     =   16777215
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
      Left            =   21600
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   7560
      TabIndex        =   17
      Top             =   8850
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   16380
      TabIndex        =   24
      Top             =   9240
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
      Left            =   15360
      TabIndex        =   25
      Top             =   9270
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
      Left            =   14400
      TabIndex        =   26
      Top             =   9270
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
      Left            =   13635
      TabIndex        =   27
      Top             =   9270
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
      Left            =   11640
      TabIndex        =   28
      Top             =   9270
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
      Left            =   3720
      TabIndex        =   29
      Top             =   9270
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
      Left            =   8640
      TabIndex        =   30
      Top             =   9270
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
      Left            =   10710
      TabIndex        =   31
      Top             =   9240
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
      Height          =   315
      Left            =   5640
      TabIndex        =   42
      Top             =   8730
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
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
      MICON           =   "FrmTravelTransactions.frx":2480
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
      Left            =   9720
      TabIndex        =   43
      Top             =   9270
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   6120
      TabIndex        =   44
      Top             =   9360
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
      Left            =   16320
      TabIndex        =   45
      Tag             =   "Delete Row"
      Top             =   8400
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
      MICON           =   "FrmTravelTransactions.frx":249C
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
      Left            =   4680
      TabIndex        =   46
      Top             =   9360
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   3270
      Left            =   0
      TabIndex        =   82
      Top             =   5040
      Width           =   18585
      _cx             =   32782
      _cy             =   5768
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
      Caption         =   "‰”» «· Ê“Ì⁄|New Tab|»Ì«‰«  «‰Ê«⁄ «·‰ﬁ·|”Ì«—«  «·€Ì—|«·„’—Ê›«  Ê Õ„Ì· «·«’‰«›|«Ê«„— «· Õ„Ì·"
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
      Flags(0)        =   2
      Flags(1)        =   2
      Flags(3)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   2850
         Left            =   -19740
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VSFlex8Ctl.VSFlexGrid GridEstimatedCost 
            Height          =   2115
            Left            =   0
            TabIndex        =   84
            Top             =   0
            Width           =   10425
            _cx             =   18389
            _cy             =   3731
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmTravelTransactions.frx":24B8
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2850
         Index           =   2
         Left            =   -20040
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   23
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   90
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2850
         Index           =   4
         Left            =   -19440
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VB.TextBox loadingInvoice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   281
            Top             =   480
            Width           =   1410
         End
         Begin VB.TextBox txtWeight 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5100
            TabIndex        =   277
            Top             =   480
            Width           =   1410
         End
         Begin VB.TextBox txtRecNo 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   8040
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   275
            Top             =   480
            Width           =   1410
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   16605
            TabIndex        =   5
            Top             =   480
            Width           =   900
         End
         Begin VB.TextBox QtyDischarge 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2055
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox CardNO2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5100
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox QtyDownload 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8025
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox CardNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11010
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox TxtQtyDischarge 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   2520
            Width           =   1365
         End
         Begin VB.TextBox txtid 
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
            Index           =   2
            Left            =   -4710
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   8970
            Width           =   2535
         End
         Begin VB.TextBox TxtQtyDownload 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9615
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   2520
            Width           =   1440
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
            Height          =   1410
            Left            =   180
            TabIndex        =   190
            Top             =   960
            Width           =   18135
            _cx             =   31988
            _cy             =   2487
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
            Rows            =   1
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":27A9
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
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2160
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   191
               Top             =   2400
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox Text10 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   192
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                  ForeColor       =   &H0000C8FF&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   0
                  Width           =   2445
               End
            End
            Begin VDSCOMBOLibCtl.SmartCombo CboDes 
               Height          =   315
               Left            =   0
               TabIndex        =   279
               ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
               Top             =   600
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
               Picture         =   "FrmTravelTransactions.frx":29A1
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
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   390
            Left            =   16950
            TabIndex        =   196
            Top             =   2400
            Width           =   1365
            _ExtentX        =   2408
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
            ButtonImage     =   "FrmTravelTransactions.frx":2F3B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbTypeTransport 
            Height          =   315
            Left            =   14055
            TabIndex        =   0
            Top             =   120
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboItems 
            Height          =   315
            Left            =   14055
            TabIndex        =   6
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   675
            Left            =   180
            TabIndex        =   7
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   1191
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmTravelTransactions.frx":34D5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker BillDate 
            Height          =   315
            Left            =   11010
            TabIndex        =   230
            Top             =   480
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Format          =   143065089
            CurrentDate     =   38784
         End
         Begin ImpulseAniLabel.ISAniLabel LblLink 
            Height          =   315
            Left            =   90
            TabIndex        =   288
            Top             =   2490
            Width           =   1920
            _ExtentX        =   3387
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
            MouseIcon       =   "FrmTravelTransactions.frx":9D37
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
            Caption         =   "—ﬁ„ ﬂ—  „Ì“«‰  ›—Ì€"
            Height          =   285
            Index           =   76
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   282
            Top             =   120
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ ›« Ê—… «· Õ„Ì· "
            Height          =   285
            Index           =   87
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   280
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ê“‰"
            Height          =   360
            Index           =   85
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   278
            Top             =   495
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «–‰ «· ”·Ì„"
            Height          =   285
            Index           =   84
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   276
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   285
            Index           =   83
            Left            =   12465
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   495
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’‰›"
            Height          =   285
            Index           =   78
            Left            =   17640
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ„Ì… «· ›—Ì€"
            Height          =   285
            Index           =   77
            Left            =   3675
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ„Ì… «· Õ„Ì·"
            Height          =   285
            Index           =   75
            Left            =   9465
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ ﬂ—   „Ì“«‰   Õ„Ì·"
            Height          =   285
            Index           =   74
            Left            =   12465
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·‰ﬁ·"
            Height          =   285
            Index           =   72
            Left            =   17130
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label20 
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
            Height          =   345
            Left            =   16770
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   930
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì« "
            Height          =   285
            Index           =   73
            Left            =   12240
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   2520
            Width           =   810
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2850
         Index           =   5
         Left            =   -19140
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VB.TextBox TxtQtyDownload2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9615
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   2520
            Width           =   1440
         End
         Begin VB.TextBox txtid 
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
            Left            =   -4710
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   8970
            Width           =   2535
         End
         Begin VB.TextBox TxtQtyDischarge2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   2520
            Width           =   1365
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
            Height          =   2010
            Left            =   180
            TabIndex        =   215
            Top             =   480
            Width           =   18135
            _cx             =   31988
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
            Rows            =   1
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":9E99
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
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2160
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   216
               Top             =   2400
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox Text16 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   217
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                  ForeColor       =   &H0000C8FF&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   218
                  Top             =   0
                  Width           =   2445
               End
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«—«  «·€Ì—"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   80
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   271
            Top             =   120
            Width           =   2730
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì« "
            Height          =   285
            Index           =   86
            Left            =   12240
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label Label14 
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
            Height          =   345
            Left            =   16770
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   930
            Width           =   1275
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2850
         Index           =   1
         Left            =   45
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VB.TextBox txtid 
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
            Index           =   0
            Left            =   -4710
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   10920
            Width           =   2535
         End
         Begin VB.TextBox TxtTotalQty 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   11520
            RightToLeft     =   -1  'True
            TabIndex        =   224
            Top             =   2385
            Width           =   1440
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   9435
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   223
            Top             =   60
            Width           =   8430
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   1635
            Left            =   9525
            TabIndex        =   226
            Top             =   600
            Width           =   8790
            _cx             =   15505
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
            Rows            =   1
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":A001
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
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2160
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   227
               Top             =   2400
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox Text6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   228
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                  ForeColor       =   &H0000C8FF&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   229
                  Top             =   0
                  Width           =   2445
               End
            End
         End
         Begin ImpulseButton.ISButton CmdDelete 
            Height          =   420
            Left            =   17130
            TabIndex        =   231
            Top             =   2385
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   741
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
            ButtonImage     =   "FrmTravelTransactions.frx":A0F4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   2250
            Left            =   180
            TabIndex        =   232
            Top             =   120
            Width           =   9165
            _cx             =   16166
            _cy             =   3969
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":A68E
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
            Begin VB.PictureBox PicDes 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2160
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   233
               Top             =   2400
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
                  TabIndex        =   234
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
                  TabIndex        =   235
                  Top             =   0
                  Width           =   2445
               End
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2175
            Left            =   -90
            TabIndex        =   237
            Top             =   105
            Visible         =   0   'False
            Width           =   13605
            _cx             =   23998
            _cy             =   3836
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
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":A825
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
               TabIndex        =   249
               Top             =   3720
               Visible         =   0   'False
               Width           =   4215
               Begin VB.CommandButton Command5 
                  Caption         =   "‰”Œ"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   250
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   255
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   252
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   3915
               Left            =   -1650
               RightToLeft     =   -1  'True
               ScaleHeight     =   3915
               ScaleWidth      =   9405
               TabIndex        =   238
               Top             =   2130
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
                  TabIndex        =   242
                  Top             =   2040
                  Width           =   8955
               End
               Begin VB.TextBox txtcodesub 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   241
                  Top             =   3600
                  Width           =   855
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "Add des"
                  Height          =   255
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   240
                  Top             =   3600
                  Width           =   1350
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Call des"
                  Height          =   255
                  Left            =   6240
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   3600
                  Width           =   1095
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   3900
                  Left            =   -5760
                  TabIndex        =   243
                  TabStop         =   0   'False
                  Top             =   480
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
                     Left            =   -3840
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   244
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
                     TabIndex        =   245
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
                  TabIndex        =   248
                  Top             =   3480
                  Width           =   735
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Height          =   495
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   247
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Code"
                  Height          =   255
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   246
                  Top             =   1320
                  Width           =   735
               End
            End
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   420
            Left            =   7920
            TabIndex        =   284
            Top             =   2400
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   741
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
            ButtonImage     =   "FrmTravelTransactions.frx":AB44
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label9 
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
            Height          =   360
            Left            =   16770
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   945
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   51
            Left            =   12420
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Top             =   2385
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«„—"
            Height          =   285
            Index           =   5
            Left            =   17130
            RightToLeft     =   -1  'True
            TabIndex        =   254
            Top             =   150
            Width           =   1275
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2850
         Index           =   3
         Left            =   19230
         TabIndex        =   257
         TabStop         =   0   'False
         Top             =   45
         Width           =   18495
         _cx             =   32623
         _cy             =   5027
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
         Begin VB.TextBox txtid 
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
            Index           =   1
            Left            =   -4710
            RightToLeft     =   -1  'True
            TabIndex        =   258
            Top             =   9000
            Width           =   2535
         End
         Begin VSFlex8Ctl.VSFlexGrid FGOrders 
            Height          =   2100
            Left            =   180
            TabIndex        =   259
            Top             =   375
            Width           =   18135
            _cx             =   31988
            _cy             =   3704
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
            Rows            =   1
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTravelTransactions.frx":B0DE
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
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2160
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   260
               Top             =   2400
               Visible         =   0   'False
               Width           =   2925
               Begin VB.TextBox Text9 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1125
                  Left            =   30
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   261
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                  ForeColor       =   &H0000C8FF&
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   262
                  Top             =   0
                  Width           =   2445
               End
            End
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   375
            Left            =   17130
            TabIndex        =   264
            Top             =   2400
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
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
            ButtonImage     =   "FrmTravelTransactions.frx":B2F6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   315
            Left            =   17310
            TabIndex        =   265
            Top             =   0
            Width           =   1005
            _Version        =   786432
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   " ÕœÌœ «·ﬂ·"
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«ÌÃ«—"
            Height          =   285
            Index           =   69
            Left            =   14955
            RightToLeft     =   -1  'True
            TabIndex        =   270
            Top             =   2505
            Width           =   1005
         End
         Begin VB.Label Label18 
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
            Height          =   360
            Left            =   16770
            RightToLeft     =   -1  'True
            TabIndex        =   269
            Top             =   945
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   68
            Left            =   12690
            RightToLeft     =   -1  'True
            TabIndex        =   268
            Top             =   2505
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì "
            Height          =   285
            Index           =   70
            Left            =   2355
            RightToLeft     =   -1  'True
            TabIndex        =   267
            Top             =   2505
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   71
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   266
            Top             =   2505
            Width           =   2085
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   131
      Top             =   9240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin MSDataListLib.DataCombo DcboCreditSide2 
      Height          =   315
      Left            =   0
      TabIndex        =   132
      Top             =   8640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   17400
      TabIndex        =   143
      Tag             =   "Delete Row"
      Top             =   8400
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
      MICON           =   "FrmTravelTransactions.frx":B890
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
      Index           =   11
      Left            =   12600
      TabIndex        =   253
      Top             =   9270
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "‰”Œ… „„«À·…"
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
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   315
      Left            =   5640
      TabIndex        =   283
      Top             =   8400
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   " „ «·«‰ Â«¡"
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
      MICON           =   "FrmTravelTransactions.frx":B8AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   390
      Index           =   0
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   263
      Top             =   8745
      Width           =   855
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
      Height          =   2115
      Index           =   27
      Left            =   -1920
      RightToLeft     =   -1  'True
      TabIndex        =   142
      Top             =   8280
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "  »⁄ GPS"
      Height          =   285
      Index           =   33
      Left            =   17520
      RightToLeft     =   -1  'True
      TabIndex        =   137
      Top             =   8880
      Width           =   930
   End
   Begin VB.Label LblValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8460
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   1380
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   9090
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   9090
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   9090
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
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   9090
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   11505
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8865
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8400
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "FrmTravelTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
