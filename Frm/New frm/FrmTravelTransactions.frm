VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
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
      Visible         =   0   'False
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
         Top             =   150
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
            Format          =   253952001
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
            Format          =   253952001
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
            Format          =   253952003
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
            Format          =   253952003
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
         Begin VB.TextBox txtCityToCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2070
            RightToLeft     =   -1  'True
            TabIndex        =   302
            Top             =   600
            Width           =   1365
         End
         Begin VB.TextBox txtCityFromCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6570
            RightToLeft     =   -1  'True
            TabIndex        =   301
            Top             =   600
            Width           =   1065
         End
         Begin VB.TextBox txtContainerNo 
            BackColor       =   &H0000FFFF&
            Height          =   345
            Left            =   0
            TabIndex        =   299
            Top             =   180
            Width           =   1725
         End
         Begin VB.TextBox TxtRent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   289
            Top             =   960
            Width           =   705
         End
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
            Width           =   2115
            _ExtentX        =   3731
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
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   2730
            TabIndex        =   159
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
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
            Left            =   5490
            TabIndex        =   186
            Top             =   240
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton ChCarType 
            Height          =   255
            Index           =   0
            Left            =   2160
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
            Left            =   960
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
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «—«„ﬂÊ"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   1710
            TabIndex        =   300
            Top             =   270
            Width           =   1050
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
            Left            =   4020
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
            Top             =   2580
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
            Top             =   2580
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
            Top             =   2940
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
            Top             =   2940
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   780
            Width           =   2715
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Format          =   207028225
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   450
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
            Top             =   1110
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
            Left            =   2175
            TabIndex        =   150
            Top             =   2340
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
            Top             =   2340
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
            Top             =   2340
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
            Top             =   1470
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   3
            Left            =   1740
            TabIndex        =   290
            Top             =   2100
            Width           =   1305
            _Version        =   786432
            _ExtentX        =   2302
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»«·· —/«·Õ„Ê·…"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyped 
            Height          =   255
            Index           =   4
            Left            =   900
            TabIndex        =   291
            Top             =   2100
            Visible         =   0   'False
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»«·Õ„Ê·…"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
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
            Top             =   1470
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
            Top             =   2940
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
            Top             =   2940
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
            Top             =   2580
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
            Top             =   2580
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
            Top             =   1110
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
            Top             =   480
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
            Top             =   780
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
            Top             =   1800
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
         Format          =   207028225
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
      Visible         =   0   'False
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
      Caption         =   "‰”» «· Ê“Ì⁄|New Tab|»Ì«‰«  «‰Ê«⁄ «·‰ﬁ·|”Ì«—«  «·€Ì—|«·„’—Ê›«  Ê Õ„Ì· «·«’‰«›|«Ê«„— «· Õ„Ì·|«·«⁄ „«œ"
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
            Top             =   360
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
            Cols            =   22
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
               Picture         =   "FrmTravelTransactions.frx":2AFA
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
            ButtonImage     =   "FrmTravelTransactions.frx":3094
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbTypeTransport 
            Height          =   315
            Left            =   14055
            TabIndex        =   0
            Top             =   30
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
            Top             =   360
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
            ButtonImage     =   "FrmTravelTransactions.frx":362E
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
            Format          =   254148609
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
            MouseIcon       =   "FrmTravelTransactions.frx":9E90
            Alignment       =   1
            Caption         =   ""
            ColorHover      =   16711680
            RightToLeft     =   -1  'True
            ImageCount      =   0
         End
         Begin MSDataListLib.DataCombo cmbUnitName 
            Height          =   315
            Left            =   14070
            TabIndex        =   292
            Top             =   660
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
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÊÕœ…"
            Height          =   285
            Index           =   89
            Left            =   16740
            RightToLeft     =   -1  'True
            TabIndex        =   293
            Top             =   690
            Width           =   645
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
            Top             =   360
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
            Top             =   30
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
            FormatString    =   $"FrmTravelTransactions.frx":9FF2
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
            Index           =   1
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
            FormatString    =   $"FrmTravelTransactions.frx":A15A
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
            ButtonImage     =   "FrmTravelTransactions.frx":A24D
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
            FormatString    =   $"FrmTravelTransactions.frx":A7E7
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
            FormatString    =   $"FrmTravelTransactions.frx":A97E
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
            ButtonImage     =   "FrmTravelTransactions.frx":AC9D
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
            FormatString    =   $"FrmTravelTransactions.frx":B237
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
            ButtonImage     =   "FrmTravelTransactions.frx":B44F
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   2850
         Left            =   19530
         TabIndex        =   294
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
            Height          =   2100
            Left            =   210
            TabIndex        =   295
            Tag             =   "1"
            Top             =   120
            Width           =   18240
            _cx             =   32173
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
            FormatString    =   $"FrmTravelTransactions.frx":B9E9
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
         Begin ImpulseButton.ISButton Accredit 
            Height          =   390
            Left            =   180
            TabIndex        =   296
            Top             =   2280
            Width           =   1080
            _ExtentX        =   1905
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
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   6780
            RightToLeft     =   -1  'True
            TabIndex        =   298
            Top             =   2400
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   13080
            RightToLeft     =   -1  'True
            TabIndex        =   297
            Top             =   4560
            Width           =   3375
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
      MICON           =   "FrmTravelTransactions.frx":BB2C
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
      MICON           =   "FrmTravelTransactions.frx":BB48
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
      Visible         =   0   'False
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
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim Destribute As Boolean
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim Account_Code_dynamic3 As String
Dim RentValue As Double

Function CuurentLogdata(Optional Currentmode As String)
TxtSerial1.text = txtNoteSerial1.text
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·”‰œ " & TxtSerial1.text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & Dcbranch & CHR(13) & "   „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPaymentType & CHR(13) & "   «·„‘—Ê⁄  " & dcproject & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "  »‰«¡ ⁄·Ï " & txtto & CHR(13) & "   »‰«¡ ⁄·Ï  " & CBoBasedON & "  »—ﬁ„  " & txt_ORDER_NO & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·”‰œ    " & XPTxtValView
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. No " & TxtSerial1.text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & Dcbranch & CHR(13) & "   CC  " & DcCostCenter & CHR(13) & "  Payment Type  " & CboPaymentType & CHR(13) & "   Project  " & dcproject & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Based On " & txtto & CHR(13) & "   Based On  " & CBoBasedON & "  No:  " & txt_ORDER_NO & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Vchr Total   " & XPTxtValView
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
    Else
        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
    End If
End Function

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "notes_all", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.text), TxtSerial1.text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
    Accredit.Caption = "Sent To Approval "
End If
    Retrive (val(Me.XPTxtID.text))

End Sub

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
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… «Ê·« ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If
            Exit Sub
        End If
        marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
        marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
        marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
        marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
        marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        marakes_taklefa_tawze3.Adodc3.Refresh
        'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub RemoveTyptransRow()
    With Me.VSFlexGrid3
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveFgOrdersRow()
    With Me.FGOrders
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.VSFlexGrid2
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub

Private Sub ALLButton3_Click()
If Me.TxtModFlg.text = "R" Then
Cn.Execute "Update TblOrderUpload set OrderStuts =1 where ID=" & val(TxtBasedNo.text) & " "
MsgBox " „ «·«‰ Â«¡"
TxtModFlg_Change
End If
End Sub
Function CheckOrderStuts() As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select OrderStuts from TblOrderUpload where  ID=" & val(TxtBasedNo.text) & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckOrderStuts = IIf(IsNull(rs2("OrderStuts").value), 0, rs2("OrderStuts").value)
Else
CheckOrderStuts = 0
End If
End Function
Private Sub CBoBasedON_Change()

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
                lbl(21).Caption = "—ﬁ„ «·› Ê—… «·„»œ∆ÌÂ"
            Else
                lbl(21).Caption = "Performa Invoice NO"
            End If
        ElseIf Me.CBoBasedON.ListIndex = 3 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If
        End If
        .TextMatrix(0, .ColIndex("order_no")) = lbl(21).Caption
    End With
End Sub
Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub
Private Sub CboPayMentType_Change()
    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DBCboClientName.text = ""
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
        lbl(19).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    Else
        lbl(18).Caption = "Cheque No"
        lbl(19).Caption = "Due Date"
    End If

    If Me.CboPaymentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.lbl(24).Enabled = True
        DBCboClientName.Enabled = True
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = True
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·ÕÊ«·… "
            lbl(19).Caption = " «—ÌŒÂ«"
        Else
            lbl(18).Caption = "Transfer No"
            lbl(19).Caption = "Date"
        End If
    Else
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
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

Private Sub ChCarType_Click(Index As Integer)
    If ChCarType(0).value = True Then
        DcbSupplem.Enabled = True
        DCCar.Enabled = True
        TxtSearchCode.Enabled = False
        TxtSearchCode.text = ""
        DBCboClientName2.Enabled = False
        DcbCar2.Enabled = False
        DcbSupplem2.Enabled = False
        DBCboClientName2.BoundText = ""
        DcbCar2.BoundText = ""
    Else
        DcbSupplem2.Enabled = True
        DcbCar2.Enabled = True
        DBCboClientName2.Enabled = True
        TxtSearchCode.Enabled = True
        DCCar.BoundText = ""
        DCCar.Enabled = False
        DcbSupplem.BoundText = ""
        DcbSupplem.Enabled = False
    End If
End Sub

Private Sub CheckBox1_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    If CheckBox1.value = vbChecked Then

        With Me.FGOrders
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Selct")) = True
            Next i

        End With
    End If

    Else

        With Me.FGOrders
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Selct")) = False
            Next i

        End With

    End If
    ReLineGrid
    
End Sub

Private Sub cmbUnitName_Click(Area As Integer)
Dim sql As String
Dim rsDummy  As New ADODB.Recordset
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    sql = "   SELECT        TblClientTransContr.ID, TblClientTransContr.CompID, TblClientTransContr.LockedID, TblClientTransContr.UserID, TblClientTransContr.FromDate, TblClientTransContr.Todate,"
    sql = sql & "                       TblClientTransContr.Remarks, TblClientTransContr.Typed,  TblVendorCars.BoardNo,  TblClientTransContrDet.ClintTransID,"
    sql = sql & "                          TblClientTransContrDet.VehicleType , TblClientTransContrDet.Price , TblClientTransContrDet.Remarks AS Expr7,  TblClientTransContrDet.FromPrice,"
    sql = sql & "                          TblClientTransContrDet.ToPrice, TblClientTransContrDet.FromCityID, TblClientTransContrDet.ToCityID, TblCountriesGovernments_1.GovernmentName as ToCityName, TblCountriesGovernments.GovernmentName AS FromCity,"
    sql = sql & "                          TblCustemers.CusName , TblCustemers.CusID, TblCustemers.CusNamee, TblVendorCars.nBoardNo, TblVendorCars.ChasisNo, TblVendorCars.BrandID, TblVendorCars.ModelID,"
    sql = sql & "                   TblItems.ItemName,TblItems.ItemNamee,TblItems.ItemCode,TblItems.ItemID,TblUnites.UnitID,TblUnites.UnitName,TblUnites.UnitNamee"
    sql = sql & " FROM            TblCountriesGovernments AS TblCountriesGovernments_1 RIGHT OUTER JOIN"
    sql = sql & "                          TblClientTransContrDet ON TblCountriesGovernments_1.GovernmentID = TblClientTransContrDet.ToCityID LEFT OUTER JOIN"
    sql = sql & "                          TblCountriesGovernments ON TblClientTransContrDet.FromCityID = TblCountriesGovernments.GovernmentID RIGHT OUTER JOIN"
    sql = sql & "                          TblClientTransContr LEFT OUTER JOIN"
    sql = sql & "                          TblVendorCars ON TblClientTransContr.VehicleType = TblVendorCars.ID LEFT OUTER JOIN"
    sql = sql & "                          TblCustemers ON TblClientTransContr.CompID = TblCustemers.CusID ON TblClientTransContrDet.ClintTransID = TblClientTransContr.ID"
    sql = sql & "                  LEFT OUTER JOIN TblItems  On TblItems.ItemID =TblClientTransContrDet.ItemID"
    sql = sql & "                  LEFT OUTER JOIN TblUnites  On TblUnites.UnitID =TblClientTransContrDet.UnitID"
'    sql = sql & " Where (1 = 1) and TblClientTransContr.ID=" & ID
    
    
    If Trim(VehicleType.text) = "" Then
    sql = sql & " Where  (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
Else
    sql = sql & " Where (dbo.TblClientTransContr.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
End If

sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""

    sql = sql & " and TblItems.ItemID = " & val(DcboItems.BoundText)
    sql = sql & " and TblUnites.UnitID = " & val(cmbUnitName.BoundText)
    
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open sql, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        
        TXTTravelPrice = rsDummy!Price & ""
        TxtPrice = rsDummy!Price & ""
    End If
    
    
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
Dim IntRes As Integer
 ' On Error GoTo ErrTrap
  
    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            DcCostCenter.text = ""
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
            BillDate.value = Date
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
            
         
             
            
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 2
            FGOrders.Clear flexClearScrollable, flexClearEverything
            FGOrders.rows = 2
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 2
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Me.VSFlexGrid1.rows = 2
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.rows = 1
            Fg_Journal.Enabled = True
            DtpChequeDueDate.value = Date
            setfoxy
            CBoBasedON.ListIndex = 0
            Me.Dcbranch.BoundText = Current_branch
            DcbBasedOn_Change
            
            RdTyped(1).value = True
               DcCityFromId.BoundText = 2
             DcCityToId.BoundText = 2
             
             CboPaymentType.ListIndex = 1
ChCarType(0).value = True

               GetTripInformations
    RetriveClinCounr
    
        Case 1
        'If SystemOptions.CanChangeTripAfterInvoiceing = False Then
                        If ChekClodePeriod(XPDtbTrans.value) = True Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                            Else
                                MsgBox "Please Change Date Becouse This is Period is Closed"
                            End If
                            Exit Sub
                        End If
        '  End If
            Dim Msg As String

            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
            
If SystemOptions.CanChangeTripAfterInvoiceing = False Then
                If CheckAllocation() = True Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·⁄„·Ì… „— »ÿ… »Õ—ﬂ«  ›Ê« Ì— «·⁄„·«¡"
                            Else
                            MsgBox "Can Not Edit .This movement is related to customer billing movements"
                            End If
                Exit Sub
                End If
 End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
           ' VSFlexGrid3.Rows = VSFlexGrid3.Rows + 1
            VSFlexGrid3.Enabled = True
            VSFlexGrid1.rows = VSFlexGrid1.rows + 1
            VSFlexGrid1.Enabled = True
              VSFlexGrid2.rows = VSFlexGrid2.rows + 1
            VSFlexGrid2.Enabled = True
            CuurentLogdata

        Case 2
           'khaled If SystemOptions.TripwithorderOnly = True And val(TxtBasedNo.Text) = 0 Then
            'khaled    If SystemOptions.UserInterface = ArabicInterface Then
            'khaled        MsgBox "·«»œ „‰ «Œ Ì«— «„—  Õ„Ì· ·«‰Â «·“«„Ì"
            'khaled     Else
            'khaled        MsgBox "ıPlease Enter Uploading Order"
            'khaled    End If
            'khaled    Exit Sub
            'khaled End If
           
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
            
         
            C1Tab1.CurrTab = 2
            
            
           
  
            If CBoBasedON.ListIndex > 0 And Trim(txt_ORDER_NO.text) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify NO For"
                Else
                    Msg = "Õœœ —ﬁ„ "
                End If

                Msg = Msg & "  " & CBoBasedON.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                txt_ORDER_NO.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
'            If Trim(Me.DcbAccount.BoundText) = "" Then
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    Msg = "Specify Account"
'                Else
'                    Msg = "Ì—ÃÏ «Œ Ì«— ÃÂ… «·’—›"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DcbAccount.SetFocus
'                Sendkeys "{F4}"
'                 Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
            
            my_branch = Me.Dcbranch.BoundText

            DcboBox_Change
            DcboBankName_Change
          '  dcEmp_Change
          '  DBCboClientName_Change
               Dim TxtNoteSerial1str As String

    If txtNoteSerial1.text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 74, 74)
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
    If ChCarType(1).value = True Then
    If val(DBCboClientName2.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê—œ"
    Else
    MsgBox "Please select supplier"
    End If
    Exit Sub
    End If
    End If
    
    If TxtLeaderName.text = "" And val(DCEmp.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·”«∆ﬁ"
    Else
    MsgBox "Please select driver"
    End If
    DCEmp.SetFocus
    Exit Sub
    End If
    Dim i As Integer
   
    If SystemOptions.TripDateInsertDefulat = True Then
   
    If VSFlexGrid3.rows = 1 Then
    DcbTypeTransport.BoundText = 1
    CardNO.text = 1
    QtyDownload.text = 1
    CardNO2.text = 1
    QtyDischarge.text = 1
    BillDate.value = XPDtbTrans.value
    DcboItems.BoundText = 1
    FillGridTypeTrans
    ReLineGrid

    End If
   
    End If
    If VSFlexGrid3.rows = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·„ Ì „ «œŒ«· ›Ê« Ì—  Õ„Ì· «Ê  ›—Ì€" & CHR(13)
    Msg = Msg & "Â·  —Ìœ «·Õ›Ÿ"
    Else
    Msg = "There are no data loading and unloading bills" & CHR(13)
    Msg = Msg & "Confirm Save"
    End If
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
     If IntRes = vbNo Then
    Exit Sub
    End If
    End If
    
    
    With VSFlexGrid3
    For i = 1 To .rows - 1
    If .TextMatrix(i, .ColIndex("BillDate")) <> "" And .TextMatrix(i, .ColIndex("CardNO")) <> "" Then
    If val(.TextMatrix(i, .ColIndex("QtyDownload"))) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·„ Ì „ «œŒ«· ﬂ„Ì… «· Õ„Ì· ›Ì «·”ÿ— —ﬁ„" & " " & i & CHR(13)
    Msg = Msg & "Â·  —Ìœ «·Õ›Ÿ"
    Else
    Msg = "The upload quantity has not been entered in line" & " " & i & CHR(13)
    Msg = Msg & "Confirm Save"
    End If
   IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
If IntRes = vbNo Then
Exit Sub
End If
    End If
    End If
    Next i
For i = 1 To .rows - 1
    If .TextMatrix(i, .ColIndex("BillDate")) <> "" And .TextMatrix(i, .ColIndex("CardNO2")) <> "" Then
    If val(.TextMatrix(i, .ColIndex("QtyDischarge"))) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·„ Ì „ «œŒ«· ﬂ„Ì… «· ›—Ì€ ›Ì «·”ÿ— —ﬁ„" & " " & i & CHR(13)
    Msg = Msg & "Â·  —Ìœ «·Õ›Ÿ"
    Else
    Msg = "The download quantity has not been entered in line" & " " & i & CHR(13)
    Msg = Msg & "Confirm Save"
    End If
   IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
If IntRes = vbNo Then
Exit Sub
End If
    End If
    End If
    Next i
    End With


            SaveData
        Case 3
            Undo
        Case 4
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
    If CheckAllocation() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ Õ–› Â–Â «·⁄„·Ì…. „— »ÿ… »Õ—ﬂ«  ›Ê« Ì— «·⁄„·«¡"
    Else
    MsgBox "Can Not delete.This movement is related to customer billing movements"
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

             FrmSearchinvestment.inde = 32
             FrmSearchinvestment.show vbModal

        Case 6
            Unload Me
        Case 7
            ViewDataList
        Case 8
            print_report (TxtSerial.text)
        Case 9
            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.text
        Case 10
            ShowGL_cc TxtSerial.text, , 200
       Case 11
        TxtSerial1.text = ""
        txtNoteSerial1.text = ""
        TxtModFlg.text = "N"
        VSFlexGrid3.rows = VSFlexGrid3.rows + 1
        VSFlexGrid3.Enabled = True
        Fg_Journal.rows = Fg_Journal.rows + 1
        Fg_Journal.Enabled = True
        VSFlexGrid2.rows = VSFlexGrid2.rows + 1
        VSFlexGrid2.Enabled = True
        FGOrders.rows = FGOrders.rows + 1
        FGOrders.Enabled = True
        Cmd(1).Enabled = True
        updaterowdate
        
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

    MySQL = "Select * From notes  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

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
    xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

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

  
    MySQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
    MySQL = MySQL & "     dbo.ACCOUNTS.Account_Name, dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteSerial, TblCountriesGovernments_2.GovernmentName AS FromCity,"
    MySQL = MySQL & "     TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.notes_all.NoteType, dbo.notes_all.CityFromId, dbo.notes_all.CityToId, dbo.notes_all.Location,"
    MySQL = MySQL & "    dbo.notes_all.CarId, dbo.notes_all.TravelPrice, dbo.notes_all.DriverValue, dbo.notes_all.DriverEra, dbo.notes_all.totalExpenses AS Desil, dbo.TblCarsData.BoardNO,"
    MySQL = MySQL & "    dbo.TblCarsData.Model, dbo.TblCarsData.Name, dbo.TblCarsData.LicenseNO, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.notes_all.Distance,"
    MySQL = MySQL & "    dbo.notes_all.KmPrice, dbo.notes_all.DriverId, dbo.notes_all.DriverPercentage, dbo.notes_all.startDate, dbo.notes_all.EndDate, dbo.notes_all.StartTime,"
    MySQL = MySQL & "    dbo.notes_all.EndTime, dbo.notes_all.KMCounterBeforeStart, dbo.notes_all.KMCounterAtEnd, dbo.notes_all.NoteCashingType, dbo.TblCustemers.CusName,"
    MySQL = MySQL & "    dbo.TblCustemers.CusNamee, dbo.notes_all.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ExpensesType.ID,"
    MySQL = MySQL & "    dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.notes_all.general_des"
    MySQL = MySQL & "    FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    MySQL = MySQL & "    dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
    MySQL = MySQL & "    dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID Left OUTER JOIN"
    MySQL = MySQL & "    dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code RIGHT OUTER JOIN"
    MySQL = MySQL & "    dbo.notes_all Left outer JOIN"
    MySQL = MySQL & "    dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id Left outer JOIN"
    MySQL = MySQL & "    dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID ON"
    MySQL = MySQL & "    dbo.Notes.notes_all = dbo.notes_all.NoteID LEFT OUTER JOIN"
    MySQL = MySQL & "    dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "    dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_Code LEFT OUTER JOIN"
    MySQL = MySQL & "    dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID"
    MySQL = MySQL & "    WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.notes_all.NoteID = " & val(XPTxtID.text) & ")"

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Transporter\TripData.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Transporter\TripData.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        'StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        'StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        'StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        'StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
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

Private Sub CmdDelete_Click()
If Me.TxtModFlg.text <> "R" Then
RemoveGridRow
End If
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Private Sub CmdRemove_Click()
TxtSerial1.text = txtNoteSerial1.text
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
    
    If Fg_Journal.rows > 1 Then
        If Fg_Journal.rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else
            If Me.Fg_Journal.rows > 1 Then
                If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                    With Me.Fg_Journal
                        If Me.TxtModFlg <> "E" Then Exit Sub
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                        LogTextA = "  Õ–› «·„’—Ê›   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                        LogTexte = "  Delete  Expensen   " & .cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                        AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
                    End With
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                End If
            End If
        End If
    End If
            
    With Fg_Journal
        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
    End With

End Sub

Private Sub DBCboClientName_Change()

    If DBCboClientName.BoundText = "" Then Exit Sub

    If CboPaymentType.ListIndex <> 1 Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        RetriveClinCounr
    End If
    
    'Text2.text = Me.DCVendor.BoundText
End Sub
Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 100
        FrmCustemerSearch.show vbModal

    End If
End Sub

Private Sub DBCboClientName2_Change()
DBCboClientName2_Click (0)
End Sub

Private Sub DBCboClientName2_Click(Area As Integer)
    Dim fullcode As String
    GetCustomersDetail val(DBCboClientName2.BoundText), , fullcode, 2
    TxtSearchCode.text = fullcode
End Sub

Private Sub DBCboClientName2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 101
        FrmCompanySearch.show vbModal
    End If
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
    If DcbAccount.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide2.BoundText = DcbAccount.BoundText
     End If
End Sub

Private Sub DcbBasedOn_Change()
    
    FGOrders.Clear flexClearScrollable, flexClearEverything
    FGOrders.rows = 1

    If val(DcbBasedOn.ListIndex) = 1 Then
        lbl(53).Visible = True
        TxtBasedNo.Visible = True
    ElseIf val(DcbBasedOn.ListIndex) = 3 Then
        lbl(53).Visible = True
        TxtBasedNo.Visible = True
    ElseIf val(DcbBasedOn.ListIndex) = 2 Then
        If Me.TxtModFlg.text <> "R" Then
            TxtBasedNo.text = ""
            TxtBasedNo.Visible = False
            lbl(53).Visible = False
            RetriveMultyOrders
        End If
    Else
        TxtBasedNo.text = ""
        TxtBasedNo.Visible = False
        lbl(53).Visible = False
    End If
    ReLineGrid
End Sub

Private Sub DcbBasedOn_Click()
DcbBasedOn_Change
End Sub

Private Sub DcbCar2_Change()
DcbCar2_Click (0)
End Sub

Private Sub DcbCar2_Click(Area As Integer)
 Dim Dcombos As New ClsDataCombos
Dcombos.GetBartCarByVonder DcbSupplem2, val(DcbCar2.BoundText)
DBCboClientName2.BoundText = GetCusIDByCarID(val(DcbCar2.BoundText))
End Sub

Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        'Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
        
        If CboPaymentType.ListIndex = 2 Or CboPaymentType.ListIndex = 3 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub
Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub DcboItems_Change()
DcboItems_Click (0)
Dim sql As String
Dim rsDummy  As New ADODB.Recordset
Dim Dcombos As New ClsDataCombos

    
    Dcombos.GetItemsUnitsDetai Me.cmbUnitName, val(DcboItems.BoundText)
    
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then



    sql = "   SELECT        TblClientTransContr.ID, TblClientTransContr.CompID, TblClientTransContr.LockedID, TblClientTransContr.UserID, TblClientTransContr.FromDate, TblClientTransContr.Todate,"
    sql = sql & "                       TblClientTransContr.Remarks, TblClientTransContr.Typed,  TblVendorCars.BoardNo,  TblClientTransContrDet.ClintTransID,"
    sql = sql & "                          TblClientTransContrDet.VehicleType , TblClientTransContrDet.Price , TblClientTransContrDet.Remarks AS Expr7,  TblClientTransContrDet.FromPrice,"
    sql = sql & "                          TblClientTransContrDet.ToPrice, TblClientTransContrDet.FromCityID, TblClientTransContrDet.ToCityID, TblCountriesGovernments_1.GovernmentName as ToCityName, TblCountriesGovernments.GovernmentName AS FromCity,"
    sql = sql & "                          TblCustemers.CusName , TblCustemers.CusID, TblCustemers.CusNamee, TblVendorCars.nBoardNo, TblVendorCars.ChasisNo, TblVendorCars.BrandID, TblVendorCars.ModelID,"
    sql = sql & "                   TblItems.ItemName,TblItems.ItemNamee,TblItems.ItemCode,TblItems.ItemID,TblUnites.UnitID,TblUnites.UnitName,TblUnites.UnitNamee,TblClientTransContrDet.VehicleType"
    sql = sql & " FROM            TblCountriesGovernments AS TblCountriesGovernments_1 RIGHT OUTER JOIN"
    sql = sql & "                          TblClientTransContrDet ON TblCountriesGovernments_1.GovernmentID = TblClientTransContrDet.ToCityID LEFT OUTER JOIN"
    sql = sql & "                          TblCountriesGovernments ON TblClientTransContrDet.FromCityID = TblCountriesGovernments.GovernmentID RIGHT OUTER JOIN"
    sql = sql & "                          TblClientTransContr LEFT OUTER JOIN"
    sql = sql & "                          TblVendorCars ON TblClientTransContr.VehicleType = TblVendorCars.ID LEFT OUTER JOIN"
    sql = sql & "                          TblCustemers ON TblClientTransContr.CompID = TblCustemers.CusID ON TblClientTransContrDet.ClintTransID = TblClientTransContr.ID"
    sql = sql & "                  LEFT OUTER JOIN TblItems  On TblItems.ItemID =TblClientTransContrDet.ItemID"
    sql = sql & "                  LEFT OUTER JOIN TblUnites  On TblUnites.UnitID =TblClientTransContrDet.UnitID"
   ' sql = sql & " Where (1 = 1) "
    'and TblClientTransContr.ID=" & ID
        If Trim(VehicleType.text) = "" Then
    sql = sql & " Where  (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
Else
    sql = sql & " Where (dbo.TblClientTransContr.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
End If

sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""

'    sql = sql & " and TblItems.ItemID = " & val(DcboItems.BoundText)
'    sql = sql & " and TblUnites.UnitID = " & val(cmbUnitName.BoundText)
'    sql = sql & " and (dbo.TblClientTransContrDet.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
'    sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
'    sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""
    

    sql = sql & " and TblItems.ItemID = " & val(DcboItems.BoundText)
    sql = sql & " and TblUnites.UnitID = " & val(cmbUnitName.BoundText)
    
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open sql, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        TxtPrice = rsDummy!Price & ""
        TXTTravelPrice = rsDummy!Price & ""
    End If
    
    
End If

End Sub

Private Sub DcboItems_Click(Area As Integer)
  Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
End Sub

Private Sub DcboItems_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 100
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.text = ""
    TxtSerial1.text = ""
    txtNoteSerial1.text = ""
End Sub

Function GetDriverInformation(ID As Integer)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TblCarsData"
        sql = sql & " Where (id = " & ID & ") "

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            DCEmp.BoundText = IIf(IsNull(rs("Emp_id").value), 0, rs("Emp_id").value)
        Else
            DCEmp = 0
        End If

    End If

End Function
Function GetTripInformations()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    
        Dim sql As String
        Dim rs As New ADODB.Recordset
        Dim rsDummy As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TBLCitiesDistance"
        sql = sql & " Where (CityFromId = " & val(DcCityFromId.BoundText) & ") And (CitytoId=" & val(DcCityToId.BoundText) & ")"
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            TxtDistance = IIf(IsNull(rs("Distance").value), 0, rs("Distance").value)
            TxtKmPrice = IIf(IsNull(rs("Desil").value), 0, rs("Desil").value)
            
            TXTTravelPrice = IIf(IsNull(rs("TravelPrice").value), 0, rs("TravelPrice").value)
            TxtDriverPercentage = IIf(IsNull(rs("DriverPercentage").value), 0, rs("DriverPercentage").value)
            txtDriverValue = IIf(IsNull(rs("DriverValue").value), 0, rs("DriverValue").value)

             If DCCar.text <> "" Then
                sql = "Select IsNull(IsUsed,0) IsUsed From TblCarsData Where Id  =  " & val(DCCar.BoundText)
                rsDummy.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Not rsDummy.EOF Then
                     If Not (IsNull(rsDummy("IsUsed").value)) Then
                        If rsDummy("IsUsed").value Then
                            TXTTravelPrice = IIf(IsNull(rs("TravelPriceUsed").value), 0, rs("TravelPriceUsed").value)
                            TxtDriverPercentage = IIf(IsNull(rs("DriverPercentageUsed").value), 0, rs("DriverPercentageUsed").value)
                            txtDriverValue = IIf(IsNull(rs("DriverValueUsed").value), 0, rs("DriverValueUsed").value)
                        End If
                    End If
                
                End If
             End If
        Else
            TxtDistance = 0
            TxtKmPrice = 0
            TXTTravelPrice = 0
            TxtDriverPercentage = 0
            txtDriverValue = 0
       
        End If

    End If

End Function

Private Sub dcCar_Change()
    GetDriverInformation (val(DCCar.BoundText))
    GetTripInformations
     Dim Dcombos As New ClsDataCombos
Dcombos.GetPartCar DcbSupplem, val(DCCar.BoundText)
End Sub
Private Sub dcCar_Click(Area As Integer)
    dcCar_Change
End Sub

Private Sub Dccar_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "TravelTrans"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcCityFromId_Change()
    GetTripInformations
    RetriveClinCounr
      Me.txtCityFromCode.text = GetGovernmentCode(val(Me.DcCityFromId.BoundText))
End Sub
Private Sub DcCityFromId_Click(Area As Integer)
    GetTripInformations
    RetriveClinCounr
End Sub
Private Sub DcCityToId_Change()
    GetTripInformations
    RetriveClinCounr
    Me.txtCityToCode.text = GetGovernmentCode(val(Me.DcCityToId.BoundText))
End Sub
Private Sub DcCityToId_Click(Area As Integer)
    GetTripInformations
    RetriveClinCounr
End Sub
Private Sub DcCostCenter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 4
    End If
End Sub
Private Sub dcEmp_Change()
Dim empSalaryAccount As String
    If DCEmp.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
         DcbAccount.BoundText = ModAccounts.GetDriverAccountCode(val(Me.DCEmp.BoundText))
         
        DcboCreditSide2.BoundText = ModAccounts.GetDriverAccountCode(val(Me.DCEmp.BoundText))
   
If SystemOptions.showEmployeeAccountIntrip = True Then


DcbAccount.BoundText = GetemployeeAccountCode(val(Me.DCEmp.BoundText))
DcboCreditSide2.BoundText = GetemployeeAccountCode(val(Me.DCEmp.BoundText))

End If

        Dim Balance As String
        Dim balanceString As String
        Dim balancetype As Integer
        WriteCustomerBalPublic Me.DcboCreditSide2.BoundText, Balance, balanceString, balancetype

        If Me.TxtModFlg.text = "N" Then
            If balancetype = 0 Then
                txtDriverEra = val(Balance)
            Else
                txtDriverEra = 0
            End If

        Else
            If balancetype = 0 Then
                txtDriverEra = val(Balance) '+ val(XPTxtVal.text)
            Else
                txtDriverEra = val(Balance)  ' - val(XPTxtVal.text))
            End If
        End If

        LblLink.Caption = balanceString
    End If
End Sub
Private Sub Dcemp_Click(Area As Integer)
    dcEmp_Change
End Sub
Private Sub dcproject_Change()

    If dcproject.text = "" Then
        VSFlexGrid1.Visible = False
        Me.Fg_Journal.Visible = True
    End If
 
End Sub

Private Sub dcproject_Click(Area As Integer)

    If SystemOptions.gldetails_or_gl_general = 0 Then 'Õ”«»«  «·„‘—Ê⁄
        VSFlexGrid1.Visible = True
        Me.Fg_Journal.Visible = False
    Else
        VSFlexGrid1.Visible = False
        Me.Fg_Journal.Visible = True
    End If

End Sub

Function CheckAllExpensesDistributed() As Boolean
    CheckAllExpensesDistributed = False
    Dim i As Integer
    Dim zeroExist As Boolean
    Dim oneexist As Boolean

    With Fg_Journal

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("Destribute")) = "0" Then
                zeroExist = True
            End If
        
            If .TextMatrix(i, .ColIndex("Destribute")) = "1" Then
                oneexist = True
            End If
        
            If zeroExist = True And oneexist = True Then
                CheckAllExpensesDistributed = False
                Exit Function
            End If
        Next i
    End With

    CheckAllExpensesDistributed = True
End Function
Function FillDestributionsToAll() As Boolean

    GridEstimatedCost.Clear flexClearScrollable, flexClearEverything
    GridEstimatedCost.rows = 1
    
    Dim Msg As String

    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“⁄Â Ê«Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ  " & CHR(13)
        Else
            Msg = " This Expenses Voucher  Have  Destribute and not  Destribute Expenses " & CHR(13)
            Msg = Msg + "can't Save"
        End If
                                 
        FillDestributionsToAll = False
        Exit Function
            
    End If
 
    Dim i As Integer
    GridEstimatedCost.Clear flexClearScrollable, flexClearEverything
    GridEstimatedCost.rows = 1
          
    With Fg_Journal
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                FillDestributions .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("AccountName")), val(.TextMatrix(i, .ColIndex("value"))) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
            End If
        Next i
    End With
End Function
 
Public Function FillDestributions(AcountCode As String, AcountName As String, value As Double)
 
    Dim StrSQL  As String
    
    StrSQL = "SELECT     dbo.TblAccountsDestributions.AccountMaster, dbo.TblAccountsDestributionsDetails.ACode, dbo.TblAccountsDestributionsDetails.Percentage, "
    StrSQL = StrSQL + "  dbo.TblAccountsDestributions.DistType , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL + " FROM         dbo.TblAccountsDestributions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblAccountsDestributionsDetails ON"
    StrSQL = StrSQL + " dbo.TblAccountsDestributions.TblAccountsDestributionsid = dbo.TblAccountsDestributionsDetails.TblAccountsDestributionsid INNER JOIN"
    StrSQL = StrSQL + "  dbo.TblBranchesData ON dbo.TblAccountsDestributionsDetails.ACode = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL + " WHERE     (dbo.TblAccountsDestributions.DistType IS NULL) AND (dbo.TblAccountsDestributions.AccountMaster = N'" & AcountCode & "')"
     
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
 
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
 
        row_count = GridEstimatedCost.rows
    
        If GridEstimatedCost.TextMatrix(row_count - 1, GridEstimatedCost.ColIndex("AcountCode")) = "" Then
            row_count = row_count - 1
        End If
     
        GridEstimatedCost.rows = RsDetails.RecordCount + row_count

        For Num = row_count To GridEstimatedCost.rows - 1 'RsDetails.RecordCount
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Ser")) = Num
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AcountCode")) = AcountCode
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AcountName")) = AcountName
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchId")) = IIf(IsNull(RsDetails("Acode")), "", (RsDetails("Acode").value))
            If SystemOptions.UserInterface = ArabicInterface Then
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
            Else
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("BranchName")) = IIf(IsNull(RsDetails("branch_namee")), "", (RsDetails("branch_namee").value))
            End If
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Percentage")) = IIf(IsNull(RsDetails("Percentage")), 0, (RsDetails("Percentage").value))
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("value")) = value
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Netvalue")) = Round(value * GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Percentage")) / 100, 2)
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If GridEstimatedCost.Rows > 10 Then
            '     If Num = 8 Then GridEstimatedCost.Refresh
            ' End If
        Next Num
    End If
End Function
Function addFixedExpenses()
If SystemOptions.TripnotUploadExpenses = True Then 'Õ«·Â «Œ Ì«— «Ê»‘‰ ⁄œ„  Õ„Ì· «·„’—Ê›«  «·Ì«
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 4
Exit Function

End If
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    Dim AccountName As String
 
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    On Error Resume Next
    Dim StrAccountCode As String
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 4
    With Me.Fg_Journal
        If TxtKmPrice <> 0 Then
            .Row = 1
            .TextMatrix(.Row, .ColIndex("AccountCode")) = Account_Code_dynamic1
            .TextMatrix(.Row, .ColIndex("Destribute")) = 0
            StrAccountCode = .TextMatrix(.Row, .ColIndex("AccountCode"))
            If CheckAccountHaveDestributions(StrAccountCode) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " Â–« «·„’—Ê› ·Â ŒÿÂ  Ê“Ì⁄  ⁄·Ï «·›—Ê⁄ Â·  —Ìœ «· Ê“Ì⁄  " & CHR(13)
                    Msg = Msg + "‰⁄„ «„ ·« "
                Else
                    Msg = " This Expenses Have Destribution Plan Do you want  Destribute  " & CHR(13)
                    Msg = Msg + "Yes Or No"
                End If
                                 
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    .TextMatrix(.Row, .ColIndex("Destribute")) = 1
                Else
                    .TextMatrix(.Row, .ColIndex("Destribute")) = 0
                End If
            End If
            FillDestributionsToAll
            .TextMatrix(.Row, .ColIndex("ExpensesID")) = get_Expenses_id(Account_Code_dynamic1)
            .TextMatrix(.Row, .ColIndex("AccountName")) = get_Expenses_id(Account_Code_dynamic1, AccountName)
            .TextMatrix(.Row, .ColIndex("AccountName")) = AccountName
            .TextMatrix(.Row, .ColIndex("LineNo1")) = setfoxy_Line
            .TextMatrix(.Row, .ColIndex("Value")) = Me.TxtKmPrice
            .TextMatrix(.Row, .ColIndex("Order_No")) = txt_ORDER_NO.text
            
            .TextMatrix(.Row, .ColIndex("des")) = "„’—Ê› œÌ“·"
        End If
        If txtDriverValue <> 0 Then
            .Row = 2
            .TextMatrix(.Row, .ColIndex("AccountCode")) = Account_Code_dynamic2
            .TextMatrix(.Row, .ColIndex("Destribute")) = 0
            StrAccountCode = .TextMatrix(.Row, .ColIndex("AccountCode"))
            If CheckAccountHaveDestributions(StrAccountCode) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " Â–« «·„’—Ê› ·Â ŒÿÂ  Ê“Ì⁄  ⁄·Ï «·›—Ê⁄ Â·  —Ìœ «· Ê“Ì⁄  " & CHR(13)
                    Msg = Msg + "‰⁄„ «„ ·« "
                Else
                    Msg = " This Expenses Have Destribution Plan Do you want  Destribute  " & CHR(13)
                    Msg = Msg + "Yes Or No"
                End If
                                 
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    .TextMatrix(.Row, .ColIndex("Destribute")) = 1
                Else
                    .TextMatrix(.Row, .ColIndex("Destribute")) = 0
                End If
            End If
 
            FillDestributionsToAll
            .TextMatrix(.Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
            .TextMatrix(.Row, .ColIndex("AccountName")) = get_Expenses_id(StrAccountCode, AccountName)
            .TextMatrix(.Row, .ColIndex("AccountName")) = AccountName
            .TextMatrix(.Row, .ColIndex("LineNo1")) = setfoxy_Line
            .TextMatrix(.Row, .ColIndex("Value")) = Me.txtDriverValue
            .TextMatrix(.Row, .ColIndex("Order_No")) = txt_ORDER_NO.text
            .TextMatrix(.Row, .ColIndex("des")) = "  „ﬂ«›√… ”«∆ﬁÌ‰"
        End If
        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
    End With
 
    'Me.TxtValue.text = ""
    'DcboBox.BoundText = ""
    ReLineGrid

End Function
Sub RetriveClinCounr()
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


'
'sql = "   SELECT        TblClientTransContr.ID, TblClientTransContr.CompID, TblClientTransContr.LockedID, TblClientTransContr.UserID, TblClientTransContr.FromDate, TblClientTransContr.Todate,"
'sql = sql & "                       TblClientTransContr.Remarks, TblClientTransContr.Typed,  TblVendorCars.BoardNo,  TblClientTransContrDet.ClintTransID,"
'sql = sql & "                          TblClientTransContrDet.VehicleType , TblClientTransContrDet.Price , TblClientTransContrDet.Remarks AS Expr7,  TblClientTransContrDet.FromPrice,"
'sql = sql & "                          TblClientTransContrDet.ToPrice, TblClientTransContrDet.FromCityID, TblClientTransContrDet.ToCityID, TblCountriesGovernments_1.GovernmentName as ToCityName, TblCountriesGovernments.GovernmentName AS FromCity,"
'sql = sql & "                          TblCustemers.CusName , TblCustemers.CusID, TblCustemers.CusNamee, TblVendorCars.nBoardNo, TblVendorCars.ChasisNo, TblVendorCars.BrandID, TblVendorCars.ModelID"
'sql = sql & " FROM            TblCountriesGovernments AS TblCountriesGovernments_1 RIGHT OUTER JOIN"
'sql = sql & "                          TblClientTransContrDet ON TblCountriesGovernments_1.GovernmentID = TblClientTransContrDet.ToCityID LEFT OUTER JOIN"
'sql = sql & "                          TblCountriesGovernments ON TblClientTransContrDet.FromCityID = TblCountriesGovernments.GovernmentID RIGHT OUTER JOIN"
'sql = sql & "                          TblClientTransContr LEFT OUTER JOIN"
'sql = sql & "                          TblVendorCars ON TblClientTransContr.VehicleType = TblVendorCars.ID LEFT OUTER JOIN"
'sql = sql & "                          TblCustemers ON TblClientTransContr.CompID = TblCustemers.CusID ON TblClientTransContrDet.ClintTransID = TblClientTransContr.ID"
'sql = sql & " Where (dbo.TblClientTransContrDet.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
'sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
'sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""
'
'
'
'
'
''
''
''sql = " SELECT     dbo.TblClientTransContrDet.Price, dbo.TblClientTransContrDet.Typed"
''sql = sql & " FROM         dbo.TblClientTransContr LEFT OUTER JOIN"
''sql = sql & "                      dbo.TblClientTransContrDet ON dbo.TblClientTransContr.ID = dbo.TblClientTransContrDet.ClintTransID"
''sql = sql & " Where (dbo.TblClientTransContrDet.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
''sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
''sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""
'rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If rs2.RecordCount > 0 Then
'rs2.MoveFirst
'TxtPrice.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
'
'    If rs2.RecordCount > 0 Then
'        If Not IsNull(rs2("Typed").value) Then
'            If val(rs2("Typed").value) = 1 Then
'                RdTyped(1).value = True
'            ElseIf val(rs2("Typed").value) = 2 Then
'                RdTyped(2).value = True
'            ElseIf val(rs2("Typed").value) = 3 Then
'                RdTyped(3).value = True
'            ElseIf val(rs2("Typed").value) = 4 Then
'                RdTyped(4).value = True
'            Else
'                RdTyped(0).value = True
'            End If
'        Else
'            RdTyped(0).value = True
'        End If
'    End If
'End If
'End If

    
sql = "   SELECT        TblClientTransContr.ID,TblUnites.UnitName,TblUnites.UnitNamee,TblUnites.UnitId, TblClientTransContr.CompID, TblClientTransContr.LockedID, TblClientTransContr.UserID, TblClientTransContr.FromDate, TblClientTransContr.Todate,"
sql = sql & "                       TblClientTransContr.Remarks, TblClientTransContr.Typed,  TblVendorCars.BoardNo,  TblClientTransContrDet.ClintTransID,"
sql = sql & "                          TblClientTransContrDet.VehicleType , TblClientTransContrDet.Price , TblClientTransContrDet.Remarks AS Expr7,  TblClientTransContrDet.FromPrice,"
sql = sql & "                          TblClientTransContrDet.ToPrice, TblClientTransContrDet.FromCityID, TblClientTransContrDet.ToCityID, TblCountriesGovernments_1.GovernmentName as ToCity, TblCountriesGovernments.GovernmentName AS FromCity,"
sql = sql & "                          TblCustemers.CusName , TblCustemers.CusID, TblCustemers.CusNamee, TblVendorCars.nBoardNo, TblVendorCars.ChasisNo, TblVendorCars.BrandID, TblVendorCars.ModelID,  TblClientTransContr.VehicleType,"
sql = sql & "                   TblItems.ItemName, QtyDownload = 1,TotalValue =TblClientTransContrDet.Price,  TblItems.ItemNamee,TblItems.ItemCode,TblItems.ItemID,TblUnites.UnitID,TblUnites.UnitName,TblUnites.UnitNamee"
sql = sql & " FROM            TblCountriesGovernments AS TblCountriesGovernments_1 RIGHT OUTER JOIN"
sql = sql & "                          TblClientTransContrDet ON TblCountriesGovernments_1.GovernmentID = TblClientTransContrDet.ToCityID LEFT OUTER JOIN"
sql = sql & "                          TblCountriesGovernments ON TblClientTransContrDet.FromCityID = TblCountriesGovernments.GovernmentID RIGHT OUTER JOIN"
sql = sql & "                          TblClientTransContr LEFT OUTER JOIN"
sql = sql & "                          TblVendorCars ON TblClientTransContr.VehicleType = TblVendorCars.ID LEFT OUTER JOIN"
sql = sql & "                          TblCustemers ON TblClientTransContr.CompID = TblCustemers.CusID ON TblClientTransContrDet.ClintTransID = TblClientTransContr.ID"
sql = sql & "                  LEFT OUTER JOIN TblItems  On TblItems.ItemID =TblClientTransContrDet.ItemID"
sql = sql & "                  LEFT OUTER JOIN TblUnites  On TblUnites.UnitID =TblClientTransContrDet.UnitID"
'sql = sql & " Where (1 = 1) and TblClientTransContr.ID=" & ID
 sql = sql & " Where  (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
If Trim(VehicleType.text) = "" Then
  '  sql = sql & " Where  (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
Else

  '  sql = sql & " Where (dbo.TblClientTransContr.VehicleType = " & val(VehicleType.BoundText) & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
End If
sql = sql & " and  (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null) And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName.BoundText) & ")"
sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""
'
'


    
    'sql = "select * from TblOrderUpload where "
    If Me.TxtModFlg.text = "N" Then
      '  sql = sql & " and IsTravel is null "
    ElseIf Me.TxtModFlg.text = "E" Then
      '  sql = sql & " and IsTravel is null or ID in (Select BasedNo from notes_all  where NoteID=" & val(XPTxtID.text) & " ) "
    End If
'sql = sql & " ORDER BY TblClientTransContr.ID"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs2.RecordCount > 0 Then
        If Not IsNull(rs2("Typed").value) Then
            If val(rs2("Typed").value) = 1 Then
                RdTyped(1).value = True
            ElseIf val(rs2("Typed").value) = 2 Then
                RdTyped(2).value = True
            ElseIf val(rs2("Typed").value) = 3 Then
                RdTyped(3).value = True
            ElseIf val(rs2("Typed").value) = 4 Then
                RdTyped(4).value = True
            Else
                RdTyped(0).value = True
            End If
        Else
            RdTyped(0).value = True
        End If
        RdTyped_Click val(rs2("Typed").value)
        'TxtRent.text = IIf(IsNull(rs2("TxtRent").value), 0, rs2("TxtRent").value)
        'TxtPartPrice.text = IIf(IsNull(rs2("PartPrice").value), 0, rs2("PartPrice").value)
        TxtPrice1.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        TxtPrice.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        txtto.text = IIf(IsNull(rs2("Remarks").value), 0, rs2("Remarks").value)
        TxtTotal1.text = TxtPrice ' IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
        DBCboClientName2.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
        DcCityFromId.BoundText = IIf(IsNull(rs2("FromCityID").value), 0, rs2("FromCityID").value)
        VehicleType.BoundText = IIf(IsNull(rs2("VehicleType").value), 0, rs2("VehicleType").value)
        DcCityFromId.BoundText = IIf(IsNull(rs2("FromCityID").value), 0, rs2("FromCityID").value)
        
        
        
'                    DcboItems.BoundText = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
'            cmbUnitName.BoundText = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
       
        TXTTravelPrice.text = TxtPrice.text
        DcCityToId.BoundText = IIf(IsNull(rs2("ToCityID").value), 0, rs2("ToCityID").value)
        TxtPrice = rs2!Price & ""
        'DCCar.BoundText = IIf(IsNull(rs2("CarID").value), 0, rs2("CarID").value)
        'DcbCar2.BoundText = IIf(IsNull(rs2("CarID2").value), 0, rs2("CarID2").value)
        'DcbSupplem.BoundText = IIf(IsNull(rs2("SupplemID").value), 0, rs2("SupplemID").value)
        'DcbSupplem2.BoundText = IIf(IsNull(rs2("SupplemID2").value), 0, rs2("SupplemID2").value)
        'TxtLeaderName.text = IIf(IsNull(rs2("LeaderName").value), "", rs2("LeaderName").value)
        TxtIDNo.text = IIf(IsNull(rs2("ID").value), "", rs2("ID").value)
        'TxtOrderNo.text = IIf(IsNull(rs2("OrderNo").value), "", rs2("OrderNo").value)
        'TxtTypGoods.text = IIf(IsNull(rs2("TypGoods").value), "", rs2("TypGoods").value)
        DBCboClientName.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
        'DCEmp.BoundText = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
    Else
'        DCEmp.BoundText = ""
'        DBCboClientName.BoundText = ""
'        TxtOrderNo.text = ""
'        TxtLeaderName.text = ""
'        TxtIDNo.text = ""
'        DBCboClientName2.BoundText = 0
'        DcbSupplem2.BoundText = 0
'        DcbSupplem.BoundText = 0
'        DcbCar2.BoundText = 0
'        DCCar.BoundText = 0
'        DcCityToId.BoundText = 0
'        DcCityFromId.BoundText = 0
'        TxtPartPrice.text = ""
'        TxtPrice1.text = ""
'        TxtTotal1.text = ""
    End If
    
        Dim i As Long
    If DcboItems.text = "" Then
        If Not rs2.EOF Then
            DcboItems.BoundText = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            cmbUnitName.BoundText = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
            
                If Not RdTyped(3).value Or Not RdTyped(2).value Then
                    TxtPrice = 0
                '
                    For i = 1 To VSFlexGrid3.rows - 1
                        TxtPrice = val(TxtPrice) + (val(VSFlexGrid3.TextMatrix(i, VSFlexGrid3.ColIndex("Price"))) * val(VSFlexGrid3.TextMatrix(i, VSFlexGrid3.ColIndex("QtyDownload"))))
                    Next
                End If
        End If
    Else
        sql = sql & " and TblClientTransContrDet.ItemID =" & DcboItems.BoundText
       rs2.Close
       rs2.Open sql, Cn, adOpenStatic, adLockReadOnly
       If Not rs2.EOF Then
            TxtPrice1.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
            TxtPrice.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
            
            'TxtPrice = 0
            '
            If VSFlexGrid3.rows > 1 Then
                VSFlexGrid3.TextMatrix(1, VSFlexGrid3.ColIndex("Price")) = TxtPrice
                VSFlexGrid3.TextMatrix(1, VSFlexGrid3.ColIndex("TotalValue")) = (val(VSFlexGrid3.TextMatrix(1, VSFlexGrid3.ColIndex("Price"))) * val(VSFlexGrid3.TextMatrix(1, VSFlexGrid3.ColIndex("QtyDownload"))))
                TxtPrice = 0
                For i = 1 To VSFlexGrid3.rows - 1
                    
                    TxtPrice = val(TxtPrice) + (val(VSFlexGrid3.TextMatrix(i, VSFlexGrid3.ColIndex("Price"))) * val(VSFlexGrid3.TextMatrix(i, VSFlexGrid3.ColIndex("QtyDownload"))))
                    
                Next
            End If
        End If
         
        
        
    End If
   ' loadgrid sql, VSFlexGrid3, True
    'VSFlexGrid3.IsSubtotal(VSFlexGrid3.rows - 1) = True
    'TxtPrice = ""


    TXTTravelPrice = TxtPrice
    TXTTravelPrice = TxtPrice
    VSFlexGrid3.Visible = True
    'TxtPrice = VSFlexGrid3.Aggregate(flexSTSum, VSFlexGrid3.FixedRows, VSFlexGrid3.ColIndex("Price"), VSFlexGrid3.rows - 1, VSFlexGrid3.ColIndex("Price"))
    TxtQtyDownload = 1
    TxtQtyDischarge = 1
    If RdTyped(3).value = True Then
    
        RdTyped_Click 3
    ElseIf RdTyped(4).value = True Then
        RdTyped_Click 4
    End If
    ReLineGrid
End If
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sgl As String
TxtSerial1.text = txtNoteSerial1.text

    With Fg_Journal
        sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
        Cn.Execute sgl, , adExecuteNoRecords

        Select Case .ColKey(Col)
            Case "ExpensesID"
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
            Case "AccountName"
                '.TextMatrix(Row, .ColIndex("userid")) = user_id
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("Destribute")) = 0
                StrAccountCode = .TextMatrix(Row, .ColIndex("AccountCode"))
                If CheckAccountHaveDestributions(StrAccountCode) = True Then
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
 
                FillDestributionsToAll
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = txt_ORDER_NO.text
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value", "opr_fullcode"
        
                Dim project_id As Integer
                
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
    
                FillDestributionsToAll
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
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

        AddToLogFile CInt(user_id), 3, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select
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

          '  CboDes.Visible = False

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

       ' CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
       ' CboDes.Visible = True
       ' CboDes.ZOrder 0
       ' CboDes.SetFocus
 
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub
Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
     Sendkeys "{F4}"
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, Shift As Integer)

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
                     FrmExpensesSearch.RetrunType = 1
                End If
 
        End Select

    End With

End Sub

Public Sub Fg_Journal_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

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

            Case "AccountName"
                '      StrSQL = "select * from Expenses_accounts"
                             
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts order by Account_Name"
                Else
                    StrSQL = "select * from Expenses_accounts_eng order by Account_Nameeng"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Nameeng", "Account_Code")
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

Private Sub FGOrders_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

Private Sub FGOrders_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FGOrders
Select Case .ColKey(Col)
Case "Selct"
.ComboList = ""
Case "RentVAlue"
.ComboList = ""
Case "LineNo"
Cancel = True
Case "OrderNo"
Cancel = True
Case "Name"
Cancel = True
Case "TypedID"
Cancel = True

Case "CarName"
Cancel = True
Case "Part"
Cancel = True
Case "Price"
Cancel = True
Case "TypedID"
Cancel = True
Case "TypeDriver"
Cancel = True
Case "DriverName"
Cancel = True
End Select
End With
End Sub

Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos

    Dim StrSQL As String

    On Error GoTo ErrTrap

    'StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    'fill_combo Me.DcCostCenter, StrSQL
  
    ScreenNameArabic = " ”ÃÌ· »Ì«‰«  «·—Õ·«   "
    ScreenNameEnglish = "Trips Data "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 370

    Account_Code_dynamic1 = get_account_code_branch(69, my_branch)
    Account_Code_dynamic2 = get_account_code_branch(70, my_branch)
    Account_Code_dynamic3 = get_account_code_branch(71, my_branch)
                
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
    
    With FGOrders
         .ColComboList(.ColIndex("TypeDriver")) = "#1;œ«Œ·Ì |#2;Œ«—ÃÌ "
         .ColComboList(.ColIndex("TypedID")) = "#1;„„·Êﬂ… |#2;«Œ—Ï "
    End With
    
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetTypesTransport Me.DcbTypeTransport
    Dcombos.GetItemsNames Me.DcboItems
    
    Dcombos.GetItemsUnits Me.cmbUnitName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetTblCarsDataGroup VehicleType
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName2
    Dcombos.GetCars Me.DCCar
    Dcombos.GetHarbors Me.DcbHarbor
    Dcombos.GetShips Me.DcbShip
    
   If SystemOptions.IsTransferByCode Then
        Dcombos.getCountriesGovernments Me.DcCityFromId
        Dcombos.getCountriesGovernments Me.DcCityToId
   Else
        Dcombos.GetCitiesDistance Me.DcCityFromId, 0
        Dcombos.GetCitiesDistance Me.DcCityToId, 1
    End If
'    Dcombos.GetEmployees Me.DCEmP, , True
    
Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
   If SystemOptions.ShowDriverOnly = True Then
   str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   End If
    fill_combo DCEmp, str

    
    
    Dcombos.GetCarByVonder DcbCar2
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide2

    Dcombos.GetBranches Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ/ ⁄ÂœÂ"
        .AddItem "«Ã·"
        .AddItem " ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "‘Ìﬂ „Õ’·"
    End With
    
    With DcbBasedOn
        .Clear
        .AddItem "»·«"
        .AddItem "«„—  Õ„Ì·"
        .AddItem "«ﬂÀ— „‰ «„—  Õ„Ì·"
        .AddItem "« ›«ﬁÌ… ⁄„·«¡"
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

    StrSQL = " select expanses_account,Project_name from projects  where not(expanses_account is null)"
    fill_combo dcproject, StrSQL
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrSQL = "select ItemID,ItemName from tblitems  where GroupID in ( "
'    Else
'        StrSQL = "select ItemID,ItemNamee from tblitems  where GroupID in ( "
'    End If
'    StrSQL = StrSQL & " SELECT     GroupID "
'    StrSQL = StrSQL & " From dbo.Groups"
'    StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"
'
'   fill_combo DcboItems, StrSQL
                
    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=370"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
TxtModFlg_Change
                     If SystemOptions.TripRevenueAuto = True Then
           C1Tab1.CurrTab = 4
           Else
            C1Tab1.CurrTab = 2
            End If
            
            
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    'MsgBox ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    hide_logo = False
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 3

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

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, ByVal SpinningEnded As Boolean)

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
Private Sub CboDes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{F4}"
    End If
End Sub
Function VIEW_ATTACH()

    'On Error Resume Next''
 
    'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.show
    imaged.Label9.Caption = "„—›ﬁ«  «·—Õ·… —ﬁ„"
    imaged.Caption = "„—›ﬁ«  «·—Õ·…  "
    imaged.txtopeation_type = "„—›ﬁ«  «·—Õ·…"
    imaged.SUBJECT_NO = XPTxtID 'TxtEmp_Code.text
    imaged.Label6.Caption = "ﬂÊœ «·—Õ·…"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·—Õ·…' and subject_no='" & XPTxtID & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Function

Private Sub ISButton1_Click()
'    If DoPremis(Do_Attach, Me.Name, True) = False Then
'        Exit Sub
'    End If
'    VIEW_ATTACH
    
    
     '   On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
If vbcheck.value = vbUnchecked Then
ShowAttachments txtNoteSerial1 & "-" & XPTxtID.text, "0411201801"
Else
ShowAttachments txtNoteSerial1, "0411201801"
End If

End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.text <> "R" Then
UnPayedFlag
RemoveFgOrdersRow
End If
End Sub

Private Sub ISButton3_Click()
If Me.TxtModFlg.text <> "R" Then
RemoveTyptransRow
End If
End Sub

Private Sub ISButton4_Click()
    Dim IntRes As Integer
    Dim Msg As String
    
    If Me.TxtModFlg.text <> "R" Then
        If val(DcboItems.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
            Else
                MsgBox "Please Select Item"
            End If
            Exit Sub
        End If
        
        If CheckIploadQty(CardNO.text) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ ﬂ—  «· Õ„Ì· „ÊÃÊœ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â·  —Ìœ «·„Ê«’·…"
            Else
                Msg = "The card number already exists" & CHR(13)
                Msg = Msg & "Confirm Continue"
            End If
            IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
            If IntRes = vbNo Then
                Exit Sub
            End If
        End If

        If CheckIDownLoadQty(CardNO2.text) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ ﬂ—  «· ›—Ì€ „ÊÃÊœ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â·  —Ìœ «·„Ê«’·…"
            Else
                Msg = "The card number already exists" & CHR(13)
                Msg = Msg & "Confirm Continue"
            End If
            IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
            If IntRes = vbNo Then
                Exit Sub
            End If
        End If

        FillGridTypeTrans
        ReLineGrid
    End If
End Sub
Function CheckIDownLoadQty(Optional CardNO As String, Optional Row As Long) As Boolean
CheckIDownLoadQty = False
If CardNO = "" Then Exit Function
If CheckUplaodSerialInData(CardNO, 1) = True Then
CheckIDownLoadQty = True
Exit Function
Else
CheckIDownLoadQty = False
End If
Dim i As Integer
With VSFlexGrid3
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("CardNO2"))) = CardNO And i <> Row Then
CheckIDownLoadQty = True
Exit Function
End If
Next i
End With
End Function

Function CheckIploadQty(Optional CardNO As String, Optional Row As Long) As Boolean
CheckIploadQty = False
If CardNO = "" Then Exit Function
If CheckUplaodSerialInData(CardNO) = True Then
CheckIploadQty = True
Exit Function
Else
CheckIploadQty = False
End If
Dim i As Integer
With VSFlexGrid3
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("CardNO"))) = CardNO And i <> Row Then
CheckIploadQty = True
Exit Function
End If
Next i
End With
End Function
Sub FillGridTypeTrans()
    Dim i As Integer
    Dim k As Integer
    
    If Me.TxtModFlg.text <> "R" Then
        With Me.VSFlexGrid3
            k = .rows
            .rows = .rows + 1
            For i = k To .rows - 1
                .TextMatrix(i, .ColIndex("CardNO")) = CardNO.text
                .TextMatrix(i, .ColIndex("QtyDownload")) = val(QtyDownload.text)
                .TextMatrix(i, .ColIndex("CardNO2")) = CardNO2.text
                .TextMatrix(i, .ColIndex("QtyDischarge")) = val(QtyDischarge.text)
                .TextMatrix(i, .ColIndex("ItemID")) = val(DcboItems.BoundText)
                .TextMatrix(i, .ColIndex("UnitID")) = val(cmbUnitName.BoundText)
                .TextMatrix(i, .ColIndex("UnitName")) = (cmbUnitName.text)
                
                .TextMatrix(i, .ColIndex("ItemCode")) = TxtItemCode.text
                .TextMatrix(i, .ColIndex("ItemName")) = DcboItems.text
                .TextMatrix(i, .ColIndex("BillDate")) = BillDate.value
                .TextMatrix(i, .ColIndex("loadingInvoice")) = loadingInvoice.text
            Next i
        End With
        CardNO.text = ""
        QtyDownload.text = ""
        CardNO2.text = ""
        QtyDischarge.text = ""
        TxtItemCode.text = ""
        DcboItems.BoundText = ""
        loadingInvoice.text = ""
        CardNO.SetFocus
    End If
End Sub

Private Sub ISButton4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
ISButton4_Click
End If
End Sub

Private Sub ISButton5_Click()
If Me.TxtModFlg.text <> "R" Then
With Me.Fg_Journal
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
    
End If

End Sub

Private Sub LblLink_Click()

    Dim FirstPeriod As Date
    
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide2.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub RdTyped_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then
If RdTyped(0).value = True Then
TxtTotal.text = val(Me.TxtPrice.text) * val(TxtDistance.text)
End If
If RdTyped(1).value = True Then
TxtTotal.text = val(Me.TxtPrice.text)
End If
If RdTyped(2).value = True Then
TxtTotal.text = val(Me.TxtPrice.text) * val(Me.TxtTotalQty.text)
End If
End If
VSFlexGrid3.Visible = True
Dim i As Long
If RdTyped(3).value Or RdTyped(4).value Then
    For i = 1 To VSFlexGrid3.Cols - 1
        VSFlexGrid3.ColHidden(i) = True
    Next
    
End If

VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCityID")) = True
VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemID")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCityID")) = True
Select Case Index
Case 0
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Vehname")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromPrice")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToPrice")) = True
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("FromPrice")) = "„‰ "
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("ToPrice")) = "«·Ï"
         VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("QtyDownload")) = False
                VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("TotalValue")) = False
   
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Price")) = False
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("UnitName")) = False
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemName")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemCode")) = False
    
    If SystemOptions.UserInterface = ArabicInterface Then
           ' VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;ﬂÃ„"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;Kg"
            End If
         
    
Case 1
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Vehname")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromPrice")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToPrice")) = True
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("FromPrice")) = "„‰ "
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("ToPrice")) = "«·Ï"
    
    If SystemOptions.UserInterface = ArabicInterface Then
            'VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰|#4;· —|#3;Õ„Ê·… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;Kg |#2;RD |#3;Weight|#4;Litr|#5;Weight"
            End If
          
    
Case 2
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCity")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Vehname")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromPrice")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToPrice")) = True
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("FromPrice")) = "„‰ "
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("ToPrice")) = "«·Ï"
    If SystemOptions.UserInterface = ArabicInterface Then
           ' VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰|#4;· —|#5;Õ„Ê·… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;Kg |#2;RD |#3;Weight|#4;Litr|#5;Weight"
            End If
          
    
Case 3
VSFlexGrid3.Visible = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCity")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCity")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Vehname")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromPrice")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToPrice")) = False
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Price")) = False
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("QtyDownload")) = False
   VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("TotalValue")) = False
     
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("UnitName")) = False
     VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemName")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemCode")) = False
     
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("FromPrice")) = "„‰"
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("ToPrice")) = "«·Ï"
    
    If SystemOptions.UserInterface = ArabicInterface Then
            'VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#4;· —"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#4;Litr"
            End If
          
    

Case 4
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCity")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCity")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Vehname")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromPrice")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToPrice")) = False
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Price")) = False
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("FromPrice")) = "„‰ "
    VSFlexGrid3.TextMatrix(0, VSFlexGrid3.ColIndex("ToPrice")) = "«·Ï"
    If SystemOptions.UserInterface = ArabicInterface Then
           ' VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;Kg"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Typed")) = "#1;Kg"
            End If
Case 5
Case 6
End Select
VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("BillDate")) = ""
VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("FromCityID")) = True
VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ItemID")) = True
    VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("ToCityID")) = True
    VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("QtyDownload")) = ""
    VSFlexGrid3.ColComboList(VSFlexGrid3.ColIndex("Price")) = ""
    'VSFlexGrid3.ComboList = ""
    
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)

    If CBoBasedON.ListIndex = 3 Then
        If KeyCode = vbKeyF3 Then
            Order_no_search2.show
            Order_no_search2.RetrunType = 3
        End If
    Else
        If KeyCode = vbKeyF3 Then
            Order_no_search.show
            Order_no_search.RetrunType = 0
        End If
    End If
End Sub
Private Sub TxtBasedNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBasedNo_LostFocus
    End If
End Sub

Private Sub TxtBasedNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If val(DcbBasedOn.ListIndex) = 1 Then
        If KeyCode = vbKeyF3 Then
            Unload FrmInsurancesSearch
            FrmInsurancesSearch.BankInx = 702
            FrmInsurancesSearch.SendForm = 7
            FrmInsurancesSearch.show
        End If
    End If
    
    If val(DcbBasedOn.ListIndex) = 3 Then
        If KeyCode = vbKeyF3 Then
            Unload FrmInsurancesSearch
            FrmInsurancesSearch.BankInx = 703
            FrmInsurancesSearch.SendForm = 7
            FrmInsurancesSearch.show
        End If
    End If
End Sub
Private Sub TxtBasedNo_LostFocus()
    If Me.TxtModFlg.text <> "R" Then
        If val(DcbBasedOn.ListIndex) = 1 Then
            checkUsedLoadingOrder
            RetriveOrders val(TxtBasedNo.text)
            ChCarType_Click (0)
        End If
         If val(DcbBasedOn.ListIndex) = 3 Then
          '  checkUsedLoadingOrder
            RetriveOrders2 val(TxtBasedNo.text)
            ChCarType_Click (0)
        End If
    End If
End Sub

Private Sub txtCityFromCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If txtCityFromCode.text = "" Then
            Me.DcCityFromId.BoundText = ""
        Else
            Me.DcCityFromId.BoundText = GetGovernmentID(Trim$(Me.txtCityFromCode.text))
        End If
    End If
End Sub

Private Sub txtCityToCode_Change()
'Me.txtCityToCode.text = GetGovernmentCode(val(Me.DcCityToId.BoundText))
End Sub

Private Sub txtCityToCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        If txtCityToCode.text = "" Then
            Me.DcCityToId.BoundText = ""
        Else
            Me.DcCityToId.BoundText = GetGovernmentID(Trim$(Me.txtCityToCode.text))
        End If
    End If
End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub
Private Sub TxtDes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If
End Sub
Private Sub checkUsedLoadingOrder()
If val(TxtBasedNo.text) = 0 Then Exit Sub
    Dim Msg As String
    Dim sql As String
    Dim i As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    If Me.TxtModFlg.text = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Then
        sql = "select BasedNo,NoteID,NoteSerial1 From notes_all where notetype=370 and BasedNo = " & val(TxtBasedNo.text)
    ElseIf Me.TxtModFlg.text = "E" Then
        sql = "select BasedNo,NoteID,NoteSerial1 From notes_all where notetype=370 and NoteID != " & XPTxtID.text & " And BasedNo = " & val(TxtBasedNo.text)
    End If
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs2.RecordCount <> 0 Then
      rs2.MoveFirst
        If Me.TxtModFlg.text = "N" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ «” Œœ«„ «„—  Õ„Ì· ﬂ«‰ „” Œœ„ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â–« «·«„— „” Œœ„ ›Ì «·”Ã· —ﬁ„" & rs2("NoteSerial1").value
            Else
                Msg = "Can't use this order it's been used befor in recored with No. " & rs2("NoteSerial1").value
            End If
            MsgBox Msg
            TxtBasedNo.text = ""
            TxtBasedNo.SetFocus
            Exit Sub
        ElseIf Me.TxtModFlg.text = "E" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ «” Œœ«„ «„—  Õ„Ì· ﬂ«‰ „” Œœ„ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â–« «·«„— „” Œœ„ ›Ì «·”Ã· —ﬁ„" & rs2("NoteSerial1").value
            Else
                Msg = "Can't use this order it's been used befor in recored with No. " & rs2("NoteSerial1").value
            End If
            MsgBox Msg
            TxtBasedNo.text = ""
            TxtBasedNo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtDistance_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        TXTTravelPrice = Round(val(TxtDistance) * val(TxtKmPrice), 2)
        txtDriverValue = Round((val(TXTTravelPrice) * val(TxtDriverPercentage)) / 100, 2)

    End If

End Sub

Private Sub txtDriverPercentage_Change()
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        txtDriverValue = Round((val(TXTTravelPrice) * val(TxtDriverPercentage)) / 100, 2)
    End If
End Sub
Function CalValues()
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        TXTTravelPrice = Round(val(TxtDistance) * val(TxtKmPrice), 2)
        txtDriverValue = Round((TXTTravelPrice * TxtDriverPercentage) / 100, 2)
    End If
End Function
Private Sub txtDriverValue_Change()
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        addFixedExpenses
    End If
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If
End Sub

Private Sub TxtItemCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 100
        FrmItemSearch.show vbModal
    End If
End Sub
Function CheckAllocation() As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select * from TblTripTypesTransport  where allocations=1 and  NotesallID =" & val(XPTxtID.text) & "  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckAllocation = True
Else
CheckAllocation = False
End If
End Function
Private Sub txtKmPrice_Change()
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        TXTTravelPrice = Round(val(TxtDistance) * val(TxtKmPrice), 2)
        txtDriverValue = Round((val(TXTTravelPrice) * val(TxtDriverPercentage)) / 100, 2)
        addFixedExpenses
    End If
End Sub

Private Sub TxtModFlg_Change()

    On Error GoTo ErrTrap
ALLButton3.Enabled = False
    Select Case Me.TxtModFlg.text
        Case "R"
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " ”ÃÌ· »Ì«‰«  «·—Õ·«   "
            Else
                Me.Caption = "Expenses"
            End If
            
            If CheckOrderStuts() = 1 Or TxtBasedNo.text = "" Or val(DcbBasedOn.ListIndex) = 0 Then
            ALLButton3.Enabled = False
            Else
            ALLButton3.Enabled = True
            End If
            Me.VSFlexGrid1.Enabled = False
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
            'XPCboProfLevel.Locked = True
            'XPTxtProfMail.Locked = True
            'XPTxtPhone.Locked = True
            'XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            
            End If
        
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "  ”ÃÌ· »Ì«‰«  «·—Õ·«    (ÃœÌœ)"
            Else
                Me.Caption = "Expenses(New Record)"
            End If
        
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
            'Me.XPBtnMove(0).Enabled = False
            'Me.XPBtnMove(1).Enabled = False
            'Me.XPBtnMove(2).Enabled = False
            'Me.XPBtnMove(3).Enabled = False
        
            'XPTxtVal.locked = False
            'XPCboProfLevel.Locked = False
            'XPTxtProfMail.Locked = False
            'XPTxtPhone.Locked = False
            'XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " ”ÃÌ· »Ì«‰«  «·—Õ·«    (  ⁄œÌ· )"
            Else
                Me.Caption = "Expenses(Edit Current Record)"
            End If
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
        
            CmdRemove.Enabled = True
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
            'XPCboProfLevel.Locked = False
            'XPTxtProfMail.Locked = False
            'XPTxtPhone.Locked = False
            'XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNoteSerial1_Change()
TxtSerial1.text = txtNoteSerial1.text
End Sub

Private Sub txtPrice_Change()
If Me.TxtModFlg.text <> "R" Then
Dim mIndex As Long
If RdTyped(0).value Then
    mIndex = 0
ElseIf RdTyped(1).value Then
    mIndex = 1
ElseIf RdTyped(2).value Then
    mIndex = 2
ElseIf RdTyped(3).value Then
    mIndex = 3
Else
    mIndex = 0
End If
RdTyped_Click (mIndex)
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
        DBCboClientName2.BoundText = CUSTID
    End If
End Sub

Private Sub TxtSerial1_Change()
TxtSerial1.text = txtNoteSerial1.text
End Sub

Private Sub txtTravelPrice_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        txtDriverValue = Round((val(TXTTravelPrice) * val(TxtDriverPercentage)) / 100, 2)

    End If

End Sub

Private Sub VehicleType_Change()
VehicleType_Click (0)
End Sub

Private Sub VehicleType_Click(Area As Integer)
RetriveClinCounr
If val(Me.VehicleType.BoundText) <> 0 Then
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
 Dcombos.GetCars Me.DCCar, , val(Me.VehicleType.BoundText)
End If
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
    Dim project_id As Integer

    With VSFlexGrid1

        Select Case .ColKey(Col)
    
            Case "Value", "opr_fullcode"
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If

                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
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
                    'Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    'Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    'Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    'Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
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

                    'Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    'Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    'Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    'Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    'Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    'Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    'Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    'Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                End If
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
       
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
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
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

                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                'End If
           
                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
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
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ReLineGrid

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)
            Case "Value"
                .ComboList = ""
            Case "Account_Serial"
                .ComboList = ""
            Case "Des"
                .ComboList = ""
                'Cancel = True
        End Select
    End With
End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
    Sendkeys "{F4}"
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    Dim project_id As Integer
    Dim whrstring As String

    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "opr_fullcode"
                    
                project_id = get_project_id(dcproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = .BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
            
            Case "AccountName"
         
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                whrstring = getProjectAccountwhereString(project_id)
                
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
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    StrSQL = StrSQL + "and (" + whrstring + ")"
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                
                End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrItemID As String
    Dim StrUnitID As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sgl As String

    With VSFlexGrid2
        Select Case .ColKey(Col)
            Case "LineNo"
                '.TextMatrix(Row, .ColIndex("LineNo")) = setID_Line
            Case "KItem"
                StrItemID = .ComboData
                LngRow = .FindRow(StrItemID, .FixedRows, .ColIndex("LineNo"), False, True)
                .TextMatrix(Row, .ColIndex("KItemID")) = StrItemID
                StrItemID = .TextMatrix(Row, .ColIndex("KItemID"))

            Case "KUnit"
                StrUnitID = .ComboData
                LngRow = .FindRow(StrUnitID, .FixedRows, .ColIndex("KUnitID"), False, True)
                .TextMatrix(Row, .ColIndex("KUnitID")) = StrUnitID
                StrUnitID = .TextMatrix(Row, .ColIndex("KUnitID"))
        End Select

        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

    End With

    ReLineGrid

End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With VSFlexGrid2
        Select Case .ColKey(Col)

            Case "Count"
                .ComboList = ""
            Case "LineNo"
                Cancel = True
                
        End Select
    End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String
    
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "KItem"
                StrSQL = "select * from tblitems  where GroupID in ( "
                StrSQL = StrSQL & " SELECT     GroupID "
                StrSQL = StrSQL & " From dbo.Groups"
                StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"

                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
           
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "KUnit"
                StrSQL = "select * from TblUnites"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                         
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList1 = VSFlexGrid2.BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList1 = VSFlexGrid2.BuildComboList(rs, "UnitNamee", "UnitID")
                End If
           
                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If
                
                .ComboList = StrComboList1
        End Select

    End With
End Sub

Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrItemID As String
    Dim StrUnitID As String
    Dim IntRes As Integer
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
     Dim StrAccountCode As String
    Dim LngRow As Long
    Dim sgl As String
    With VSFlexGrid3
        Select Case .ColKey(Col)
            Case "ItemName"
                StrItemID = .ComboData
                LngRow = .FindRow(StrItemID, .FixedRows, .ColIndex("LineNo"), False, True)
                .TextMatrix(Row, .ColIndex("ItemID")) = StrItemID
                StrItemID = .TextMatrix(Row, .ColIndex("ItemID"))
                StrSQL = " SELECT     Fullcode"
                StrSQL = StrSQL & " From dbo.TblItems"
                StrSQL = StrSQL & "  Where (ItemID = " & val(.TextMatrix(Row, .ColIndex("ItemID"))) & ")"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                Else
                .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                End If
            Case "ItemCode"
                 StrSQL = " SELECT   * "
                StrSQL = StrSQL & " From dbo.TblItems"
                StrSQL = StrSQL & " WHERE     (Fullcode = N'" & .TextMatrix(Row, .ColIndex("ItemCode")) & "')"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                Else
                .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("ItemName")) = ""
                .TextMatrix(Row, .ColIndex("ItemID")) = 0
                End If
             Case "FromCity"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("VehicleType"), False, True)
             .TextMatrix(Row, .ColIndex("FromCityID")) = StrAccountCode
        Case "ToCity"
             StrAccountCode = .ComboData
             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("VehicleType"), False, True)
             .TextMatrix(Row, .ColIndex("ToCityID")) = StrAccountCode
           Case "CardNO"
           If CheckIploadQty(.TextMatrix(Row, .ColIndex("CardNO")), Row) = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ ﬂ—  «· Õ„Ì· „ÊÃÊœ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â·  —Ìœ «·„Ê«’·…"
             Else
                Msg = "The card number already exists" & CHR(13)
                Msg = Msg & "Confirm Continue"
             End If
             IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
             If IntRes = vbNo Then
                .TextMatrix(Row, .ColIndex("CardNO")) = ""
                Exit Sub
             End If
             End If
          Case "CardNO2"
            If CheckIDownLoadQty(.TextMatrix(Row, .ColIndex("CardNO2")), Row) = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ ﬂ—  «· ›—Ì€ „ÊÃÊœ „”»ﬁ« " & CHR(13)
                Msg = Msg & "Â·  —Ìœ «·„Ê«’·…"
             Else
                Msg = "The card number already exists" & CHR(13)
                Msg = Msg & "Confirm Continue"
             End If
            IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
            If IntRes = vbNo Then
              .TextMatrix(Row, .ColIndex("CardNO2")) = ""
            Exit Sub
            End If
            End If
            
        Case "Price", "QtyDownload"
            .TextMatrix(Row, .ColIndex("TotalValue")) = val(.TextMatrix(Row, .ColIndex("Price"))) * val(.TextMatrix(Row, .ColIndex("QtyDownload")))
        
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
            Case "CardNO"
                .ComboList = ""
            Case "BillDate"
                .ComboList = ""
            Case "QtyDownload"
                .ComboList = ""
             Case "CardNO2", "QtyDownload"
                .ComboList = ""
            Case "QtyDischarge"
                .ComboList = ""
             Case "ItemCode"
                .ComboList = ""
              Case "Price", "FromPrice", "ToPrice", "FromCity", "ToCity"
                .ComboList = ""
        End Select
    End With
End Sub

Private Sub VSFlexGrid3_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    With VSFlexGrid3
        Select Case .ColKey(Col)
            Case "ItemName"
                StrSQL = "select * from tblitems  where GroupID in ( "
                StrSQL = StrSQL & " SELECT     GroupID "
                StrSQL = StrSQL & " From dbo.Groups"
                StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid3.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = VSFlexGrid3.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
           
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
              Case "FromCity"
               StrSQL = " Select TblCountriesGovernments.GovernmentID ID,TblCountriesGovernments.GovernmentName Name,TblCountriesGovernments.GovernmentName NameE from TblCountriesGovernments "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "ID")
              Else
                    StrComboList = .BuildComboList(rs, "NameE", "ID")
             End If
             .ComboList = StrComboList

            Case "ToCity"
               StrSQL = " Select TblCountriesGovernments.GovernmentID ID,TblCountriesGovernments.GovernmentName Name,TblCountriesGovernments.GovernmentName NameE from TblCountriesGovernments "
               Set rs = New ADODB.Recordset
               rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "ID")
              Else
                    StrComboList = .BuildComboList(rs, "NameE", "ID")
             End If
             .ComboList = StrComboList
        Case "QtyDownload"
            .ComboList = ""
        Case "BillDate"
            .ComboList = ""
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
    Me.TxtModFlg.text = "R"
    TxtModFlg_Change
    Exit Sub
ErrTrap:
End Sub
Sub RetriveOrders(Optional ID As Double)
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim CarTypeID As Integer
    
    sql = "select * from TblOrderUpload where ID=" & ID & ""
    If Me.TxtModFlg.text = "N" Then
        sql = sql & " and IsTravel is null "
        sql = sql & " and TblOrderUpload.ID Not In  (Select IsNull(BasedNo,0) from notes_all where NoteType = 370)"
    ElseIf Me.TxtModFlg.text = "E" Then
        sql = sql & " and IsTravel is null or ID in (Select BasedNo from notes_all  where NoteID=" & val(XPTxtID.text) & " ) "
    End If

    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs2.RecordCount > 0 Then
        If Not IsNull(rs2("CarType").value) Then
            If val(rs2("CarType").value) = 1 Then
                ChCarType(1).value = True
            Else
                ChCarType(0).value = True
            End If
        Else
            ChCarType(0).value = True
        End If
        TxtRent.text = IIf(IsNull(rs2("TxtRent").value), 0, rs2("TxtRent").value)
        TxtPartPrice.text = IIf(IsNull(rs2("PartPrice").value), 0, rs2("PartPrice").value)
        TxtPrice1.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        TxtPrice.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        txtContainerNo = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        txtto.text = IIf(IsNull(rs2("Remarks").value), 0, rs2("Remarks").value)
        TxtTotal1.text = IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
        DBCboClientName2.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
        DcCityFromId.BoundText = IIf(IsNull(rs2("CityID").value), 0, rs2("CityID").value)
        DcCityToId.BoundText = IIf(IsNull(rs2("CityID2").value), 0, rs2("CityID2").value)
        DCCar.BoundText = IIf(IsNull(rs2("CarID").value), 0, rs2("CarID").value)
        DcbCar2.BoundText = IIf(IsNull(rs2("CarID2").value), 0, rs2("CarID2").value)
        DcbSupplem.BoundText = IIf(IsNull(rs2("SupplemID").value), 0, rs2("SupplemID").value)
        DcbSupplem2.BoundText = IIf(IsNull(rs2("SupplemID2").value), 0, rs2("SupplemID2").value)
        TxtLeaderName.text = IIf(IsNull(rs2("LeaderName").value), "", rs2("LeaderName").value)
        TxtIDNo.text = IIf(IsNull(rs2("IDNo").value), "", rs2("IDNo").value)
        TxtOrderNo.text = IIf(IsNull(rs2("OrderNo").value), "", rs2("OrderNo").value)
        txtContainerNo.text = IIf(IsNull(rs2("OrderNo").value), "", rs2("OrderNo").value)
        TxtTypGoods.text = IIf(IsNull(rs2("TypGoods").value), "", rs2("TypGoods").value)
        DBCboClientName.BoundText = IIf(IsNull(rs2("CustId1").value), "", rs2("CustId1").value)
        DCEmp.BoundText = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
        
        
        
        sql = " Select TravKItemDet1.* , " & ID & " as loadingInvoice, TravKItemDet1.Count as QtyDownload,tblItems.ItemName,tblItems.ItemNamee,TblUnites.UnitName,TblUnites.UnitNamee from TravKItemDet1 "
        sql = sql & "  left OUTER join tblItems On TblItems.ItemID = TravKItemDet1.ItemID"
        sql = sql & " left OUTER join TblUnites On TblUnites.UnitID= TravKItemDet1.UnitID where MasterID = " & val(rs2!ID & "")
        sql = sql & " and IsNull(TravKItemDet1.ItemID,0) <> 0"
        VSFlexGrid3.rows = 1
       loadgrid sql, VSFlexGrid3, True, False
       rs2.Close
       rs2.Open sql, Cn, adOpenStatic, adLockReadOnly
       If Not rs2.EOF Then
            DcboItems.BoundText = val(rs2!ItemID & "")
            cmbUnitName.BoundText = val(rs2!UnitID & "")
            QtyDownload = val(rs2!count & "")
       End If
       ReLineGrid
    Else
        DCEmp.BoundText = ""
        DBCboClientName.BoundText = ""
        TxtOrderNo.text = ""
        TxtLeaderName.text = ""
        TxtIDNo.text = ""
        DBCboClientName2.BoundText = 0
        DcbSupplem2.BoundText = 0
        DcbSupplem.BoundText = 0
        DcbCar2.BoundText = 0
        DCCar.BoundText = 0
        DcCityToId.BoundText = 0
        DcCityFromId.BoundText = 0
        TxtPartPrice.text = ""
        TxtPrice1.text = ""
        TxtTotal1.text = ""
        VSFlexGrid3.rows = 1
        VSFlexGrid3.rows = 2
    End If
End Sub
Public Sub RetriveOrders2(Optional ID As Double)
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim CarTypeID As Integer
    
    
sql = "   SELECT        TblClientTransContr.ID, TblClientTransContr.CompID, TblClientTransContr.LockedID, TblClientTransContr.UserID, TblClientTransContr.FromDate, TblClientTransContr.Todate,"
sql = sql & "                       TblClientTransContr.Remarks, TblClientTransContr.Typed,  TblVendorCars.BoardNo,  TblClientTransContrDet.ClintTransID,"
sql = sql & "                          TblClientTransContrDet.VehicleType , TblClientTransContrDet.Price , TblClientTransContrDet.Remarks AS Expr7,  TblClientTransContrDet.FromPrice,"
sql = sql & "                          TblClientTransContrDet.ToPrice, TblClientTransContrDet.FromCityID, TblClientTransContrDet.ToCityID, TblCountriesGovernments_1.GovernmentName as ToCityName, TblCountriesGovernments.GovernmentName AS FromCity,"
sql = sql & "                          TblCustemers.CusName , TblCustemers.CusID, TblCustemers.CusNamee, TblVendorCars.nBoardNo, TblVendorCars.ChasisNo, TblVendorCars.BrandID, TblVendorCars.ModelID,"
sql = sql & "                   TblItems.ItemName,TblItems.ItemNamee,TblItems.ItemCode,TblItems.ItemID,TblUnites.UnitID,TblUnites.UnitName,TblUnites.UnitNamee"
sql = sql & " FROM            TblCountriesGovernments AS TblCountriesGovernments_1 RIGHT OUTER JOIN"
sql = sql & "                          TblClientTransContrDet ON TblCountriesGovernments_1.GovernmentID = TblClientTransContrDet.ToCityID LEFT OUTER JOIN"
sql = sql & "                          TblCountriesGovernments ON TblClientTransContrDet.FromCityID = TblCountriesGovernments.GovernmentID RIGHT OUTER JOIN"
sql = sql & "                          TblClientTransContr LEFT OUTER JOIN"
sql = sql & "                          TblVendorCars ON TblClientTransContr.VehicleType = TblVendorCars.ID LEFT OUTER JOIN"
sql = sql & "                          TblCustemers ON TblClientTransContr.CompID = TblCustemers.CusID ON TblClientTransContrDet.ClintTransID = TblClientTransContr.ID"
sql = sql & "                  LEFT OUTER JOIN TblItems  On TblItems.ItemID =TblClientTransContrDet.ItemID"
sql = sql & "                  LEFT OUTER JOIN TblUnites  On TblUnites.UnitID =TblClientTransContrDet.UnitID"
sql = sql & " Where (1 = 1) and TblClientTransContr.ID=" & ID


    
    'sql = "select * from TblOrderUpload where "
    If Me.TxtModFlg.text = "N" Then
      '  sql = sql & " and IsTravel is null "
    ElseIf Me.TxtModFlg.text = "E" Then
      '  sql = sql & " and IsTravel is null or ID in (Select BasedNo from notes_all  where NoteID=" & val(XPTxtID.text) & " ) "
    End If
sql = sql & " ORDER BY TblClientTransContr.ID"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs2.RecordCount > 0 Then
        If Not IsNull(rs2("Typed").value) Then
            If val(rs2("Typed").value) = 1 Then
                RdTyped(1).value = True
            ElseIf val(rs2("Typed").value) = 2 Then
                RdTyped(2).value = True
            ElseIf val(rs2("Typed").value) = 3 Then
                RdTyped(3).value = True
            ElseIf val(rs2("Typed").value) = 4 Then
                RdTyped(4).value = True
            Else
                RdTyped(0).value = True
            End If
        Else
            RdTyped(0).value = True
        End If
        RdTyped_Click val(rs2("Typed").value)
        'TxtRent.text = IIf(IsNull(rs2("TxtRent").value), 0, rs2("TxtRent").value)
        'TxtPartPrice.text = IIf(IsNull(rs2("PartPrice").value), 0, rs2("PartPrice").value)
        TxtPrice1.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        TxtPrice.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
        txtto.text = IIf(IsNull(rs2("Remarks").value), 0, rs2("Remarks").value)
        TxtTotal1.text = TxtPrice ' IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
        DBCboClientName2.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
        DcCityFromId.BoundText = IIf(IsNull(rs2("FromCityID").value), 0, rs2("FromCityID").value)
        
        DcCityFromId.BoundText = IIf(IsNull(rs2("FromCityID").value), 0, rs2("FromCityID").value)
        DcboItems.BoundText = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
        cmbUnitName.BoundText = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
        TXTTravelPrice.text = TxtPrice.text
        DcCityToId.BoundText = IIf(IsNull(rs2("ToCityID").value), 0, rs2("ToCityID").value)
        'DCCar.BoundText = IIf(IsNull(rs2("CarID").value), 0, rs2("CarID").value)
        'DcbCar2.BoundText = IIf(IsNull(rs2("CarID2").value), 0, rs2("CarID2").value)
        'DcbSupplem.BoundText = IIf(IsNull(rs2("SupplemID").value), 0, rs2("SupplemID").value)
        'DcbSupplem2.BoundText = IIf(IsNull(rs2("SupplemID2").value), 0, rs2("SupplemID2").value)
        'TxtLeaderName.text = IIf(IsNull(rs2("LeaderName").value), "", rs2("LeaderName").value)
        TxtIDNo.text = IIf(IsNull(rs2("ID").value), "", rs2("ID").value)
        'TxtOrderNo.text = IIf(IsNull(rs2("OrderNo").value), "", rs2("OrderNo").value)
        'TxtTypGoods.text = IIf(IsNull(rs2("TypGoods").value), "", rs2("TypGoods").value)
        DBCboClientName.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
        'DCEmp.BoundText = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
    Else
        DCEmp.BoundText = ""
        DBCboClientName.BoundText = ""
        TxtOrderNo.text = ""
        TxtLeaderName.text = ""
        TxtIDNo.text = ""
        DBCboClientName2.BoundText = 0
        DcbSupplem2.BoundText = 0
        DcbSupplem.BoundText = 0
        DcbCar2.BoundText = 0
        DCCar.BoundText = 0
        DcCityToId.BoundText = 0
        DcCityFromId.BoundText = 0
        TxtPartPrice.text = ""
        TxtPrice1.text = ""
        TxtTotal1.text = ""
    End If
    
    loadgrid sql, VSFlexGrid3, True
    'VSFlexGrid3.IsSubtotal(VSFlexGrid3.rows - 1) = True
    TxtPrice = ""
    Dim i As Long
    For i = 1 To VSFlexGrid3.rows - 1
        TxtPrice = val(TxtPrice) + val(VSFlexGrid3.TextMatrix(i, VSFlexGrid3.ColIndex("Price")))
    Next
    If Not rs2.EOF Then
        TxtPrice = rs2!Price & ""
    End If
    XPTxtVal = TxtPrice
    XPTxtValView = TxtPrice
    'TxtPrice = VSFlexGrid3.Aggregate(flexSTSum, VSFlexGrid3.FixedRows, VSFlexGrid3.ColIndex("Price"), VSFlexGrid3.rows - 1, VSFlexGrid3.ColIndex("Price"))
    TxtQtyDownload = 1
    TxtQtyDischarge = 1
    If RdTyped(3).value = True Then
    
        RdTyped_Click 3
    ElseIf RdTyped(4).value = True Then
        RdTyped_Click 4
    End If
End Sub
Public Function FindRecbyNoteserial1(ByVal NoteSerial1 As String)
    On Error GoTo ErrTrap
     
    rs.Find "NoteSerial1='" & NoteSerial1 & "'", , adSearchForward, 1
    
    Retrive
    Exit Function
ErrTrap:

  End Function

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

    If 1 = 2 Then
        Exit Sub
    Else
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
    
    If Not IsNull(rs("CarType").value) Then
        If val(rs("CarType").value) = 1 Then
            ChCarType(1).value = True
        Else
            ChCarType(0).value = True
        End If
    Else
        ChCarType(0).value = True
    End If

    Me.txtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.TxtTotalQty.text = IIf(IsNull(rs("TotalQty").value), 0, rs("TotalQty").value)
    Me.TxtTotal.text = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
    Me.TxtPrice.text = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
    Me.VehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), "", rs("VehicleType").value)
    If Not IsNull(rs("Typed").value) Then
        If (rs("Typed").value) = 2 Then
            RdTyped(2).value = True
        ElseIf (rs("Typed").value) = 3 Then
            RdTyped(3).value = True
            RdTyped_Click 3
        ElseIf (rs("Typed").value) = 4 Then
            RdTyped(4).value = True
            RdTyped_Click 4
        ElseIf (rs("Typed").value) = 1 Then
            RdTyped(1).value = True
        Else
            RdTyped(0).value = True
        End If
    Else
        RdTyped(0).value = True
    End If
    TxtManualNo.text = IIf(IsNull(rs("ManualNO").value), "", rs("ManualNO").value)
    Me.DcbAccount.BoundText = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
        txtContainerNo = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
    DcboItems.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
    cmbUnitName.BoundText = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
    
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.TxtOrderNo.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TXT_A_NoteID.text = IIf(IsNull(rs("A_NoteID").value), "", (rs("A_NoteID").value))
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)
    ''/////////////
    Me.DcbBasedOn.ListIndex = IIf(IsNull(rs("BasedOnID2").value), 0, rs("BasedOnID2").value)
    TxtBasedNo.text = IIf(IsNull(rs("BasedNo").value), "", rs("BasedNo").value)
    TxtPartPrice.text = IIf(IsNull(rs("PartPrice").value), "", rs("PartPrice").value)
    TxtPrice1.text = IIf(IsNull(rs("Price1").value), "", rs("Price1").value)
    TxtTotal1.text = IIf(IsNull(rs("Total1").value), "", rs("Total1").value)
    TxtQtyDownload.text = IIf(IsNull(rs("QtyDownload").value), "", rs("QtyDownload").value)
    TxtQtyDischarge.text = IIf(IsNull(rs("QtyDischarge").value), "", rs("QtyDischarge").value)
    DcbTypeTransport.BoundText = IIf(IsNull(rs("TypeTransportID").value), "", rs("TypeTransportID").value)
    Me.DBCboClientName2.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
    TxtIDNo.text = IIf(IsNull(rs("IDNo").value), "", rs("IDNo").value)
    TxtLeaderName.text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
    'TxtNationality.Text = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
    Me.DcbCar2.BoundText = IIf(IsNull(rs("CarID2").value), "", rs("CarID2").value)
    Me.DcbSupplem.BoundText = IIf(IsNull(rs("SupplemID").value), "", rs("SupplemID").value)
    Me.DcbSupplem2.BoundText = IIf(IsNull(rs("SupplemID2").value), "", rs("SupplemID2").value)
    TxtTypGoods.text = IIf(IsNull(rs("TypGoods").value), "", rs("TypGoods").value)
    TxtOrderNo.text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
    Me.DcbHarbor.BoundText = IIf(IsNull(rs("HarborID").value), "", rs("HarborID").value)
    Me.DcbShip.BoundText = IIf(IsNull(rs("ShipID").value), "", rs("ShipID").value)
    If Not IsNull(rs("TripType").value) Then
        If (rs("TripType").value) = 1 Then
            ChTripType(1).value = True
        Else
            ChTripType(0).value = True
        End If
    Else
        ChTripType(0).value = True
    End If
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    If IsNull(rs("Destribute").value) Then
        chkDestribute.value = vbUnchecked
    ElseIf (rs("Destribute").value) = False Then
        chkDestribute.value = vbUnchecked
    Else
        chkDestribute.value = vbChecked
    End If
    'Me.DBCboClientName.BoundText = rs("CusID").value
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 1 Then '«Ã·
        'Me.CboPaymentType.ListIndex = 1
        'Me.DcboBox.BoundText = ""
        'Me.DcboBankName.BoundText = rs("BankID").value
        'Me.TxtChequeNumber.text = rs("ChqueNum").value
        'Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.CboPaymentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPaymentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPaymentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    End If
    CboPayMentType_Change
    If Not IsNull(rs("BasedONID").value) Then
        Me.CBoBasedON.ListIndex = rs("BasedONID").value
    Else
        Me.CBoBasedON.ListIndex = 0
    End If
    'Me.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Or rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    End If
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.oldTxtSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    Me.DcCityFromId.BoundText = IIf(IsNull(rs("CityFromId").value), "", rs("CityFromId").value)
    Me.DcCityToId.BoundText = IIf(IsNull(rs("CityToId").value), "", rs("CityToId").value)
    Me.TxtLocation.text = IIf(IsNull(rs("Location").value), "", rs("Location").value)
    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    Me.DCEmp.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
    Me.TxtDistance.text = IIf(IsNull(rs("Distance").value), 0, rs("Distance").value)
    Me.TxtKmPrice.text = IIf(IsNull(rs("KmPrice").value), 0, rs("KmPrice").value)
    Me.TXTTravelPrice.text = IIf(IsNull(rs("TravelPrice").value), 0, rs("TravelPrice").value)
    Me.TxtDriverPercentage.text = IIf(IsNull(rs("DriverPercentage").value), 0, rs("DriverPercentage").value)
    Me.txtDriverValue.text = IIf(IsNull(rs("DriverValue").value), 0, rs("DriverValue").value)
    Me.txtDriverEra.text = IIf(IsNull(rs("DriverEra").value), 0, rs("DriverEra").value)
    Me.txtNoR.text = IIf(IsNull(rs("NoR").value), 0, rs("NoR").value)
    
    Me.txtTotalExpenses.text = IIf(IsNull(rs("TotalExpenses").value), 0, rs("TotalExpenses").value)
    Me.TxtComm.text = IIf(IsNull(rs("comm").value), 0, rs("comm").value)
    Me.DtpStartDate.value = rs("StartDate").value
    Me.DtpEndDate.value = rs("EndDate").value
    loadingInvoice.text = IIf(IsNull(rs("loadingInvoice").value), "", rs("loadingInvoice").value)
    txtRecNo = IIf(IsNull(rs("RecNo").value), "", rs("RecNo").value)    '  (rs!RecNo & "")
    txtWeight = IIf(IsNull(rs("Weight").value), 0, rs("Weight").value)   'val(rs!Weight & "")
    If Not IsNull(rs("StartTime").value) Then
        StartTime.value = FormatDateTime(rs("StartTime").value, vbShortTime)
    End If
    If Not IsNull(rs("EndTime").value) Then
        EndTime.value = FormatDateTime(rs("EndTime").value, vbShortTime)
    End If
    Me.txtKMCounterBeforeStart.text = IIf(IsNull(rs("KMCounterBeforeStart").value), 0, rs("KMCounterBeforeStart").value)
    Me.TxtKMCounterAtEnd.text = IIf(IsNull(rs("KMCounterAtEnd").value), 0, rs("KMCounterAtEnd").value)
     Me.TxtRent.text = IIf(IsNull(rs("TxtRent").value), 0, rs("TxtRent").value)
     
    
    
    lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)
    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then 'Õ”«Ì« 
        Me.VSFlexGrid1.Visible = True
        Me.Fg_Journal.Visible = False

        StrSQL = "SELECT TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value], dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode"
        StrSQL = StrSQL + " FROM dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
        StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
        StrSQL = StrSQL + " Where (not (Trip=1 )) and (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
            .rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .rows
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value) / IIf(val(txtNoR) <> 0, val(txtNoR), 1)
                
                .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                RsDev.MoveNext
            Next i
    
        End With
        Exit Sub
    End If

    Me.VSFlexGrid1.Visible = False
    Me.Fg_Journal.Visible = True

    '«·„’—Ê›« 
    '-----------------------------------------------------------------------------
    If chkDestribute.value = vbUnchecked Then
        'StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        'StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        
        'StrSQL = "SELECT dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"
        
        'StrSQL = "SELECT dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + "Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
        StrSQL = "SELECT dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
        StrSQL = StrSQL + " FROM dbo.ACCOUNTS INNER JOIN"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        StrSQL = StrSQL + " Where ( trip is null ) and   (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value))
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    Else
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                    End If
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value) / IIf(val(txtNoR) <> 0, val(txtNoR), 1)
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
            End With
        End If
    End If

    '-----------------------------------------------------------------------------«·„’—Ê›«  «·„Ê“⁄Â
    If chkDestribute.value = vbChecked Then
    
        'StrSQL = "SELECT dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
        'StrSQL = StrSQL + " FROM dbo.ACCOUNTS INNER JOIN"
        'StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        'StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & Val(Me.XPTxtID.text) & ")"
        'StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        StrSQL = "Select * from ExpensesDetails where noteid=" & val(XPTxtID.text)
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            'Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            'Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst
            'For i = 1 To RsDev.RecordCount
                'If RsDev("Credit_Or_Debit").value = 0 Then
                    'Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                'ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    'Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                'End If
                'RsDev.MoveNext
            'Next i
    
            'RsDev.MoveFirst
    
            With Me.Fg_Journal
                If Me.dcproject.BoundText = "" Then
                    .rows = .FixedRows + RsDev.RecordCount
                Else
                    .rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("Destribute")) = IIf(IsNull(RsDev("Destribute").value), 0, RsDev("Destribute").value)
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value))
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("ExpensesName").value), "", RsDev("ExpensesName").value)
                    Else
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("ExpensesName").value), "", RsDev("ExpensesName").value)
                    End If
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Des").value), "", RsDev("Des").value)
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value) / IIf(val(txtNoR) <> 0, val(txtNoR), 1)
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
            End With
        End If
    End If
    TxtSerial1.text = txtNoteSerial1.text
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    FillDestributionsToAll
    fillItemsGrid
    RetriveGridOrders
    RetriveTypeTransport
    ChCarType_Click (0)
    
    
         If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
    fillapprovData
    TxtModFlg_Change
    Exit Sub
ErrTrap:
End Sub
Sub UnPayedFlag(Optional Row As Integer = 0)
Dim i As Integer
With FGOrders
If Row = 0 Then
Row = .rows - 1
End If
Cn.Execute "Update TblOrderUpload set  IsTravel = null where Id=" & val(TxtBasedNo.text) & " "
For i = 1 To Row
Cn.Execute "Update TblOrderUpload set  IsTravel = null where Id=" & val(.TextMatrix(i, .ColIndex("OrderNo"))) & " "
Next i
End With
End Sub

Private Sub SaveData()

    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim bankDes As String
    Dim AcountSup_Exp As String
    'On Error GoTo ErrTrap

RentValue = 0
    If Me.TxtModFlg.text <> "R" Then
        Account_Code_dynamic = get_account_code_branch(2, my_branch)
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If
            Exit Sub
        ElseIf Account_Code_dynamic = "NO account" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
                MsgBox "Sales Account Not Defined in this Branch", vbCritical
            End If
            Exit Sub
        End If
                
        Account_Code_dynamic3 = get_account_code_branch(71, my_branch)
        
        If Account_Code_dynamic3 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If
            Exit Sub
        ElseIf Account_Code_dynamic3 = "NO account" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ«  «·⁄„Ê·«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
                MsgBox "Sales Account Not Defined in this Branch", vbCritical
            End If
            Exit Sub
        End If
                
        If Me.CboPaymentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPaymentType.SetFocus
            Exit Sub
        End If
    
        If Trim(Me.DBCboClientName.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì·..!!"
            Else
                Msg = "Select Customer..!!"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
    
        If Me.CboPaymentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            If Trim(Me.DBCboClientName.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì·..!!"
                Else
                    Msg = "Select Customer..!!"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
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
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
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
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
        End If
        'If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            'Else
                'Msg = "Cheque Due Date Not Valid...!!"
            'End If
            'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            'DtpChequeDueDate.SetFocus
            'SendKeys "{F4}"
            'Exit Sub
        'End If
    End If
    
    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“—⁄Â «Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
        Else
            MsgBox "This Voucher Have Distributed and not Distributed Expenses", vbCritical
        End If
        Exit Sub
    End If
                
'    If Trim(Me.DcCityFromId.BoundText) = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«— «·—Õ·… „‰ ..!!"
'        Else
'            Msg = "Select Trip From..!!"
'        End If
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        DcCityFromId.SetFocus
'        Sendkeys "{F4}"
'        Exit Sub
'    End If
'
'    If Trim(Me.DcCityToId.BoundText) = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«— «·—Õ·… «·Ï ..!!"
'        Else
'            Msg = "Select Trip To..!!"
'        End If
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        DcCityToId.SetFocus
'        Sendkeys "{F4}"
'        Exit Sub
'    End If
'
'    If Trim(Me.DCCar.BoundText) = "" And DcbCar2.text = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«—   «·„⁄œÂ/«·”Ì«—… ..!!"
'        Else
'            Msg = "Select Car  ..!!"
'        End If
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        'DCCar.SetFocus
'        Sendkeys "{F4}"
'        Exit Sub
'    End If
'
    'If Trim(Me.DCEmP.BoundText) = "" And TxtLeaderName.Text = "" Then
        'If SystemOptions.UserInterface = ArabicInterface Then
            'Msg = "ÌÃ» ≈Œ Ì«—   «·”«∆ﬁ ..!!"
        'Else
            'Msg = "Select Driver  ..!!"
        'End If
        'MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        'DCEmP.SetFocus
        'SendKeys "{F4}"
        'Exit Sub
    'End If
    
    DcboCreditSide2.BoundText = DcbAccount.BoundText
    
    'If DcboCreditSide2.BoundText = "" Then
        'If SystemOptions.UserInterface = ArabicInterface Then
            'MsgBox "·«ÌÊÃœ  ⁄ÂœÂ ··”«∆ﬁ"
        'Else
            'MsgBox "Pleae enter hand data of driver"
        'End If
        'Exit Sub
    'End If
    
    If Me.TxtModFlg.text = "N" Then
        If Me.CboPaymentType.ListIndex = 0 Then
            If val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
                    Exit Sub
                End If
            End If
        End If
    ElseIf Me.TxtModFlg.text = "E" Then
        If Me.CboPaymentType.ListIndex = 0 Then
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

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then
        GoTo xx
    End If

    Dim i As Integer
    If SystemOptions.AllowSaveTripWithoutExpen = True Then
    Else
        With Fg_Journal
            For i = .FixedRows To .rows - 1
                If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                    '////////////////////////////////////////notes
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·«  ÌÊÃœ „’—Ê› ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                    Else
                        MsgBox "Select Expenses in line no" & i, vbCritical
                    End If
                    Exit Sub
                End If
            Next i
        End With
    
        With Fg_Journal
            For i = .FixedRows To .rows - 1
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
xx:
    calcnets     '-------------------------------------------------------------------------------------------
 
    '-------------------------------------------------------------------------------------------
    
    Dim Vchr_result As String
    Dim notes_result As String

    'If TxtSerial1.Text = "" Then
        'Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 37, 370)
        'If Vchr_result = "error" Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ —Õ·«   ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            'Else
                'MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
            'End If
        'Else
            'If Vchr_result = "" Then
                'If SystemOptions.UserInterface = ArabicInterface Then
                    'MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                'Else
                    'MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                'End If
            'Else
            'End If
        'End If
    'End If
    Dim mm As Long
    Dim mValue As Double
    For mm = 1 To Fg_Journal.rows - 1
        
        mValue = mValue + val(Fg_Journal.TextMatrix(mm, Fg_Journal.ColIndex("value")))
    Next
    
    If TxtSerial.text = "" Or TxtSerial.text = "0" Then
        notes_result = Notes_coding(val(Dcbranch.BoundText), XPDtbTrans.value)
        If notes_result = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
            End If
        Else
            If notes_result = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    MsgBox "You must Define JE Coding ": Exit Sub
                End If
            Else
            End If
        End If
    End If
    TxtSerial = notes_result
    If RdTyped(3).value Or RdTyped(0).value Then mValue = val(TxtPrice)
   ' If mValue = 0 Then TxtSerial.text = ""
    
    'TxtSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value) 'kk
    
    Cn.BeginTrans
    BeginTrans = True
    Dim A_NoteID As Long

    '///////////////NOTESALL
    If TxtModFlg.text = "N" Then
        XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
        Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=370"))
        
        rs.AddNew
        rs("NoteID").value = val(XPTxtID.text)
        TxtSerial1.text = txtNoteSerial1.text
        Me.oldTxtSerial1.text = Trim$(Me.TxtSerial1.text)
    ElseIf Me.TxtModFlg.text = "E" Then
        UnPayedFlag
        Cn.Execute "Delete from TblTripTypesTransport where NotesallID=" & val(XPTxtID.text) & " "
        Cn.Execute "Delete from TblTravelTransDet where NotesallID=" & val(XPTxtID.text) & " "
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        'dcEmp_Change
        StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        
        StrSQL = "Delete From ExpensesDetails Where Noteid =" & val(XPTxtID.text) & "  or NoteSerial1='" & Me.TxtSerial1.text & "'"
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
        If txtNoteSerial1.text = "" Then
              txtNoteSerial1.text = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 74, 74)
        End If
        rs("NoteSerial1").value = IIf(Me.txtNoteSerial1 <> "", val(txtNoteSerial1.text), Null)
        TxtSerial1.text = txtNoteSerial1.text
        'Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        ''/////////////
        rs("ManualNo").value = TxtManualNo.text
        rs("IDNo").value = TxtIDNo.text
        rs("LeaderName").value = TxtLeaderName.text
        'rs("Nationality").value = TxtNationality.Text
        rs("SupplemID").value = val(Me.DcbSupplem.BoundText)
        rs("SupplemID2").value = val(Me.DcbSupplem2.BoundText)
        rs("VendorID").value = val(DBCboClientName2.BoundText)
        rs("CarID2").value = val(Me.DcbCar2.BoundText)
        rs("OrderNo").value = TxtOrderNo.text
        rs("TypGoods").value = TxtTypGoods.text
        rs("BasedNo").value = val(TxtBasedNo.text)
        rs("BasedOnID2").value = val(Me.DcbBasedOn.ListIndex)
        rs("HarborID").value = val(Me.DcbHarbor.BoundText)
        rs("ShipID").value = val(Me.DcbShip.BoundText)
        rs("PartPrice").value = val(TxtPartPrice.text)
        rs("Price1").value = val(TxtPrice1.text)
        rs("ContainerNo").value = IIf(txtContainerNo.text = "", Null, Trim(txtContainerNo.text))
        rs("Total1").value = val(TxtTotal1.text)
        If ChTripType(1).value = True Then
            rs("TripType").value = 1
        Else
            rs("TripType").value = 0
        End If
        rs("ShipID").value = val(Me.DcbShip.BoundText)
        '//////////
        rs("foxy_no").value = val(Text1.text)
        rs("order_no").value = txt_ORDER_NO.text
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, XPTxtVal.text)
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
        rs("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text)) & bankDes
        rs("branch_no").value = val(Me.Dcbranch.BoundText)
        rs("QtyDownload").value = val(Me.TxtQtyDownload.text)
        rs("QtyDischarge").value = val(Me.TxtQtyDischarge.text)
        rs("TypeTransportID").value = val(Me.DcbTypeTransport.BoundText)
        rs("AccountCode").value = Me.DcbAccount.BoundText
        If ChCarType(1).value = True Then
            rs("CarType").value = 1
        Else
            rs("CarType").value = 0
        End If
        rs("CusID").value = Null
        rs("NoteType").value = 370
        rs("NoteDate").value = XPDtbTrans.value
        rs("UserID").value = user_id
        rs("TotalQty").value = val(TxtTotalQty.text)
        rs("Total").value = val(TxtTotal.text)
        rs("Price").value = val(TxtPrice.text)
        rs("VehicleType").value = val(VehicleType.BoundText)
        If RdTyped(2).value = True Then
            rs("Typed").value = 2
        ElseIf RdTyped(1).value = True Then
            rs("Typed").value = 1
        ElseIf RdTyped(3).value = True Then
            rs("Typed").value = 3
        ElseIf RdTyped(4).value = True Then
            rs("Typed").value = 4
        Else
            rs("Typed").value = 0
        End If
        rs("UserID").value = user_id
        If chkDestribute.value = vbChecked Then
            Destribute = True
        Else
            Destribute = False
        End If
        rs("Destribute").value = Destribute
        rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
        If CBoBasedON.ListIndex > -1 Then
            rs("BasedONID").value = CBoBasedON.ListIndex
        Else
            rs("BasedONID").value = 0
        End If
  
    If Me.CboPaymentType.ListIndex = 0 Then
        rs("BoxID").value = val(DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 0
        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "   ’—› „‰  " & DcboBox.text
        Else
            bankDes = "   Payed From  " & DcboBox.text
        End If
    ElseIf Me.CboPaymentType.ListIndex = 1 Then '«Ã·
        'rs("BoxID").value = Null
        'rs("BankID").value = Val(Me.DcboBankName.BoundText)
        'rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        'rs("DueDate").value = Me.DtpChequeDueDate.value
        'rs("NoteCashingType").value = 1
        'If SystemOptions.UserInterface = ArabicInterface Then
            'bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        'Else
            'bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        'End If
    
        rs("BoxID").value = Null
        rs("BankID").value = Null
        rs("CusID").value = val(Me.DBCboClientName.BoundText)
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 1
        'If SystemOptions.UserInterface = ArabicInterface Then
            'bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        'Else
            'bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        
        'End If
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 3
        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »‘Ìﬂ „”œœ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        Else
            bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        End If
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 2
        If SystemOptions.UserInterface = ArabicInterface Then
            bankDes = "  ’—› »ÕÊ«·…  —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        Else
            bankDes = "  Bank Transfere No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        End If
    End If
    rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
    rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    rs("Buy").value = "0"
    rs("Remark").value = XPMTxtRemarks.text
    rs!RecNo = (txtRecNo)
    rs!Weight = val(txtWeight)
    'If TxtSerial1.Text = "" Then
        'TxtSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 37, 370)
    'End If
    'rs("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
    
    If TxtSerial.text = "" And mValue <> 0 Then
        TxtSerial.text = Notes_coding(val(Dcbranch.BoundText), XPDtbTrans.value)
    End If
    rs("NoteSerial").value = val(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
    rs("OldNoteSerial1").value = Trim$(Me.oldTxtSerial1.text) '
    rs("CusID").value = val(Me.DBCboClientName.BoundText)
    rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    rs("numbering_type1").value = sand_numbering_type(37) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
    rs("sanad_year").value = year(XPDtbTrans.value)
    rs("sanad_month").value = Month(XPDtbTrans.value)
    rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    If Me.TxtModFlg.text = "N" Then
        A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
        TXT_A_NoteID.text = A_NoteID
    Else
        A_NoteID = val(TXT_A_NoteID.text)
    End If
    

    rs("ItemID").value = val(DcboItems.BoundText)
    rs("UnitID").value = val(cmbUnitName.BoundText)
    
    
    rs("A_NoteID").value = val(A_NoteID)
    rs("CityFromId").value = val(Me.DcCityFromId.BoundText)
    rs("CityToId").value = val(Me.DcCityToId.BoundText)
    rs("Location").value = Trim$(Me.TxtLocation.text)
    rs("CarId").value = val(Me.DCCar.BoundText)
    rs("DriverId").value = val(Me.DCEmp.BoundText)
    rs("Distance").value = val(Me.TxtDistance.text)
    rs("KmPrice").value = val(Me.TxtKmPrice.text)
    rs("TravelPrice").value = val(Me.TXTTravelPrice.text)
    rs("DriverPercentage").value = val(Me.TxtDriverPercentage.text)
    rs("DriverValue").value = val(Me.txtDriverValue.text)
    rs("DriverEra").value = val(Me.txtDriverEra.text)
    rs("NoR").value = val(Me.txtNoR.text)

    rs("comm").value = val(Me.TxtComm.text)
    rs("TotalExpenses").value = val(Me.txtTotalExpenses.text)
    rs("StartDate").value = Me.DtpStartDate.value
    rs("EndDate").value = Me.DtpEndDate.value
    rs("StartTime").value = FormatDateTime(Me.StartTime.value, vbShortTime)
    rs("EndTime").value = FormatDateTime(Me.EndTime.value, vbShortTime)
    rs("KMCounterBeforeStart").value = val(Me.txtKMCounterBeforeStart.text)
    rs("KMCounterAtEnd").value = val(Me.TxtKMCounterAtEnd.text)
    rs("TxtRent").value = val(Me.TxtRent.text)
    
    rs.update
    
    Dim project_id As Integer
    project_id = get_project_id(dcproject.BoundText, "expanses_account")
    
    '/////////////////////Accounts Õ”«Ì« 
    Dim line_no  As Integer

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then
        Set RsNotes = New ADODB.Recordset
        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
        If TxtModFlg.text = "N" Then
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        
        'Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        'rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        'rs("foxy_no").value = Val(Text1.text)
        
        'œ«∆‰ Õ”«»«  «·„‘—Ê€
        RsNotes.AddNew
        RsNotes("NoteID").value = A_NoteID
        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
        RsNotes("order_no").value = txt_ORDER_NO.text
        RsNotes("notes_all").value = Me.XPTxtID.text
        RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.text), 0, val(XPTxtVal.text))
        RsNotes("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        RsNotes("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
        'RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
        If Me.CboPaymentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Then '«Ã·
            'RsNotes("BoxID").value = Null
            'RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
            'RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            'RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            'RsNotes("NoteCashingType").value = 1
            RsNotes("BoxID").value = Null
            RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 2
        'ElseIf Me.CboPaymentType.ListIndex = 2 Then
            'RsNotes("CusID").value = DCVendor.BoundText
        End If
     
        'RsNotes("BasedONID").value = Me.CBoBasedON.ListIndex
        RsNotes("NoteType").value = 370
        RsNotes("NoteDate").value = XPDtbTrans.value
        RsNotes("UserID").value = user_id
        'rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
        'rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
        RsNotes("Buy").value = "0"
        RsNotes("Remark").value = txt_general_des.text & bankDes
        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„   ”‰œ ’—›
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
        rs("loadingInvoice").value = loadingInvoice.text
        RsNotes.update
    
        Dim IntDEV_Type As Integer
        Dim SngDEV_Value As Single
        line_no = 1
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            
        If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, IIf(Not IsNumeric(XPTxtVal.text), 0, val(XPTxtVal.text)), 1, txt_general_des.text & bankDes, A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.text), , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
            
        '„œÌ‰ Õ”«»«  «·„‘—Ê€
        With VSFlexGrid1
            line_no = 2
            For i = .FixedRows To .rows - 1
                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.text), project_id, .TextMatrix(i, .ColIndex("opr_fullcode")), , , , , , val(Me.Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    line_no = line_no + 1
                End If
            Next i
        End With
        'TxtModFlg.text = "R"
        GoTo ll
    End If

    '„’—Ê›« 
    '//////////////////////////////////////Notes////////////////////////////////////
    If Destribute = True Then
        If createDest = True Then
            GoTo ll
        Else
            Exit Sub
        End If
    End If
    Dim s As String
    Set RsNotes = New ADODB.Recordset
    
    GoTo 1010
    s = "Select * from Notes WHERE NoteID = -1"
    RsNotes.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        Set RsDev = New ADODB.Recordset
        
           
           s = "Select * from DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID = 0"
            RsDev.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
        ' RsDev.Open "Select * from DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Dim NoteID As String
        ' ﬁÌœ «·«Ì—«œ« 
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '   «·⁄„Ì· «Ê  «·ÿ—› «·„œÌ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        RsNotes.AddNew
        NoteID = CStr(new_id("Notes", "NoteID", "", True))
        RsNotes("NoteID").value = CStr(NoteID)
        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
        
        If val(XPTxtVal) = 0 Then
        
            RsNotes("Note_Value").value = IIf(IsNumeric(TxtPrice.text), TxtPrice.text, 0) '* IIf(val(txtNoR) <> 0, val(txtNoR), 1)
        Else
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '* IIf(val(txtNoR) <> 0, val(txtNoR), 1)
        End If
        RsNotes("Remark").value = txt_general_des.text & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text
        RsNotes("foxy_no").value = val(Text1.text)
        If Me.CboPaymentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            'RsNotes("BoxID").value = Null
            'RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
            'RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            'RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            'RsNotes("NoteCashingType").value = 1
            RsNotes("BoxID").value = Null
            RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 2
        End If
        RsNotes("CusID").value = Null
        RsNotes("NoteType").value = 370
        RsNotes("NoteDate").value = XPDtbTrans.value
        RsNotes("UserID").value = user_id
        RsNotes("notes_all").value = Me.XPTxtID.text
        RsNotes("NoteSerial").value = val(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = val(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(37) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
     '   RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
        RsNotes("Remark").value = txt_general_des.text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsNotes.update
    
        '«·ÿ—› «·„œÌ‰  «·Õ“Ì‰… «Ê  «·⁄„Ì· «Ê «·»‰ﬂ
        
        Dim TotalValue As Double
   '     Totalvalue = val(txtTravelPrice.Text) + val(TxtComm.Text)
        
   '    If SystemOptions.CarsRevenuePerOwner = True Then
   '             Totalvalue = val(TxtComm.Text)
   '     End If
        
        
        If val(TXTTravelPrice.text) + val(TxtComm.text) > 0 And SystemOptions.CarsRevenuePerOwner = False Then
            'IIf(IsNumeric(Me.TXTTravelPrice.Text), val(TXTTravelPrice.Text), 0) + IIf(IsNumeric(Me.TxtComm.Text), val(TxtComm.Text), 0)
            line_no = 1
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            If Trim(DcboCreditSide.BoundText) = "" Then
                DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            End If
            RsDev("Account_Code").value = DcboCreditSide.BoundText
            RsDev("Value").value = (IIf(IsNumeric(Me.TXTTravelPrice.text), val(TXTTravelPrice.text), 0) + IIf(IsNumeric(Me.TxtComm.text), val(TxtComm.text), 0)) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
            RsDev("Credit_Or_Debit").value = 0
            'rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("Double_Entry_Vouchers_Description").value = "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text    '  txt_general_des.text & "  «Ì—«œ«  «·—Õ·… —ﬁ„  " & Me.TxtSerial1.text     'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
            RsDev("trip").value = 1
            RsDev("CarId").value = val(DCCar.BoundText)
            RsDev.update
        End If
        line_no = line_no + 1

        ' «·ÿ—› «·œ«∆‰ «Ì—«œ«  «·„»Ì⁄« 
        If val(TXTTravelPrice.text) > 0 And SystemOptions.CarsRevenuePerOwner = False Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = Account_Code_dynamic
            RsDev("Value").value = (IIf(IsNumeric(Me.TXTTravelPrice.text), val(TXTTravelPrice.text), 0)) * IIf(val(txtNoR) <> 0, val(txtNoR), 1) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            'rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
            RsDev("trip").value = 1
            RsDev("CarId").value = val(DCCar.BoundText)
            RsDev.update
        End If
               
        ' «·ÿ—› «·œ«∆‰ ⁄„Ê·«   «·„»Ì⁄« 
        If val(TxtComm.text) > 0 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = Account_Code_dynamic3
            RsDev("Value").value = IIf(IsNumeric(Me.TxtComm.text), val(TxtComm.text), 0) * IIf(val(txtNoR) <> 0, val(txtNoR), 1) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            'rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
            RsDev("trip").value = 1
            RsDev("CarId").value = val(DCCar.BoundText)
            RsDev.update
        End If
                
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
  
        line_no = line_no + 1
   
        Dim ExpensesID As Double
        With Fg_Journal
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
                    RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
                    RsNotes("Destribute").value = IIf(.TextMatrix(i, .ColIndex("Destribute")) = "", 0, Destribute)
                    RsNotes("Remark").value = txt_general_des.text & bankDes & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text
                    RsNotes("ExpensesRemark").value = .TextMatrix(i, .ColIndex("des"))
                    RsNotes("foxy_no").value = val(Text1.text)
                    RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
                    If Me.CboPaymentType.ListIndex = 0 Then
                        RsNotes("BoxID").value = val(DcboBox.BoundText)
                        RsNotes("BankID").value = Null
                        RsNotes("ChqueNum").value = Null
                        RsNotes("DueDate").value = Null
                        RsNotes("NoteCashingType").value = 0
                    ElseIf Me.CboPaymentType.ListIndex = 1 Then
                        'RsNotes("BoxID").value = Null
                        'RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                        'RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                        'RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        'RsNotes("NoteCashingType").value = 1
                        RsNotes("BoxID").value = Null
                        RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
                        RsNotes("BankID").value = Null
                        RsNotes("ChqueNum").value = Null
                        RsNotes("DueDate").value = Null
                        RsNotes("NoteCashingType").value = 1
                    ElseIf Me.CboPaymentType.ListIndex = 3 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 3
                    ElseIf Me.CboPaymentType.ListIndex = 2 Then
                        RsNotes("BoxID").value = Null
                        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                        RsNotes("NoteCashingType").value = 2
                    End If
                    
                    If txt_ORDER_NO.text <> "" Then
                        RsNotes("order_no").value = txt_ORDER_NO.text
                    Else
                        RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                    End If

                    RsNotes("CusID").value = Null
                    RsNotes("NoteType").value = 370
                    RsNotes("NoteDate").value = XPDtbTrans.value
                    RsNotes("UserID").value = user_id
                    RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                    RsNotes("notes_all").value = Me.XPTxtID.text
                    RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                    RsNotes("NoteSerial1").value = val(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                    RsNotes("numbering_type1").value = sand_numbering_type(37) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
                    RsNotes("sanad_year").value = year(XPDtbTrans.value)
                    RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                    RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                    RsNotes("remark").value = txt_general_des.text & bankDes & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text
                    RsNotes.update
              
                    '////////////////////////////////////////notes
   
                    project_id = get_project_id(dcproject.BoundText, "expanses_account")
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    
                    If ChCarType(1).value = True Then
                'salimhere
                
                        'AcountSup_Exp = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName2.BoundText))
                    AcountSup_Exp = .TextMatrix(i, .ColIndex("AccountCode"))
                    If checkRentAccount(.TextMatrix(i, .ColIndex("AccountCode"))) = True Then
                           RentValue = RentValue + .TextMatrix(i, .ColIndex("value"))
                    End If
                    
                    Else
                        AcountSup_Exp = .TextMatrix(i, .ColIndex("AccountCode"))
                    End If

                    If Destribute = False Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, AcountSup_Exp, .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), 0, .TextMatrix(i, .ColIndex("des")) & txt_general_des.text & bankDes & "—Õ·… —ﬁ„:" & Me.txtNoteSerial1.text & " ( " & DcCityFromId.text & "- " & DcCityToId.text & ")" & " ‘«Õ‰… —ﬁ„" & DCCar.text, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.Dcbranch.BoundText), val(Me.DCCar.BoundText)) = False Then
                            GoTo ErrTrap
                            
                        End If
            
                        line_no = line_no + 1
                        
                        
                        
                        
                    End If
                End If
            Next i
        End With
    TotalValue = (val(TXTTravelPrice.text) + val(TxtComm.text) + val(XPTxtVal)) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
        '«·ÿ—› «·œ«∆‰      ⁄Âœ… «·”«∆ﬁ
        RsNotes.AddNew
        NoteID = CStr(new_id("Notes", "NoteID", "", True))
        RsNotes("NoteID").value = CStr(NoteID)
        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
        RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
        RsNotes("Remark").value = txt_general_des.text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsNotes("foxy_no").value = val(Text1.text)

        If Me.CboPaymentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            'RsNotes("BoxID").value = Null
            'RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
            'RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            'RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            'RsNotes("NoteCashingType").value = 1
            RsNotes("BoxID").value = Null
            RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 2
        End If
        RsNotes("CusID").value = Null
        RsNotes("NoteType").value = 370
        RsNotes("NoteDate").value = XPDtbTrans.value
        RsNotes("UserID").value = user_id
        RsNotes("notes_all").value = Me.XPTxtID.text
        RsNotes("NoteSerial").value = val(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = val(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(37) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
        RsNotes("note_value_by_characters").value = WriteNo(Format(IIf(IsNumeric(TotalValue), TotalValue, 0), "0.00"), 0, True, ".")
        RsNotes("Remark").value = txt_general_des.text & bankDes 'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsNotes.update
    
        '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
 If val(XPTxtVal) > 0 Then
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = DcboCreditSide2.BoundText
        RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) * IIf(val(txtNoR) <> 0, val(txtNoR), 1) - RentValue '.TextMatrix(I, .ColIndex("VALUE"))-rentvalue
        RsDev("Credit_Or_Debit").value = 1
        'rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text & "  „’—Ê›«   «·—Õ·… —ﬁ„  " & Me.TxtSerial1.text     'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("notes_all").value = Me.XPTxtID.text
        RsDev("carid").value = val(Me.DCCar.BoundText)
                        
        RsDev.update
End If
        
        
      'ÿ«·«ÌÃ«— ··„Ê—œ
 If val(RentValue) > 0 Then
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName2.BoundText))
        RsDev("Value").value = RentValue
        RsDev("Credit_Or_Debit").value = 1
        'rsdev("Double_Entry_Vouchers_Description").value = txtto ' .TextMatrix(I, .ColIndex("des"))
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text & "  «ÌÃ«—  «·—Õ·… —ﬁ„  " & Me.TxtSerial1.text     'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("notes_all").value = Me.XPTxtID.text
        RsDev("carid").value = val(Me.DCCar.BoundText)
                        
        RsDev.update
End If
        'GoTo ll
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        If Me.dcproject.BoundText <> "" Then
            '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
            RsNotes.AddNew
            NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = CStr(NoteID)
            RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
          
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) ' * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
            RsNotes("Remark").value = txt_general_des.text & bankDes

            If Me.CboPaymentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPaymentType.ListIndex = 1 Then
                'RsNotes("BoxID").value = Null
                'RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                'RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                'RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                'RsNotes("NoteCashingType").value = 1
                RsNotes("BoxID").value = Null
                RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPaymentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
            End If
               
            'If TXT_order_no.text <> "" Then
                'RsNotes("order_no").value = TXT_order_no.text
            'Else
                'RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
            'End If
            
            RsNotes("CusID").value = Null
            RsNotes("NoteType").value = 370
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
            'rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
            RsNotes("notes_all").value = Me.XPTxtID.text
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(37) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
            RsNotes.update
            
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = dcproject.BoundText
            RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) * IIf(val(txtNoR) <> 0, val(txtNoR), 1) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text & bankDes  ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
            'RsDev("project_id").value = project_id
                        
            RsDev.update
                    
            line_no = line_no + 1

            With Fg_Journal
                For i = .FixedRows To .rows - 1
                    'line_no = 2
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                        '////////////////////////////////////////notes
                        If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                            Else
                                MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                            End If
                        End If
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), 1, txt_general_des.text & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), , , , , setfoxy_Line, val(Me.XPTxtID.text), , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
                        line_no = line_no + 1
                    End If
                Next i
            End With

            Dim sql As String
            'sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(Val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & Val(TxtSerial.text)
            'Cn.Execute sql
            sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text) + val(Me.TXTTravelPrice.text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text)
            Cn.Execute sql
        End If

1010:
        saveItemsGrid
        SaveOrders
        SaveTypeTransport
        Cn.Execute " Update TblOrderUpload set  IsTravel = null where Id=" & val(TxtBasedNo.text) & " "
        
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        LblDevID.Caption = LngDevID
        lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If

ll:
    Cn.CommitTrans
    BeginTrans = False
   ' sql = "Update   notes set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) + val(Me.txtTravelPrice.Text) + val(Me.TxtComm.Text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text)
   ' Cn.Execute sql
    
   updateNotesValueAndNobytext val(XPTxtID.text) ',  Format(XPTxtVal.Text, "###.00")
   
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
            Retrive
Me.TxtModFlg.text = "R"
        Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
        
            lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
        
            Fg_Journal.Enabled = False
            Retrive
            Me.TxtModFlg.text = "R"
    End Select
    TxtModFlg_Change
    '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
    'If Me.DcCostCenter.BoundText <> "" Then
    save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "”‰œ ’—›", Me.XPDtbTrans.value
    'End If
    save_cost_center
        
    'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
    ' If saveExpensesDetails(0, TxtSerial.text, TxtSerial1.text, TXT_order_no.text, XPDtbTrans.value, Val(XPTxtID.text)) = True Then
    ' End If
    
    'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
    saveChequeBoxContents1 (val(Me.XPTxtID.text))
    
    TxtModFlg.text = "R"
Retrive
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

Function saveChequeBoxContents1(NoteID As Double)
    Exit Function

    If SystemOptions.banks_Accounts3 = False Then Exit Function
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    If CboPaymentType.ListIndex = 1 Then
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

Function createDest() As Boolean

    '„’—Ê›« 
    If CheckAllExpensesDistributed = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·”‰œ ÌÕ ÊÏ ⁄·Ï „’«—Ì› „Ê“—⁄Â «Œ—Ï €Ì— „Ê“⁄Â Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
        Else
            MsgBox "This Voucher Have Distributed and not Distributed Expenses", vbCritical
        End If

        createDest = False
        Exit Function
    End If

    '//////////////////////////////////////Notes////////////////////////////////////
    Dim RsNotes As ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
    Dim ExpensesID As Double
    Dim NoteID As String
 
    RsNotes.AddNew
    NoteID = CStr(new_id("Notes", "NoteID", "", True))
    RsNotes("NoteID").value = CStr(NoteID)
    RsNotes("Note_Value").value = val(XPTxtVal.text)
    RsNotes("Remark").value = txt_general_des.text
    RsNotes("foxy_no").value = val(Text1.text)
    RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)

    If Me.CboPaymentType.ListIndex = 0 Then
        RsNotes("BoxID").value = val(DcboBox.BoundText)
        RsNotes("BankID").value = Null
        RsNotes("ChqueNum").value = Null
        RsNotes("DueDate").value = Null
        RsNotes("NoteCashingType").value = 0
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        ' RsNotes("BoxID").value = Null
        ' RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
        ' RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        ' RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        ' RsNotes("NoteCashingType").value = 1
        RsNotes("BoxID").value = Null
        RsNotes("CusID").value = val(Me.DBCboClientName.BoundText)
        RsNotes("BankID").value = Null
        RsNotes("ChqueNum").value = Null
        RsNotes("DueDate").value = Null
        RsNotes("NoteCashingType").value = 1
                           
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 3
                            
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 2
                        
    End If

    If txt_ORDER_NO.text <> "" Then
        RsNotes("order_no").value = txt_ORDER_NO.text
    Else
              
    End If

    RsNotes("CusID").value = Null
    RsNotes("NoteType").value = 3
    RsNotes("NoteDate").value = XPDtbTrans.value
    RsNotes("UserID").value = user_id
    'RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
    RsNotes("notes_all").value = Me.XPTxtID.text
    RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(1) '‰Ê⁄  —ﬁÌ„ ”‰œ «·’—›
    RsNotes("sanad_year").value = year(XPDtbTrans.value)
    RsNotes("sanad_month").value = Month(XPDtbTrans.value)
    RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
    RsNotes("remark").value = txt_general_des.text
    RsNotes.update
              
    Dim line_no As Integer
    Dim i As Integer
    Dim project_id As Integer
    Dim LngDevID As Long

    With GridEstimatedCost
 
        line_no = 1

        For i = .FixedRows To .rows - 1
   
            If .TextMatrix(i, .ColIndex("AcountCode")) <> "" Then
                '////////////////////////////////////////notes
   
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                If Destribute = True Then
                    If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AcountCode")), .TextMatrix(i, .ColIndex("Netvalue")), 0, .TextMatrix(i, .ColIndex("Remarks")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                              
                    End If
                     
                    line_no = line_no + 1

                    If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, .TextMatrix(i, .ColIndex("Netvalue")), 1, .TextMatrix(i, .ColIndex("Remarks")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1), , , , , .TextMatrix(i, Fg_Journal.ColIndex("LineNo1")), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                              
                    End If
     
                    line_no = line_no + 1
                End If
        
            End If

        Next i

    End With

    createDest = True
    '
ErrTrap:
End Function

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
        rs("Remark").value = "”‰œ ’—› —ﬁ„ " & TxtSerial1 & "    " & Me.txt_general_des
 
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
 
    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 and  kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
        
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg_Journal
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("general_des").value = 1
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value")) * IIf(val(txtNoR) <> 0, val(txtNoR), 1)
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.text
                rs("kedno").value = Me.Text1.text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Function calcnets()

    If GridEstimatedCost.rows > 1 Then
        chkDestribute.value = vbChecked
    Else
        chkDestribute.value = vbUnchecked
    End If

    With Fg_Journal
      '  Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        
        
        If .rows > 1 Then
      Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
Else
 Me.XPTxtVal.text = 0
End If

    End With

    If SystemOptions.gldetails_or_gl_general = 0 And Me.dcproject.BoundText <> "" Then

        With Me.VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Undo()
    On Error GoTo ErrTrap
    Dim sql As String
    Dim sgl As String

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
        UnPayedFlag
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute "Delete from TblTripTypesTransport where NotesallID=" & val(XPTxtID.text) & " "
            StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute "Delete from TblTravelTransDet where NotesallID=" & val(XPTxtID.text) & " "
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            '        StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & Val(Me.TXT_A_NoteID)
            '   Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TravKItemDet Where NotesID = " & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            Cn.Execute "Update TblOrderUpload set OrderStuts =null where ID=" & val(TxtBasedNo.text) & " "
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.rows = 3
                    Fg_Journal.Enabled = False
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
    Dim SumRent As Double
    Dim SumPrice As Double
    Dim SumPrice2 As Double
    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

If .rows > 1 Then
      Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
Else
 Me.XPTxtVal.text = 0
End If
    End With
    IntCounter = 0
SumRent = 0
SumPrice = 0
    With Me.VSFlexGrid3

        For i = .FixedRows To .rows - 1
            If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
               SumRent = SumRent + val(.TextMatrix(i, .ColIndex("QtyDownload")))
               SumPrice = SumPrice + val(.TextMatrix(i, .ColIndex("QtyDischarge")))
               SumPrice2 = SumPrice2 + val(.TextMatrix(i, .ColIndex("TotalValue")))
            End If

        Next i

    End With
    TxtQtyDownload.text = SumRent
    TxtQtyDischarge.text = SumPrice
       IntCounter = 0
    With Me.VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With
  '''///////////
      IntCounter = 0
      SumRent = 0
      SumPrice = 0
    With Me.FGOrders

        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("OrderNo")) <> "" And .cell(flexcpChecked, i, .ColIndex("Selct")) = flexChecked Then
                SumRent = SumRent + val(.TextMatrix(i, .ColIndex("RentVAlue")))
               SumPrice = SumPrice + val(.TextMatrix(i, .ColIndex("Price")))
            End If

        Next i

    End With
    lbl(68).Caption = SumRent
    If SumPrice = 0 Then SumPrice = SumPrice2
    lbl(71).Caption = SumPrice
    
       IntCounter = 0
Dim SumQty  As Double
SumQty = 0
    With Me.VSFlexGrid2

        For i = .FixedRows To .rows - 1

            If val(.TextMatrix(i, .ColIndex("KItemID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
               SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
                
            End If

        Next i

    End With
TxtTotalQty.text = SumQty
If Me.TxtModFlg.text <> "R" And Trim(Me.TxtModFlg.text) <> "" Then


Dim mIndex As Long
If RdTyped(0).value Then
    mIndex = 0
ElseIf RdTyped(1).value Then
    mIndex = 1
ElseIf RdTyped(2).value Then
    mIndex = 2
ElseIf RdTyped(3).value Then
    mIndex = 3
Else
    mIndex = 0
End If
RdTyped_Click (mIndex)


End If
If SumPrice2 <> 0 Then TxtPrice = SumPrice2: TXTTravelPrice = SumPrice2: txtTotalExpenses = SumPrice2
If RdTyped(3) Then
    XPTxtValView = TxtPrice
    XPTxtValView = TxtPrice
End If
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
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
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
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=3 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
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
Function setID_Line() As Double
    
    Dim X As Double
    X = CStr(new_id("TravKDet", "id", "", True))
    setID_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "TravKDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    rs.AddNew
    
    rs("id").value = X
 
    rs.update
    
End Function
Sub RetriveMultyOrders()

    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Set rs2 = New ADODB.Recordset

    FGOrders.Clear flexClearScrollable, flexClearEverything
    FGOrders.rows = 1

    sql = " SELECT dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate, dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.EmpID,"
    sql = sql & " dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO,"
    sql = sql & " dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.TblOrderUpload.CarID2, dbo.TblVendorCars.BoardNo AS BoardNo2,"
    sql = sql & " dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.Total, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.CustId1,"
    sql = sql & " dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
    sql = sql & " FROM  dbo.TblCustemers RIGHT OUTER JOIN"
    sql = sql & " dbo.TblOrderUpload ON dbo.TblCustemers.CusID = dbo.TblOrderUpload.CustId1 LEFT OUTER JOIN"
    sql = sql & " dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
    sql = sql & " dbo.TblVendorCars ON dbo.TblOrderUpload.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
    sql = sql & " dbo.FixedAssets RIGHT OUTER JOIN"
    sql = sql & " dbo.TblCarsDataDet ON dbo.FixedAssets.id = dbo.TblCarsDataDet.PartID ON dbo.TblOrderUpload.SupplemID = dbo.TblCarsDataDet.PartID LEFT OUTER JOIN"
    sql = sql & " dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
    sql = sql & "  dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID"

    If Me.TxtModFlg.text = "N" Then
        sql = sql & "  where dbo.TblOrderUpload.IsTravel is null"
    ElseIf Me.TxtModFlg.text = "E" Then
        sql = sql & "  where dbo.TblOrderUpload.IsTravel is null or dbo.TblOrderUpload.ID in(select OrderNo from TblTravelTransDet where NotesallID=" & val(XPTxtID.text) & " )"
    End If
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs2.RecordCount > 0 Then
        rs2.MoveFirst
        With FGOrders
            .rows = .rows + rs2.RecordCount
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs2("ID").value), "", rs2("ID").value)
                .TextMatrix(i, .ColIndex("CustID")) = IIf(IsNull(rs2("CustId1").value), "", rs2("CustId1").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
                End If
                .TextMatrix(i, .ColIndex("TypedID")) = IIf(IsNull(rs2("CarType").value), 0, rs2("CarType").value) + 1
                .TextMatrix(i, .ColIndex("TypeDriver")) = IIf(IsNull(rs2("DrievType").value), 0, rs2("DrievType").value) + 1
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs2("Total").value), "", rs2("Total").value)
                If val(.TextMatrix(i, .ColIndex("TypedID"))) = 1 Then
                    .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(rs2("CarID").value), "", rs2("CarID").value)
                    .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs2("BoardNO").value), "", rs2("BoardNO").value)
                    .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs2("SupplemID").value), "", rs2("SupplemID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("Name").value), "", rs2("Name").value)
                    Else
                        .TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
                    End If
                Else
                    .TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("accessory").value), "", rs2("accessory").value)
                    .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs2("BoardNo2").value), "", rs2("BoardNo2").value)
                    .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(rs2("CarID2").value), "", rs2("CarID2").value)
                    .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs2("SupplemID2").value), "", rs2("SupplemID2").value)
                End If
                If val(.TextMatrix(i, .ColIndex("TypeDriver"))) = 1 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
                    Else
                        .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
                    End If
                Else
                    .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs2("LeaderName").value), "", rs2("LeaderName").value)
                End If
                rs2.MoveNext
            Next i
        End With
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Exit Sub
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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

    If Trim(TxtSerial1.text) <> "" Then
        oldTxtSerial1.text = TxtSerial1.text
        
    End If

    TxtSerial.text = ""
    TxtSerial1.text = ""
    txtNoteSerial1.text = ""
    
    RetriveClinCounr
    updaterowdate
End Sub
Function updaterowdate()
On Error Resume Next
Dim i As Integer
If VSFlexGrid3.rows = 0 Then Exit Function
With VSFlexGrid3

            For i = 1 To 1
                                  If .TextMatrix(i, .ColIndex("loadingInvoice")) = "" Then
                                 .TextMatrix(i, .ColIndex("BillDate")) = XPDtbTrans.value
                                 
                                 
                                End If
            Next i
End With

End Function
Private Sub XPTxtVal_Change()
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
    Dim fg As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set fg = FrmView.vsfGroup1.VSFlexGrid

    With fg
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
 
    Me.C1Tab1.TabCaption(0) = "Expenses"

    With Me.CBoBasedON
        .Clear
        .AddItem "Without"
        .AddItem "Purchase Invoices"
        .AddItem "Performa Invoices"
        .AddItem "Production Order"
    
    End With
lbl(79).Caption = "Account"
    ISButton1.Caption = "Attachments"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(29).Caption = "Driver Cost"
    lbl(4).Caption = "Trip #"
    lbl(1).Caption = "Date"

    Label3.Caption = "Branch"
  
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(24).Caption = "Customer"
    lbl(19).Caption = "Due Date"

    lbl(5).Caption = "To"
    lbl(0).Caption = "Based On"
    Frame2.Caption = "Trip Data"

    lbl(25).Caption = "From"
    lbl(42).Caption = "To"
    lbl(41).Caption = "Location"

    lbl(26).Caption = "Car"
    lbl(28).Caption = "Driver"
    Frame4.Caption = "Fin. Informations"

    lbl(31).Caption = "Distance KM"
    lbl(30).Caption = "Trip Cost"
    lbl(43).Caption = "Driver %"
    lbl(33).Caption = "Driver Era"
    lbl(44).Caption = "Comm."
    lbl(33).Caption = "Driver Era"
    lbl(32).Caption = "Desil"
    lbl(34).Caption = "Expenses"
    lbl(39).Caption = "Driver Cost"
    lbl(33).Caption = "Driver Era"

    Frame5.Caption = "Trips Info"
    lbl(35).Caption = "Start Date"
    lbl(37).Caption = "Start Time"

    lbl(36).Caption = "End Date"
    lbl(38).Caption = "End Time"
    lbl(39).Caption = "Km Before"
    lbl(40).Caption = "Km After"
    lbl(20).Caption = "Remarks"

    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Credit"
        .AddItem "Bank Transfer"
        .AddItem "Collected Cheque"
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Trips Data"
    Me.Ele(0).Caption = Me.Caption
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(1).Caption = "Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Total"
    Me.lbl(0).Caption = "Based On"
    Me.lbl(22).Caption = "Based On"
    Label3.Caption = "Branch"

    Me.lbl(5).Caption = "TO"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."
    Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"

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
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("des")) = "description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("order_no")) = "order no"

    End With

    With Me.GridEstimatedCost
        .TextMatrix(0, .ColIndex("Ser")) = "Index"
        .TextMatrix(0, .ColIndex("AcountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("BranchName")) = " Branch Name "

        .TextMatrix(0, .ColIndex("value")) = "Total Value"
        .TextMatrix(0, .ColIndex("Percentage")) = "Percentage"
        .TextMatrix(0, .ColIndex("Netvalue")) = "Distr Value"
        .TextMatrix(0, .ColIndex("REMARKS")) = "REMARKS "

    End With

End Sub

Private Sub XPTxtValView_Change()
    txtTotalExpenses.text = val(XPTxtValView.text)
End Sub
Sub SaveOrders()
    Dim i As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql_str As String
    Dim StrSQL As String
    If Me.TxtModFlg.text <> "R" Then
          sql_str = "select * from TblTravelTransDet "
            rs2.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    With Me.FGOrders
        For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Selct")) = flexChecked Then
            rs2.AddNew
            rs2("NotesallID").value = val(XPTxtID.text)
            rs2("RentVAlue").value = IIf((.TextMatrix(i, .ColIndex("RentVAlue"))) = "", Null, val(.TextMatrix(i, .ColIndex("RentVAlue"))))
            rs2("Selct").value = 1
            rs2("OrderNo").value = IIf((.TextMatrix(i, .ColIndex("OrderNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("OrderNo"))))
            rs2("CustID").value = IIf((.TextMatrix(i, .ColIndex("CustID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CustID"))))
            rs2("TypedID").value = IIf((.TextMatrix(i, .ColIndex("TypedID"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypedID"))))
            rs2("CarID").value = IIf((.TextMatrix(i, .ColIndex("CarID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CarID"))))
            rs2("PartID").value = IIf((.TextMatrix(i, .ColIndex("PartID"))) = "", Null, val(.TextMatrix(i, .ColIndex("PartID"))))
            rs2("TypeDriver").value = IIf((.TextMatrix(i, .ColIndex("TypeDriver"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDriver"))))
            rs2("DriverName").value = IIf((.TextMatrix(i, .ColIndex("DriverName"))) = "", Null, (.TextMatrix(i, .ColIndex("DriverName"))))
            rs2("Price").value = IIf((.TextMatrix(i, .ColIndex("Price"))) = "", Null, val(.TextMatrix(i, .ColIndex("Price"))))
            rs2.update
         End If
        Next i
    End With
  End If
  With FGOrders
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Selct")) = flexChecked Then
Cn.Execute "Update TblOrderUpload set  IsTravel =1 where Id=" & val(.TextMatrix(i, .ColIndex("OrderNo"))) & " "
End If
Next i
End With
End Sub
Function CheckUplaodSerialInData(Optional CardNO1 As String, Optional Typ As Integer = 0) As Boolean
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " select * from TblTripTypesTransport where NotesallID<>" & val(XPTxtID.text) & ""
If Typ = 1 Then
sql = sql & " and      (CardNO2 = N'" & CardNO1 & "') "
Else
sql = sql & " and      (CardNO = N'" & CardNO1 & "') "
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckUplaodSerialInData = True
Else
CheckUplaodSerialInData = False
End If
End Function

Sub SaveTypeTransport()
    Dim i As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql_str As String
    Dim StrSQL As String
    
    If Me.TxtModFlg.text <> "R" Then
        sql_str = "select * from TblTripTypesTransport "
        rs2.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Me.VSFlexGrid3
            For i = .FixedRows To .rows - 1
                rs2.AddNew
                If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 And SystemOptions.TripDateInsertDefulat = False Then
                    
                    rs2("NotesallID").value = val(XPTxtID.text)
                    rs2("CardNO").value = IIf((.TextMatrix(i, .ColIndex("CardNO"))) = "", Null, (.TextMatrix(i, .ColIndex("CardNO"))))
                    rs2("allocations").value = IIf((.TextMatrix(i, .ColIndex("allocations"))) = "", 0, (.TextMatrix(i, .ColIndex("allocations"))))
                    
                    
                    rs2("QtyDownload").value = IIf((.TextMatrix(i, .ColIndex("QtyDownload"))) = "", Null, val(.TextMatrix(i, .ColIndex("QtyDownload"))))
                    rs2("CardNO2").value = IIf((.TextMatrix(i, .ColIndex("CardNO2"))) = "", Null, (.TextMatrix(i, .ColIndex("CardNO2"))))
                    rs2("QtyDischarge").value = IIf((.TextMatrix(i, .ColIndex("QtyDischarge"))) = "", Null, val(.TextMatrix(i, .ColIndex("QtyDischarge"))))
                    rs2("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                    rs2("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                    rs2("BillDate").value = IIf((.TextMatrix(i, .ColIndex("BillDate"))) = "", Null, (.TextMatrix(i, .ColIndex("BillDate"))))
                    rs2("loadingInvoice").value = IIf((.TextMatrix(i, .ColIndex("loadingInvoice"))) = "", Null, (.TextMatrix(i, .ColIndex("loadingInvoice"))))
                    rs2.update
                ElseIf (.TextMatrix(i, .ColIndex("CardNO"))) <> "" And SystemOptions.TripDateInsertDefulat = True Then
                    'rs2.AddNew
                    rs2("NotesallID").value = val(XPTxtID.text)
                    rs2("CardNO").value = IIf((.TextMatrix(i, .ColIndex("CardNO"))) = "", Null, (.TextMatrix(i, .ColIndex("CardNO"))))
                    rs2("allocations").value = IIf((.TextMatrix(i, .ColIndex("allocations"))) = "", 0, (.TextMatrix(i, .ColIndex("allocations"))))
                    
                    rs2("QtyDownload").value = IIf((.TextMatrix(i, .ColIndex("QtyDownload"))) = "", Null, val(.TextMatrix(i, .ColIndex("QtyDownload"))))
                    rs2("CardNO2").value = IIf((.TextMatrix(i, .ColIndex("CardNO2"))) = "", Null, (.TextMatrix(i, .ColIndex("CardNO2"))))
                    rs2("QtyDischarge").value = IIf((.TextMatrix(i, .ColIndex("QtyDischarge"))) = "", Null, val(.TextMatrix(i, .ColIndex("QtyDischarge"))))
                    rs2("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                    rs2("BillDate").value = IIf((.TextMatrix(i, .ColIndex("BillDate"))) = "", Null, (.TextMatrix(i, .ColIndex("BillDate"))))
                    rs2("loadingInvoice").value = IIf((.TextMatrix(i, .ColIndex("loadingInvoice"))) = "", Null, (.TextMatrix(i, .ColIndex("loadingInvoice"))))
                    
                End If
                rs2("QtyDownload").value = IIf((.TextMatrix(i, .ColIndex("QtyDownload"))) = "", Null, val(.TextMatrix(i, .ColIndex("QtyDownload"))))
                rs2("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                    rs2("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                rs2("NotesallID").value = val(XPTxtID.text)
                rs2("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                
               rs2("FromPrice").value = val(.TextMatrix(i, .ColIndex("FromPrice")))
               rs2("ToPrice").value = val(.TextMatrix(i, .ColIndex("ToPrice")))
               rs2("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
               rs2("FromCityID").value = val(.TextMatrix(i, .ColIndex("FromCityID")))
               rs2("ToCityID").value = val(.TextMatrix(i, .ColIndex("ToCityID")))
               rs2.update
            Next i
        End With
    End If
End Sub
Sub RetriveGridOrders()
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim i As Integer

FGOrders.Clear flexClearScrollable, flexClearEverything
FGOrders.rows = 1
sql = " SELECT     dbo.TblTravelTransDet.ID, dbo.TblTravelTransDet.NotesallID, dbo.TblTravelTransDet.RentVAlue, dbo.TblTravelTransDet.Selct, dbo.TblTravelTransDet.OrderNo, "
sql = sql & "                       dbo.TblTravelTransDet.TypedID, dbo.TblTravelTransDet.DriverName, dbo.TblTravelTransDet.TypeDriver, dbo.TblTravelTransDet.CustID, dbo.TblCustemers.CusName,"
sql = sql & "                       dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblTravelTransDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2,"
sql = sql & "                       dbo.TblTravelTransDet.PartID , dbo.FixedAssets.Name, dbo.FixedAssets.NameE, TblVendorCars_1.accessory ,dbo.TblTravelTransDet.Price"
sql = sql & "  FROM         dbo.TblTravelTransDet LEFT OUTER JOIN"
sql = sql & "                       dbo.TblVendorCars TblVendorCars_1 ON dbo.TblTravelTransDet.PartID = TblVendorCars_1.ID LEFT OUTER JOIN"
sql = sql & "                       dbo.FixedAssets ON dbo.TblTravelTransDet.PartID = dbo.FixedAssets.id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblVendorCars ON dbo.TblTravelTransDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCarsData ON dbo.TblTravelTransDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.TblTravelTransDet.CustID = dbo.TblCustemers.CusID"
sql = sql & "  WHERE     (dbo.TblTravelTransDet.NotesallID = " & val(XPTxtID.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With FGOrders
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Selct")) = 1
.TextMatrix(i, .ColIndex("LineNo")) = i
.TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs2("OrderNo").value), "", rs2("OrderNo").value)
.TextMatrix(i, .ColIndex("RentVAlue")) = IIf(IsNull(rs2("RentVAlue").value), "", rs2("RentVAlue").value)
.TextMatrix(i, .ColIndex("CustID")) = IIf(IsNull(rs2("CustID").value), "", rs2("CustID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
End If
.TextMatrix(i, .ColIndex("TypedID")) = IIf(IsNull(rs2("TypedID").value), "", rs2("TypedID").value)
.TextMatrix(i, .ColIndex("TypeDriver")) = IIf(IsNull(rs2("TypeDriver").value), "", rs2("TypeDriver").value)
.TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
.TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(rs2("CarID").value), "", rs2("CarID").value)
.TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs2("PartID").value), "", rs2("PartID").value)
If val(.TextMatrix(i, .ColIndex("TypedID"))) = 1 Then
.TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs2("BoardNO").value), "", rs2("BoardNO").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("Name").value), "", rs2("Name").value)
Else
.TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
End If
Else
.TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs2("accessory").value), "", rs2("accessory").value)
.TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs2("BoardNo2").value), "", rs2("BoardNo2").value)
End If
.TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs2("DriverName").value), "", rs2("DriverName").value)
rs2.MoveNext
Next i
End With
End If
ReLineGrid
End Sub
Sub RetriveTypeTransport()

    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim i As Integer

    VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.rows = 1

    sql = " SELECT allocations , dbo.TblTripTypesTransport.ID, dbo.TblTripTypesTransport.NotesallID, dbo.TblTripTypesTransport.CardNO, dbo.TblTripTypesTransport.QtyDownload, dbo.TblTripTypesTransport.loadingInvoice,"
    sql = sql & " dbo.TblTripTypesTransport.CardNO2, dbo.TblTripTypesTransport.QtyDischarge, dbo.TblTripTypesTransport.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,TblUnites.UnitName,TblUnites.UnitNamee,TblUnites.UnitId,"
    sql = sql & " dbo.TblItems.fullcode ,dbo.TblTripTypesTransport.BillDate,"
    sql = sql & " dbo.TblTripTypesTransport.Price , dbo.TblTripTypesTransport.FromPrice, dbo.TblTripTypesTransport.ToPrice ,dbo.TblTripTypesTransport.Price,TblTripTypesTransport.FromCityID,TblTripTypesTransport.ToCityID ,CC2.GovernmentName as ToCity,TblCountriesGovernments.GovernmentName as FromCity"
    sql = sql & " FROM dbo.TblTripTypesTransport LEFT OUTER JOIN"
    sql = sql & " dbo.TblItems ON dbo.TblTripTypesTransport.ItemID = dbo.TblItems.ItemID"
    sql = sql & " LEFT OUTER JOIN dbo.TblUnites ON dbo.TblTripTypesTransport.UnitId = dbo.TblUnites.UnitId"
     sql = sql & "                  LEFT OUTER JOIN TblCountriesGovernments On TblCountriesGovernments.GovernmentID =TblTripTypesTransport.FromCityID "
    sql = sql & "                  LEFT OUTER JOIN TblCountriesGovernments CC2 On CC2.GovernmentID =TblTripTypesTransport.ToCityID "
    sql = sql & " Where (dbo.TblTripTypesTransport.NotesallID = " & val(XPTxtID.text) & ") "
    
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs2.RecordCount > 0 Then
        rs2.MoveFirst
        With VSFlexGrid3
            .rows = .rows + rs2.RecordCount
            For i = 1 To .rows - 1
            
            
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("allocations")) = IIf(IsNull(rs2("allocations").value), 0, rs2("allocations").value)
                
                .TextMatrix(i, .ColIndex("CardNO")) = IIf(IsNull(rs2("CardNO").value), "", rs2("CardNO").value)
                .TextMatrix(i, .ColIndex("QtyDownload")) = IIf(IsNull(rs2("QtyDownload").value), "", rs2("QtyDownload").value)
                .TextMatrix(i, .ColIndex("CardNO2")) = IIf(IsNull(rs2("CardNO2").value), "", rs2("CardNO2").value)
                .TextMatrix(i, .ColIndex("BillDate")) = IIf(IsNull(rs2("BillDate").value), "", rs2("BillDate").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
                Else
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
                End If
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(rs2("UnitId").value), "", rs2("UnitId").value)
                .TextMatrix(i, .ColIndex("QtyDischarge")) = IIf(IsNull(rs2("QtyDischarge").value), "", rs2("QtyDischarge").value)
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs2("Fullcode").value), 0, rs2("Fullcode").value)
                .TextMatrix(i, .ColIndex("loadingInvoice")) = IIf(IsNull(rs2("loadingInvoice").value), "", rs2("loadingInvoice").value)
                
                .TextMatrix(i, .ColIndex("FromCityID")) = IIf(IsNull(rs2("FromCityID").value), "", rs2("FromCityID").value)
                .TextMatrix(i, .ColIndex("ToCityID")) = IIf(IsNull(rs2("ToCityID").value), "", rs2("ToCityID").value)
                .TextMatrix(i, .ColIndex("FromPrice")) = IIf(IsNull(rs2("FromPrice").value), "", rs2("FromPrice").value)
                .TextMatrix(i, .ColIndex("ToPrice")) = IIf(IsNull(rs2("ToPrice").value), "", rs2("ToPrice").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs2("Price").value), "", rs2("Price").value)
                .TextMatrix(i, .ColIndex("FromCity")) = IIf(IsNull(rs2("FromCity").value), "", rs2("FromCity").value)
                .TextMatrix(i, .ColIndex("ToCity")) = IIf(IsNull(rs2("ToCity").value), "", rs2("ToCity").value)
                
                .TextMatrix(i, .ColIndex("TotalValue")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("QtyDownload")))
                rs2.MoveNext
            Next i
        End With
    End If
ReLineGrid
End Sub

Sub saveItemsGrid()
    Dim i As Integer
    Dim rss As ADODB.Recordset
    Set rss = New ADODB.Recordset
    Dim sql_str As String
    Dim StrSQL As String
    If Me.TxtModFlg.text <> "R" Then
        If TxtModFlg.text = "E" Then
            StrSQL = "Delete From TravKItemDet Where NotesID = " & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
          sql_str = "select * from TravKItemDet "
            rss.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    With Me.VSFlexGrid2
        For i = .FixedRows To .rows - 1
            rss.AddNew
            rss("NotesID").value = val(XPTxtID.text)
            rss("ItemID").value = IIf((.TextMatrix(i, .ColIndex("KItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("KItemID"))))
            rss("Count").value = IIf((.TextMatrix(i, .ColIndex("Count"))) = "", Null, val((.TextMatrix(i, .ColIndex("Count")))))
            rss("UnitID").value = IIf((.TextMatrix(i, .ColIndex("KUnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("KUnitID"))))
            rss.update
        Next i
    End With
  End If
End Sub
Sub fillItemsGrid()

    Dim i As Integer
    Dim rs_ItemsGrid As ADODB.Recordset
    Set rs_ItemsGrid = New ADODB.Recordset
    Dim StrSQL As String
        
    StrSQL = " SELECT     dbo.TravKItemDet.NotesID, dbo.TravKItemDet.[Count], dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TravKItemDet.ItemID, "
    StrSQL = StrSQL & "                  dbo.TravKItemDet.UnitID , dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & "     FROM         dbo.TravKItemDet INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.TravKItemDet.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.TravKItemDet.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where dbo.TravKItemDet.NotesID = " & val(XPTxtID.text)
    
    rs_ItemsGrid.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid2.rows = 1
    
    If rs_ItemsGrid.RecordCount > 0 Then
        rs_ItemsGrid.MoveFirst
        With VSFlexGrid2
            .rows = rs_ItemsGrid.RecordCount + 1
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("KItemID")) = IIf(IsNull(rs_ItemsGrid("ItemID").value), 0, rs_ItemsGrid("ItemID").value)
                .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(rs_ItemsGrid("Count").value), 0, rs_ItemsGrid("Count").value)
                .TextMatrix(i, .ColIndex("KUnitID")) = IIf(IsNull(rs_ItemsGrid("UnitID").value), 0, rs_ItemsGrid("UnitID").value)
               ' .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_ItemsGrid("namee").value), "", rs_ItemsGrid("namee").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("KItem")) = IIf(IsNull(rs_ItemsGrid("ItemName").value), "", rs_ItemsGrid("ItemName").value)
                    .TextMatrix(i, .ColIndex("KUnit")) = IIf(IsNull(rs_ItemsGrid("UnitName").value), "", rs_ItemsGrid("UnitName").value)
                Else
                    .TextMatrix(i, .ColIndex("KItem")) = IIf(IsNull(rs_ItemsGrid("ItemNamee").value), "", rs_ItemsGrid("ItemNamee").value)
                    .TextMatrix(i, .ColIndex("KUnit")) = IIf(IsNull(rs_ItemsGrid("UnitNamee").value), "", rs_ItemsGrid("UnitNamee").value)
                End If
                rs_ItemsGrid.MoveNext
            Next
        End With
    End If
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

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label15.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label15.Caption = "Approved"
                                 End If
                            Label15.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label15.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label15.Caption = "Currently required Approve"
                            End If
                 Label15.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.rows = 1
    End If
RsDetails.Close

End Function



