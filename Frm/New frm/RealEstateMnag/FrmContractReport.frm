VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmContractReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "FrmContractReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ãÓÍ"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ýì ÇáÝÊÑÉ"
      Height          =   1185
      Left            =   4320
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830593
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830593
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ãä"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Åáì"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   6405
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   10515
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÍÏÏ ÝäÑÉ äåÇíÉ ÇáÚÞÏ"
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   4560
         Width           =   6555
         Begin MSComCtl2.DTPicker STARTDATE 
            Height          =   330
            Left            =   3495
            TabIndex        =   53
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94830593
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal STARTDATEH 
            Height          =   330
            Left            =   3480
            TabIndex        =   54
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal ENDDATEH 
            Height          =   330
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker ENDDATE 
            Height          =   330
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94830593
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÇáì"
            Height          =   435
            Index           =   6
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ãä"
            Height          =   315
            Index           =   4
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   480
            Width           =   945
         End
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
         Left            =   4320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1560
         Width           =   855
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
         Left            =   4320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Çáßá"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÚÞæÏ ÇáÊí áã  ÊÕÝì"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇáÚÞæÏ ÇáÊí Êã ÊÕÝíÊåÇ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   -480
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtContNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   4935
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
         Left            =   16080
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtDes 
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
         Height          =   435
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox TxtEmployeeID 
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
         Left            =   16080
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   24
         Top             =   1920
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÍÏÏ ÝÊÑÉ ÈÏÇíÉ ÇáÚÞÏ"
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3480
         Width           =   6555
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   18
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94830593
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   19
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94830593
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ãä"
            Height          =   315
            Index           =   3
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÇáì"
            Height          =   435
            Index           =   14
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   6960
         TabIndex        =   14
         Top             =   0
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2895
            Left            =   120
            Picture         =   "FrmContractReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3300
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÓÇÊÑíÉ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   4695
            Left            =   120
            TabIndex        =   15
            Top             =   3120
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Index           =   0
         Left            =   11640
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Tag             =   "ÚÝæÇ íÑÌì ÇÏÎÇá ÃÓã ÇáÍí"
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   12000
         TabIndex        =   30
         Tag             =   "ÚÝæÇ íÑÌì ÇÎÊíÇÑ ÃÓã ÇáãÓÊÇÌÑ"
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpaqar 
         Height          =   315
         Left            =   12000
         TabIndex        =   33
         Tag             =   "ÚÝæÇ íÑÌì ÇÎÊíÇÑ ÃÓã ÇáãÓÊÇÌÑ"
         Top             =   2280
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranches 
         Height          =   315
         Left            =   240
         TabIndex        =   38
         Tag             =   "ÚÝæÇ íÑÌì ÇÏÎÇá ÃÓã ÇáÍí"
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   240
         TabIndex        =   40
         Tag             =   "ÚÝæÇ íÑÌì ÇÏÎÇá ÃÓã ÇáÍí"
         Top             =   2640
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   240
         TabIndex        =   46
         Tag             =   "ÚÝæÇ íÑÌì ÇÎÊíÇÑÃÓã ÇáãÇáß"
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcCustomer 
         Height          =   315
         Left            =   240
         TabIndex        =   49
         Tag             =   "ÚÝæÇ íÑÌì ÇÎÊíÇÑ ÃÓã ÇáãÓÊÇÌÑ"
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ØÈÞÇ áÍÇáÉ ÇáÚÞÏ"
         Height          =   195
         Left            =   5280
         TabIndex        =   51
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ØÈÞÇ  ááãÓÊÃÌÑ"
         Height          =   195
         Index           =   7
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ ááãÇáß"
         Height          =   195
         Index           =   2
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ØÈÞÇ áÑÞã ÇáÚÞÏ"
         Height          =   195
         Left            =   5280
         TabIndex        =   39
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ØÈÞÇ áÝÑÚ ãÚíä"
         Height          =   195
         Index           =   6
         Left            =   5565
         TabIndex        =   37
         Top             =   240
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   735
         Left            =   120
         Top             =   5640
         Width           =   6855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "íÑÌì ÇÎÊíÇÑ ÇáÊÇÑíÎ Çæ ÓæÝ íßæä ÇáÊÞÑíÑ ÇÌãÇáí áßá æÇáãÏÉ"
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
         Height          =   690
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   5640
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ áÔÑæØ ÎÇÕÉ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3000
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ áãÍÕá ÚÞÇÑ ãÍÏÏ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   16800
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ áãÓæÞ ÚÞÏ ãÍÏÏ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   17040
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ áæÍÏÉ ãÍÏÏå"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ØÈÞÇ áäæÚ ãÍÏÏ"
         Height          =   195
         Index           =   15
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ØÈÞÇ áÚÞÇÑ ãÚíä"
         Height          =   195
         Index           =   1
         Left            =   5565
         TabIndex        =   8
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ØÈÞÇ áÝÑÚ ãÚíä"
         Height          =   195
         Index           =   0
         Left            =   12795
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   7200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   7200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "ÎÑæÌ"
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   480
      Picture         =   "FrmContractReport.frx":10A48
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÔÇÔÉ ÊÞÇÑíÑ  ÇáÚÞæÏ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   -60
      TabIndex        =   6
      Top             =   0
      Width           =   10515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmContractReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
'

Private Sub btnClear_Click()
Cmd_Click (1)
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

            GetData
        Case 1
   
            clear_all Me
           
         Fromdate.value = ""
    ToDate.value = ""
             StartDate.value = ""
    EndDate.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "äÊíÌÉ ÇáÈÍË"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub










Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
DcbUnitType_Click (0)
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 
    GetIqarCode , , dcbAqarType.BoundText, EmpCode
    
    Me.TxtSearch.Text = EmpCode
End Sub

Private Sub DcbUnitType_Change()
DcbUnitType_Click (0)
End Sub

Private Sub DcbUnitType_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(dcbAqarType.BoundText) > 0 Then
idd = val(dcbAqarType.BoundText)

idd1 = val(DcbUnitType.BoundText)
'If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
'Else
'Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
'End If
End If
End Sub

Private Sub dcCustomer_Change()
If val(dcCustomer.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode
    Me.Text15.Text = EmpCode
End Sub

Private Sub dcCustomer_Click(Area As Integer)
dcCustomer_Change
End Sub

Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode, 57
    Me.Text3.Text = EmpCode
End Sub



Private Sub ENDDATE_Change()
If Not IsNull(EndDate.value) Then
   EndDateH.value = ToHijriDate(EndDate.value)
   End If
End Sub

Private Sub ENDDATEH_LostFocus()
 VBA.Calendar = vbCalGreg
            EndDate.value = ToGregorianDate(EndDateH.value)
End Sub


Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches DcbBranches
    Dcombos.GetIqar dcbAqarType
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Fromdate.value = ""
    ToDate.value = ""
    Cmd_Click (1)
    
    Set cSearch = New clsDCboSearch
  '  My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
   ' RsSavRec.CursorLocation = adUseClient
   ' RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Resize_Form Me
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
   

 StrSQL = "SELECT     dbo.TblContract.ContNo, dbo.TblContract.ContType, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, "
 StrSQL = StrSQL & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS Nameunitno,"
 StrSQL = StrSQL & "                     dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContract.StrDate, dbo.TblContract.EndDate,dbo.TblContract.NoteSerial1,"
 StrSQL = StrSQL & "                     dbo.TblContract.OthersRules, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblContract.CusID,"
 StrSQL = StrSQL & "                     TblCustemers_1.CusName AS RenCusName, TblCustemers_1.CusNamee AS RenCusNameE, TblCustemers_1.Fullcode AS RenFullcode, dbo.TblContract.FromdateH,"
 StrSQL = StrSQL & "                     dbo.TblContract.todateH , dbo.TblContract.EndContract"
 StrSQL = StrSQL & "  FROM         dbo.TblContract LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers TblCustemers_1 ON dbo.TblContract.CusID = TblCustemers_1.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.TblContract.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
StrSQL = StrSQL & " where 1=1"

    StrWhere = ""
        
If Me.TxtDes.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.OthersRules like '%" & Me.TxtDes.Text & "%'"

End If
If val(Me.DcbBranches.BoundText) <> 0 Or Me.DcbBranches.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Branch_NO = " & val(DcbBranches.BoundText)
End If
If val(Me.dcsupplier.BoundText) <> 0 Or Me.dcsupplier.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.ownerid= " & val(dcsupplier.BoundText)

End If
If val(Me.dcCustomer.BoundText) <> 0 Or Me.dcCustomer.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.CusID= " & val(dcCustomer.BoundText)

End If
If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Iqar= " & val(dcbAqarType.BoundText)

End If
If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitType= " & val(DcbUnitType.BoundText)

End If
If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitNo= " & val(DcbUnitNo.BoundText)

End If
If Me.TxtContNo.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.NoteSerial= N'" & TxtContNo.Text & " '"

End If
If Opt(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblContract.EndContract= 1 "
End If
If Opt(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContract.EndContract IS NULL "
End If

If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContract.StrDate >=" & SQLDate(Me.Fromdate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContract.StrDate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If
   If Not IsNull(Me.StartDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContract.EndDate >=" & SQLDate(Me.StartDate.value, True) & ""
      End If

    If Not IsNull(Me.EndDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContract.EndDate <=" & SQLDate(Me.EndDate.value, True) & ""
     
    End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL & " order by  dbo.TblContract.NoteSerial1 "

  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
         
        End If

        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ ÊæÇÝÞ ÔÑæØ ÇáÊÞÑíÑ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 
 rs.MoveFirst

 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "äÊíÌÉ ÇáÈÍË=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaContructReport.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaContructReport.rpt"
            
       End If
  


    If Dir(StrFileName) = "" Then
       
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
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
       
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If
  Dim Total As String
  Dim totl As Double

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

Private Sub FromDate_Change()
If Not IsNull(Fromdate.value) Then
   FromDateH.value = ToHijriDate(Fromdate.value)
   End If
End Sub
Private Sub Fromdateh_LostFocus()

 VBA.Calendar = vbCalGreg
            Fromdate.value = ToGregorianDate(FromDateH.value)
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub STARTDATE_Change()
If Not IsNull(StartDate.value) Then
   StartDateh.value = ToHijriDate(StartDate.value)
   End If
End Sub

Private Sub STARTDATEH_LostFocus()
 VBA.Calendar = vbCalGreg
            StartDate.value = ToGregorianDate(StartDateh.value)
End Sub

Private Sub Text1_Change()
DcboEmpaqar.BoundText = GeTEmpIDByEmpCode(Me.Text1.Text, True)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text3.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub

Private Sub ToDate_Change()
If Not IsNull(ToDate.value) Then
   todateH.value = ToHijriDate(ToDate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()
 VBA.Calendar = vbCalGreg
            ToDate.value = ToGregorianDate(todateH.value)
End Sub
Private Sub TxtEmployeeID_Change()
DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.Text, True)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
