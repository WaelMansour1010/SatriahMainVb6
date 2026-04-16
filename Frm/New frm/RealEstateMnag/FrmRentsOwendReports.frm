VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmRentsOwendReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   Icon            =   "FrmRentsOwendReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   5880
      TabIndex        =   40
      Top             =   8190
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   9
      Top             =   8970
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
         Format          =   178520065
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
         Format          =   178520065
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
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
         Caption         =   "≈·Ï"
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
      Height          =   7425
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   10905
      Begin VB.Frame Frame9 
         Caption         =   "‘∆Ê‰ ﬁ«‰Ê‰Ì…"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   1290
         Width           =   2775
         Begin VB.OptionButton legal 
            Alignment       =   1  'Right Justify
            Caption         =   "‰⁄„"
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   80
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton legal 
            Alignment       =   1  'Right Justify
            Caption         =   "·« "
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   79
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton legal 
            Caption         =   "«·ﬂ·"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ ‰Ê⁄ «·„ÿ«·»…"
         Height          =   615
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   4470
         Width           =   6615
         Begin VB.ComboBox DcbStatusOper 
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   180
            Width           =   1065
         End
         Begin XtremeSuiteControls.RadioButton TypedID 
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   65
            Top             =   240
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " ’›Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton TypedID 
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   66
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "›Ê« Ì— ﬂÂ—»«¡"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton TypedID 
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   67
            Top             =   240
            Width           =   615
            _Version        =   786432
            _ExtentX        =   1085
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ’›Ì…"
            Height          =   255
            Index           =   7
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ ‰Ê⁄ «· ﬁ—Ì—"
         Height          =   1185
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   0
         Width           =   7335
         Begin VB.CheckBox chkIsLegalAffairsNo 
            Alignment       =   1  'Right Justify
            Caption         =   "«ŸÂ«— «· ’›Ì«  «·€Ì— „Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì… ›ﬁÿ"
            Height          =   255
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   210
            Width           =   3555
         End
         Begin VB.CheckBox chkIsLegalAffairs 
            Alignment       =   1  'Right Justify
            Caption         =   "«ŸÂ«— «· ’›Ì«  «·„Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì… ›ﬁÿ"
            Height          =   255
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   210
            Width           =   3285
         End
         Begin VB.OptionButton OptRep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— ⁄„— «·œÌ‰"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   5100
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   930
            Width           =   2055
         End
         Begin VB.Frame Frame7 
            Caption         =   "«· ÊÀÌﬁ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   780
            Width           =   2775
            Begin VB.OptionButton Accredit 
               Caption         =   "«·ﬂ·"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   76
               Top             =   120
               Width           =   615
            End
            Begin VB.OptionButton Accredit 
               Alignment       =   1  'Right Justify
               Caption         =   "·„ Ì „"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   75
               Top             =   120
               Width           =   735
            End
            Begin VB.OptionButton Accredit 
               Alignment       =   1  'Right Justify
               Caption         =   " „"
               ForeColor       =   &H00004000&
               Height          =   195
               Index           =   0
               Left            =   1560
               TabIndex        =   74
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.OptionButton OptRep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— «·«ÌÃ«—«  «·„” Õﬁ…Ê«· ’›Ì« "
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   2
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   690
            Width           =   2655
         End
         Begin VB.OptionButton OptRep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— «· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   1
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   540
            Width           =   2535
         End
         Begin VB.OptionButton OptRep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— «·«ÌÃ«—«  «·„” Õﬁ…"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   0
            Left            =   5100
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   450
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· — Ì»"
         Height          =   480
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1170
         Width           =   3795
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " — Ì» ÿ»ﬁ« ··ÊÕœ…"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   1
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   1515
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " — Ì» ÿ»ﬁ« ·· √—ÌŒ"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1200
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   5040
         Width           =   6675
         Begin VB.TextBox tXtD 
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
            Left            =   480
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtDay 
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
            Left            =   480
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   720
            Width           =   735
         End
         Begin XtremeSuiteControls.RadioButton paym 
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   48
            Top             =   720
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»œÊ‰ œ›⁄Â «Ê·Ï "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton paym 
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   49
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "œ›⁄Â «Ê·Ï ›ﬁÿ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton paym 
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   50
            Top             =   720
            Width           =   2295
            _Version        =   786432
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "»œÊ‰ «·œ›⁄Â «·«Ê·Ï «·«ﬁ· „‰ "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton paym 
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   54
            Top             =   240
            Width           =   2295
            _Version        =   786432
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "œ›⁄Â «Ê·Ï «ﬁ· „‰ "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton paym 
            Height          =   255
            Index           =   4
            Left            =   5400
            TabIndex        =   56
            Top             =   480
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ﬂ· «·œ›⁄« "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÌÊ„"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   -600
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÌÊ„"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   -600
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   720
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   4110
         Width           =   3375
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   42
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   43
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   44
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "="
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   45
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<="
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">="
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
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
         Left            =   4320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   3810
         Width           =   855
      End
      Begin VB.TextBox TxtAmount 
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
         Left            =   3720
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   4170
         Width           =   1455
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
         Left            =   4320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3450
         Width           =   855
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   26
         Top             =   2370
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1080
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   6270
         Width           =   6675
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   19
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   222822401
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   20
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   222822401
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   3
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   7470
         TabIndex        =   16
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   4095
            Left            =   120
            Picture         =   "FrmRentsOwendReports.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3300
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
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
            Height          =   5295
            Left            =   120
            TabIndex        =   17
            Top             =   4200
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Top             =   7380
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Top             =   7380
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1650
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
         Top             =   2370
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3090
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   2730
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   3450
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpaqar 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   3810
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcaqartypeid 
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
         Top             =   2010
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·⁄ﬁ«—"
         Height          =   195
         Index           =   16
         Left            =   5835
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   2010
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„Õ’· ⁄ﬁ«— „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3810
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„»·€ „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5505
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   4170
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„”Êﬁ ⁄ﬁœ „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3450
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·ÊÕœ… „ÕœœÂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3090
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·‰Ê⁄ „Õœœ"
         Height          =   195
         Index           =   15
         Left            =   5745
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2730
         Width           =   1110
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   7410
         Width           =   6975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
         Height          =   450
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   7380
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5790
         TabIndex        =   8
         Top             =   2370
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   0
         Left            =   5835
         TabIndex        =   5
         Top             =   1650
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   8190
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—  ›’Ì·Ì"
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
      Left            =   120
      TabIndex        =   1
      Top             =   8190
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
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
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   39
      Top             =   8190
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   57
      Top             =   8190
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷  Ã„Ì⁄Ì ‘Â—Ì"
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   720
      Picture         =   "FrmRentsOwendReports.frx":10A48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·«ÌÃ«—«  «·„” Õﬁ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   -105
      TabIndex        =   6
      Top             =   0
      Width           =   10560
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
Attribute VB_Name = "FrmRentsOwendReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public indexx As Integer

Private Sub btnClear_Click()
Cmd_Click (7)
End Sub



Private Sub Cmd_Click(index As Integer)

    Select Case index


        Case 0
       indexx = 1
If OptRep(0).value = True Then
 GetData

 Else
 GetDataExpEleec
 End If
        Case 1
   
       indexx = 0
      If OptRep(0).value = True Then
 GetData
  ElseIf OptRep(2).value = True Then
        indexx = 3
GetData101
ElseIf OptRep(3).value = True Then
GetData0003
 Else
 GetDataExpEleec
 End If
   Case 3
       indexx = 2
 GetData
  Case 7
  clear_all Me
  OptRep(0).value = True
  FromDate.value = ""
    ToDate.value = ""
TypedID(2).value = True
        Case 2
        
            Unload Me
            Case 3
'print_report
    End Select

End Sub
Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
DcbUnitType_Change
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , dcbAqarType.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.text = EmpCode
End Sub

Private Sub DcboEmp_Change()
 If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
End Sub

Private Sub DcboEmpaqar_Change()
 If val(Me.DcboEmpaqar.BoundText) = 0 Then Exit Sub
           Me.Text1.text = get_EMPLOYEE_Data(val(Me.DcboEmpaqar.BoundText), "Fullcode")
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(dcbAqarType.BoundText) > 0 Then
idd = val(dcbAqarType.BoundText)

idd1 = val(DcbUnitType.BoundText)

Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"

End If

End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Sub HideElemint()
If OptRep(0).value = True Or OptRep(3).value = True Then
Frame2.Visible = True
Cmd(0).Enabled = True
Cmd(3).Enabled = True
Frame6.Visible = False
Else
Frame6.Visible = True
Frame2.Visible = False
Cmd(0).Enabled = False
Cmd(3).Enabled = False
End If
End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    FromDate.value = Date
    ToDate.value = Date
    Set Dcombos = New ClsDataCombos
    OptRep(0).value = True
    Dcombos.GetIqar dcbAqarType
    
   ' Dcombos.GetAlarm Me.DcbAlarm
   Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetSalesRepData Me.DcboEmpaqar
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.getAkarType Me.dcaqartypeid
    'Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    Dcombos.GetBranches DcbBranch
With DcbStatusOper
.Clear
.AddItem "⁄«œÌ"
.AddItem "Â—Ê»"
End With

    Set cSearch = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
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
 StrSQL = "SELECT   LegalIssue,Accredit,    dbo.TblContractInstallments.installValue, ISNULL(dbo.TblContractInstallments.installValue, 0) - ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) "
 StrSQL = StrSQL & "                     AS remains1, ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) AS collectedValue, dbo.TblContractInstallments.ContNo,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.InstalldateH, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblContract.UnitNo,"
 StrSQL = StrSQL & "                     dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.NoteID, dbo.TblContractInstallments.RentValue, dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric, dbo.TblContractInstallments.TelandNet, dbo.TblContractInstallments.payed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Remains AS Expr1, dbo.TblContractInstallments.RentValuePayed, dbo.TblContractInstallments.CommissionsPayed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed, dbo.TblContractInstallments.ElectricPayed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.TelandNetPayed, dbo.TblContractInstallments.lastPayedDate, dbo.TblContractInstallments.lastPayedDateH,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Doneofall, dbo.TblContractInstallments.hijri,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue, dbo.TblContractInstallments.des,"
 StrSQL = StrSQL & "                     dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContractInstallments.id, dbo.TblContract.Emp_ID, TblEmployee_1.Emp_Name,"
 StrSQL = StrSQL & "                     TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblAqar.SalesEmp, TblEmployee_1.Emp_Name AS Emp_NameAqr,"
 StrSQL = StrSQL & "                     TblEmployee_1.Fullcode AS FullcodeAqar, TblEmployee_1.Emp_Namee AS Emp_NameAqrE, dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1,"
 StrSQL = StrSQL & "                     dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblCustemers.Cus_Phone,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.Cus_mobile, dbo.TblContract.StrMerg, DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) AS mon,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstallNo, dbo.TblAqar.BranchId, dbo.TblContract.EndContract, dbo.TblContract.FromdateH,"
 StrSQL = StrSQL & "                     dbo.TblContract.TodateH, dbo.TblContract.RecorddateH, dbo.TblContract.ContType, dbo.TblContract.ContNo AS Expr2, dbo.TblContract.RentType,"
 StrSQL = StrSQL & "                     dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini,"
 StrSQL = StrSQL & "                     dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water AS WaterCont, dbo.TblContract.Electricity, dbo.TblContract.Phone,"
 StrSQL = StrSQL & "                     dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate,"
 StrSQL = StrSQL & "                     dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing, dbo.TblContract.Remarks, dbo.TblContract.FirstInstallDateH,"
 StrSQL = StrSQL & "                     dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblContract.OutContract, dbo.TblContract.OldRent, dbo.TblContract.OldWater,"
 StrSQL = StrSQL & "                     dbo.TblContract.OldElectric, dbo.TblContract.oldCommi, dbo.TblContract.DivWater, dbo.TblContract.DivElectric, dbo.TblContract.OldInsurance,"
 StrSQL = StrSQL & "                     dbo.TblContract.balanceDate, dbo.TblContract.balanceDateH, dbo.TblContract.balanceDes, dbo.TblContract.Renew, dbo.TblContract.ContNoOld,"
 StrSQL = StrSQL & "                     dbo.TblContract.FromdateHO, dbo.TblContract.FromdateO, dbo.TblContract.Employeecontract, dbo.TblContract.Emp_IDContract, dbo.TblContract.OutOffice,"
 StrSQL = StrSQL & "                     dbo.TblContract.LegalIssue, dbo.TblContract.NotValue, dbo.TblContract.UnitElectric, dbo.TblContract.RetValue2, dbo.TblContract.WaterValue2,"
 StrSQL = StrSQL & "                     dbo.TblContract.CommValue2,dbo.TblContract.InstrunceValue2, dbo.TblContract.ElectricityValue2, dbo.TblContract.MiniRentValue, dbo.TblContract.Servce, dbo.TblContract.MethodDeci,"
 StrSQL = StrSQL & "                     dbo.TblContract.FlagContrNew2"
 StrSQL = StrSQL & "  FROM         dbo.TblAqarDetai RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblContract ON dbo.TblCustemers.CusID = dbo.TblContract.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.TblContract.Emp_ID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id ON dbo.TblAqarDetai.Id = dbo.TblContract.UnitNo LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqar ON TblEmployee_1.Emp_ID = dbo.TblAqar.SalesEmp ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo"
  StrSQL = StrSQL & " WHERE    (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL) and ((dbo.TblContract.EndContract IS NULL) OR (dbo.TblContract.EndContract=0))  "


    BolBegine = False
    StrWhere = ""
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Branch_NO = " & val(Me.DcbBranch.BoundText)

End If
If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(Me.dcaqartypeid.BoundText)
End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.text <> "" Then

StrWhere = StrWhere & " AND dbo.TblContract.Iqar = " & val(Me.dcbAqarType.BoundText)

End If

If val(Me.DcboEmpaqar.BoundText) <> 0 Or Me.DcboEmpaqar.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.SalesEmp   = " & val(DcboEmpaqar.BoundText)

End If

If val(Me.DcboEmp.BoundText) <> 0 Or Me.DcboEmp.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Emp_ID    = " & val(DcboEmp.BoundText)

End If

If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitType  = " & val(DcbUnitType.BoundText)

End If
If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitNo  = " & val(DcbUnitNo.BoundText)

End If
If val(TxtAmount.text) <> 0 Then

If opt2(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  > " & val(TxtAmount.text)
ElseIf opt2(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  < " & val(TxtAmount.text)
ElseIf opt2(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  <= " & val(TxtAmount.text)
ElseIf opt2(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  >= " & val(TxtAmount.text)
ElseIf opt2(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  = " & val(TxtAmount.text)
End If
End If
If paym(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo = 1 "
End If
If paym(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo <> 1"
End If
If paym(2).value = True Then
StrWhere = StrWhere & " AND ( (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) OR"
StrWhere = StrWhere & " (dbo.TblContractInstallments.InstallNo > 1)) "
End If
If paym(3).value = True Then
StrWhere = StrWhere & " AND (dbo.TblContractInstallments.InstallNo = 1) and  (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) "
End If

   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContractInstallments.Installdate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContractInstallments.Installdate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If

 

If Accredit(0).value = True Then
StrWhere = StrWhere & " AND   Accredit=1"
ElseIf Accredit(1).value = True Then
StrWhere = StrWhere & " AND   (Accredit is null or Accredit=0)"
End If

If legal(0).value = True Then
StrWhere = StrWhere & " AND   LegalIssue=1"
ElseIf legal(1).value = True Then
StrWhere = StrWhere & " AND   (LegalIssue is null or LegalIssue=0)"
End If


 


    StrSQL = StrSQL & StrWhere
 If opt(1).value = True Then

 StrSQL = StrSQL & " order by   dbo.TblContract.unitno "
 ElseIf opt(0).value = True Then
 StrSQL = StrSQL & " order by  dbo.TblContractInstallments.Installdate "
 Else
  StrSQL = StrSQL & " order by  dbo.TblContractInstallments.ID "
  End If
    Set rs = New ADODB.Recordset
    Cn.CommandTimeout = 10000
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
    End If

End Sub

Public Sub GetData0003()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  
StrSQL = " SELECT "
StrSQL = StrSQL & "   account_serial="
StrSQL = StrSQL & "    ("
StrSQL = StrSQL & "    SELECT     Account_Serial"
StrSQL = StrSQL & "     FROM         ACCOUNTS A"
StrSQL = StrSQL & "     Where a.Account_code = dbo.TblCustemers.Account_code"
StrSQL = StrSQL & "     ),"
StrSQL = StrSQL & "        OpeningBalance=("
StrSQL = StrSQL & "        SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
StrSQL = StrSQL & "                                  From"
StrSQL = StrSQL & "        "
StrSQL = StrSQL & "        ("
StrSQL = StrSQL & "        SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
StrSQL = StrSQL & "                                                                                 DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
                                                   StrSQL = StrSQL & "        FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
StrSQL = StrSQL & "        WHERE  do.Account_Code = dbo.TblCustemers.Account_Code   and(do.Posted IS NULL)"

StrSQL = StrSQL & "        )x"

StrSQL = StrSQL & "        )"
StrSQL = StrSQL & "        ,"
  StrSQL = StrSQL & "         LegalIssue,Accredit,    dbo.TblContractInstallments.installValue, ISNULL(dbo.TblContractInstallments.installValue, 0) - ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) "
 StrSQL = StrSQL & "                     AS remains1, ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) AS collectedValue, dbo.TblContractInstallments.ContNo,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.InstalldateH, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblContract.UnitNo,"
 StrSQL = StrSQL & "                     dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.NoteID, dbo.TblContractInstallments.RentValue, dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric, dbo.TblContractInstallments.TelandNet, dbo.TblContractInstallments.payed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Remains AS Expr1, dbo.TblContractInstallments.RentValuePayed, dbo.TblContractInstallments.CommissionsPayed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed, dbo.TblContractInstallments.ElectricPayed,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.TelandNetPayed, dbo.TblContractInstallments.lastPayedDate, dbo.TblContractInstallments.lastPayedDateH,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Doneofall, dbo.TblContractInstallments.hijri,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue, dbo.TblContractInstallments.des,"
 StrSQL = StrSQL & "                     dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContractInstallments.id, dbo.TblContract.Emp_ID, TblEmployee_1.Emp_Name,"
 StrSQL = StrSQL & "                     TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblAqar.SalesEmp, TblEmployee_1.Emp_Name AS Emp_NameAqr,"
 StrSQL = StrSQL & "                     TblEmployee_1.Fullcode AS FullcodeAqar, TblEmployee_1.Emp_Namee AS Emp_NameAqrE, dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1,"
 StrSQL = StrSQL & "                     dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblCustemers.Cus_Phone,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.Cus_mobile, dbo.TblContract.StrMerg, DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) AS mon,"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstallNo, dbo.TblAqar.BranchId, dbo.TblContract.EndContract, dbo.TblContract.FromdateH,"
 StrSQL = StrSQL & "                     dbo.TblContract.TodateH, dbo.TblContract.RecorddateH, dbo.TblContract.ContType, dbo.TblContract.ContNo AS Expr2, dbo.TblContract.RentType,"
 StrSQL = StrSQL & "                     dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini,"
 StrSQL = StrSQL & "                     dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water AS WaterCont, dbo.TblContract.Electricity, dbo.TblContract.Phone,"
 StrSQL = StrSQL & "                     dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate,"
 StrSQL = StrSQL & "                     dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing, dbo.TblContract.Remarks, dbo.TblContract.FirstInstallDateH,"
 StrSQL = StrSQL & "                     dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblContract.OutContract, dbo.TblContract.OldRent, dbo.TblContract.OldWater,"
 StrSQL = StrSQL & "                     dbo.TblContract.OldElectric, dbo.TblContract.oldCommi, dbo.TblContract.DivWater, dbo.TblContract.DivElectric, dbo.TblContract.OldInsurance,"
 StrSQL = StrSQL & "                     dbo.TblContract.balanceDate, dbo.TblContract.balanceDateH, dbo.TblContract.balanceDes, dbo.TblContract.Renew, dbo.TblContract.ContNoOld,"
 StrSQL = StrSQL & "                     dbo.TblContract.FromdateHO, dbo.TblContract.FromdateO, dbo.TblContract.Employeecontract, dbo.TblContract.Emp_IDContract, dbo.TblContract.OutOffice,"
 StrSQL = StrSQL & "                     dbo.TblContract.LegalIssue, dbo.TblContract.NotValue, dbo.TblContract.UnitElectric, dbo.TblContract.RetValue2, dbo.TblContract.WaterValue2,"
 StrSQL = StrSQL & "                     dbo.TblContract.CommValue2,dbo.TblContract.InstrunceValue2, dbo.TblContract.ElectricityValue2, dbo.TblContract.MiniRentValue, dbo.TblContract.Servce, dbo.TblContract.MethodDeci,"
 StrSQL = StrSQL & "                     dbo.TblContract.FlagContrNew2"
 StrSQL = StrSQL & "  FROM         dbo.TblAqarDetai RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblContract ON dbo.TblCustemers.CusID = dbo.TblContract.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.TblContract.Emp_ID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id ON dbo.TblAqarDetai.Id = dbo.TblContract.UnitNo LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqar ON TblEmployee_1.Emp_ID = dbo.TblAqar.SalesEmp ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo"
  StrSQL = StrSQL & " WHERE    (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL) and ((dbo.TblContract.EndContract IS NULL) OR (dbo.TblContract.EndContract=0))  "


    BolBegine = False
    StrWhere = ""
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Branch_NO = " & val(Me.DcbBranch.BoundText)

End If
If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(Me.dcaqartypeid.BoundText)
End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.text <> "" Then

StrWhere = StrWhere & " AND dbo.TblContract.Iqar = " & val(Me.dcbAqarType.BoundText)

End If

If val(Me.DcboEmpaqar.BoundText) <> 0 Or Me.DcboEmpaqar.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.SalesEmp   = " & val(DcboEmpaqar.BoundText)

End If

If val(Me.DcboEmp.BoundText) <> 0 Or Me.DcboEmp.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Emp_ID    = " & val(DcboEmp.BoundText)

End If

If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitType  = " & val(DcbUnitType.BoundText)

End If
If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitNo  = " & val(DcbUnitNo.BoundText)

End If
If val(TxtAmount.text) <> 0 Then

If opt2(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  > " & val(TxtAmount.text)
ElseIf opt2(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  < " & val(TxtAmount.text)
ElseIf opt2(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  <= " & val(TxtAmount.text)
ElseIf opt2(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  >= " & val(TxtAmount.text)
ElseIf opt2(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  = " & val(TxtAmount.text)
End If
End If
If paym(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo = 1 "
End If
If paym(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo <> 1"
End If
If paym(2).value = True Then
StrWhere = StrWhere & " AND ( (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) OR"
StrWhere = StrWhere & " (dbo.TblContractInstallments.InstallNo > 1)) "
End If
If paym(3).value = True Then
StrWhere = StrWhere & " AND (dbo.TblContractInstallments.InstallNo = 1) and  (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) "
End If

   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContractInstallments.Installdate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContractInstallments.Installdate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If

 

If Accredit(0).value = True Then
StrWhere = StrWhere & " AND   Accredit=1"
ElseIf Accredit(1).value = True Then
StrWhere = StrWhere & " AND   (Accredit is null or Accredit=0)"
End If

If legal(0).value = True Then
StrWhere = StrWhere & " AND   LegalIssue=1"
ElseIf legal(1).value = True Then
StrWhere = StrWhere & " AND   (LegalIssue is null or LegalIssue=0)"
End If


 


    StrSQL = StrSQL & StrWhere
 If opt(1).value = True Then

 StrSQL = StrSQL & " order by   dbo.TblContract.unitno "
 ElseIf opt(0).value = True Then
 StrSQL = StrSQL & " order by  dbo.TblContractInstallments.Installdate "
 Else
  StrSQL = StrSQL & " order by  dbo.TblContractInstallments.ID "
  End If
    Set rs = New ADODB.Recordset
    Cn.CommandTimeout = 10000
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL, 30
    End If

End Sub

Public Sub GetData101()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    If IsNull(Me.ToDate.value) Or IsNull(Me.FromDate.value) Then
    MsgBox ("Ì—ÃÏ  ÕœÌœ «·› —…")
    Exit Sub
    End If
 StrSQL = " SELECT      dbo.TblContractInstallments.installValue, ISNULL(dbo.TblContractInstallments.installValue, 0) "
  StrSQL = StrSQL & "                    - ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) AS remains1, ISNULL(dbo.InstallmentValue(dbo.TblContractInstallments.id), 0) AS collectedValue,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.ContNo, dbo.TblContractInstallments.InstalldateH, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname,"
  StrSQL = StrSQL & "                    dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name,"
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_namee, dbo.TblContractInstallments.NoteID, dbo.TblContractInstallments.RentValue, dbo.TblContractInstallments.Commissions,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric, dbo.TblContractInstallments.TelandNet,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.payed, dbo.TblContractInstallments.Remains AS Expr1, dbo.TblContractInstallments.RentValuePayed,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.CommissionsPayed, dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.ElectricPayed, dbo.TblContractInstallments.TelandNetPayed, dbo.TblContractInstallments.lastPayedDate,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.lastPayedDateH, dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Doneofall,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.hijri, dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.des, dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContractInstallments.id, dbo.TblContract.Emp_ID,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblAqar.SalesEmp, TblEmployee_1.Emp_Name AS Emp_NameAqr,"
  StrSQL = StrSQL & "                    TblEmployee_1.Fullcode AS FullcodeAqar, TblEmployee_1.Emp_Namee AS Emp_NameAqrE, dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1,"
  StrSQL = StrSQL & "                    dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblCustemers.Cus_Phone,"
  StrSQL = StrSQL & "                    dbo.TblCustemers.Cus_mobile, dbo.TblContract.StrMerg, DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) AS mon,"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstallNo, dbo.TblAqar.BranchId, dbo.TblContract.EndContract, dbo.TblContract.FromdateH,"
  StrSQL = StrSQL & "                    dbo.TblContract.TodateH, dbo.TblContract.RecorddateH, dbo.TblContract.ContType, dbo.TblContract.ContNo AS Expr2, dbo.TblContract.RentType,"
  StrSQL = StrSQL & "                    dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini,"
  StrSQL = StrSQL & "                    dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water AS WaterCont, dbo.TblContract.Electricity, dbo.TblContract.Phone,"
  StrSQL = StrSQL & "                    dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate,"
  StrSQL = StrSQL & "                    dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing, dbo.TblContract.Remarks, dbo.TblContract.FirstInstallDateH,"
  StrSQL = StrSQL & "                    dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblContract.OutContract, dbo.TblContract.OldRent, dbo.TblContract.OldWater,"
  StrSQL = StrSQL & "                    dbo.TblContract.OldElectric, dbo.TblContract.oldCommi, dbo.TblContract.DivWater, dbo.TblContract.DivElectric, dbo.TblContract.OldInsurance,"
  StrSQL = StrSQL & "                    dbo.TblContract.balanceDate, dbo.TblContract.balanceDateH, dbo.TblContract.balanceDes, dbo.TblContract.Renew, dbo.TblContract.ContNoOld,"
  StrSQL = StrSQL & "                    dbo.TblContract.FromdateHO, dbo.TblContract.FromdateO, dbo.TblContract.Employeecontract, dbo.TblContract.Emp_IDContract, dbo.TblContract.OutOffice,"
  StrSQL = StrSQL & "                    dbo.TblContract.LegalIssue, dbo.TblContract.NotValue, dbo.TblContract.UnitElectric, dbo.TblContract.RetValue2, dbo.TblContract.WaterValue2,"
  StrSQL = StrSQL & "                    dbo.TblContract.CommValue2, dbo.TblContract.InstrunceValue2, dbo.TblContract.ElectricityValue2, dbo.TblContract.MiniRentValue, dbo.TblContract.Servce,"
  StrSQL = StrSQL & "                    dbo.TblContract.MethodDeci, dbo.TblContract.FlagContrNew2, dbo.TblFiterWaiver.FilterDateH, dbo.TblFiterWaiver.FilterDate,"
  StrSQL = StrSQL & "                    ISNULL(dbo.Getfitterdata(dbo.TblFiterWaiver.ID, 5," & SQLDate(FromDate.value, True) & ", " & SQLDate(ToDate.value, True) & "), 0) AS totalpayed, ISNULL(dbo.Getfitterdata(dbo.TblFiterWaiver.ID, 4,"
  StrSQL = StrSQL & "                    " & SQLDate(FromDate.value, True) & ", " & SQLDate(ToDate.value, True) & "), 0) AS totalcollected, ISNULL(dbo.TblFiterWaiver.net, 0) AS net, dbo.TblFiterWaiver.ID AS FiterNo"
  StrSQL = StrSQL & "    FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblFiterWaiver RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblContract ON dbo.TblFiterWaiver.ContNo = dbo.TblContract.ContNo LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.TblContract.Emp_ID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id ON dbo.TblBranchesData.branch_id = dbo.TblContract.Branch_NO LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqar ON TblEmployee_1.Emp_ID = dbo.TblAqar.SalesEmp ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo"
  StrSQL = StrSQL & " WHERE    (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL) and  isnull(dbo.TblFiterWaiver.ID,0)<>0 "
  ''((dbo.TblContract.EndContract IS NULL) OR (dbo.TblContract.EndContract=0))  "


    BolBegine = False
    StrWhere = ""
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Branch_NO = " & val(Me.DcbBranch.BoundText)

End If
If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(Me.dcaqartypeid.BoundText)
End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.text <> "" Then

StrWhere = StrWhere & " AND dbo.TblContract.Iqar = " & val(Me.dcbAqarType.BoundText)

End If

If val(Me.DcboEmpaqar.BoundText) <> 0 Or Me.DcboEmpaqar.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.SalesEmp   = " & val(DcboEmpaqar.BoundText)

End If

If val(Me.DcboEmp.BoundText) <> 0 Or Me.DcboEmp.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.Emp_ID    = " & val(DcboEmp.BoundText)

End If

If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitType  = " & val(DcbUnitType.BoundText)

End If
If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.UnitNo  = " & val(DcbUnitNo.BoundText)

End If
If val(TxtAmount.text) <> 0 Then

If opt2(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  > " & val(TxtAmount.text)
ElseIf opt2(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  < " & val(TxtAmount.text)
ElseIf opt2(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  <= " & val(TxtAmount.text)
ElseIf opt2(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  >= " & val(TxtAmount.text)
ElseIf opt2(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.installValue  = " & val(TxtAmount.text)
End If
End If
If paym(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo = 1 "
End If
If paym(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblContractInstallments.InstallNo <> 1"
End If
If paym(2).value = True Then
StrWhere = StrWhere & " AND ( (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) OR"
StrWhere = StrWhere & " (dbo.TblContractInstallments.InstallNo > 1)) "
End If
If paym(3).value = True Then
StrWhere = StrWhere & " AND (dbo.TblContractInstallments.InstallNo = 1) and  (NOT (DATEDIFF(d, dbo.TblContractInstallments.Installdate, GETDATE()) BETWEEN 0 AND " & val(TxtDay.text) & ")) "
End If

   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContractInstallments.Installdate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContractInstallments.Installdate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If

    StrSQL = StrSQL & StrWhere
 If opt(1).value = True Then

 StrSQL = StrSQL & " order by   dbo.TblContract.unitno "
 ElseIf opt(0).value = True Then
 StrSQL = StrSQL & " order by  dbo.TblContractInstallments.Installdate "
 Else
  StrSQL = StrSQL & " order by  dbo.TblContractInstallments.ID "
  End If
    Set rs = New ADODB.Recordset
    Cn.CommandTimeout = 10000
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
    End If

End Sub
Public Sub GetDataExpEleec()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 StrSQL = "SELECT    TblOtheExpensAqar.StatusOper, dbo.TblOtheExpensAqar.ID, dbo.TblOtheExpensAqar.RecordDateH, dbo.TblOtheExpensAqar.RecordDate, dbo.TblOtheExpensAqar.BranchID, "
 StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblOtheExpensAqar.Valuee, dbo.TblOtheExpensAqar.AqarID, dbo.TblAqar.aqarNo,"
 StrSQL = StrSQL & "                     dbo.TblAqar.aqarname, dbo.TblOtheExpensAqar.UnitTypID, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOtheExpensAqar.UnitID, dbo.TblAqarDetai.unitno,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblOtheExpensAqar.EmpID,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Expr1, dbo.TblEmployee.Emp_Namee, dbo.TblOtheExpensAqar.Mobile, dbo.TblOtheExpensAqar.BillNo,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.AccountNo, dbo.TblOtheExpensAqar.AccountBank, dbo.TblOtheExpensAqar.TypID, dbo.TblOtheExpensAqar.Remarks,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.PayedDateH, dbo.TblOtheExpensAqar.PayedDate, dbo.TblOtheExpensAqar.FromDateH, dbo.TblOtheExpensAqar.FromDate,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.ToDateH, dbo.TblOtheExpensAqar.ToDate, dbo.TblOtheExpensAqar.StatusOper, dbo.TblOtheExpensAqar.RemainRent,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Electricity, dbo.TblOtheExpensAqar.Maintenance, dbo.TblOtheExpensAqar.MaintCondition, dbo.TblOtheExpensAqar.MaintDoors,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.MaintKitchen, dbo.TblOtheExpensAqar.MaintClean, dbo.TblOtheExpensAqar.MaintOther, dbo.TblOtheExpensAqar.Insurance,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.DelayDay, dbo.TblOtheExpensAqar.Noliquidation, dbo.TblOtheExpensAqar.Paints, dbo.TblOtheExpensAqar.Windows,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Total, dbo.TblOtheExpensAqar.Net - IsNull(TblOtheExpensAqar.Discount2,0) Net, dbo.TblOtheExpensAqar.Name AS Expr2, dbo.TblOtheExpensAqar.TotalAfterIns,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Discount,TblOtheExpensAqar.Discount2,  dbo.TblOtheExpensAqar.NoteSerial1, dbo.TblOtheExpensAqar.Prefix, dbo.TblOtheExpensAqar.NoteSerial,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.FlgPayed, dbo.TblOtheExpensAqar.BankID, dbo.BanksData.BankName, dbo.BanksData.BankNamee,"
 
 StrSQL = StrSQL & "                     dbo.GetValuePayedElectric(dbo.TblOtheExpensAqar.ID) AS PayedValue,"
 StrSQL = StrSQL & "                      TblOtheExpensAqar.IsLegalAffairs,TblOtheExpensAqar.LegalAffairs,TblOtheExpensAqar.LegalAffairsDate"
 StrSQL = StrSQL & "   FROM         dbo.TblOtheExpensAqar LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.BanksData ON dbo.TblOtheExpensAqar.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblOtheExpensAqar.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.TblOtheExpensAqar.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqarDetai ON dbo.TblOtheExpensAqar.UnitID = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit ON dbo.TblOtheExpensAqar.UnitTypID = dbo.TblAkarUnit.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqar ON dbo.TblOtheExpensAqar.AqarID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblOtheExpensAqar.BranchID = dbo.TblBranchesData.branch_id"
 StrSQL = StrSQL & " where 1=1"
    BolBegine = False
    StrWhere = ""
    If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(Me.dcaqartypeid.BoundText)
End If

If DcbStatusOper.ListIndex <> -1 Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.StatusOper=" & DcbStatusOper.ListIndex & ""
End If
  If TypedID(0).value = True Then
  StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.TypID=0"
  End If
  If TypedID(1).value = True Then
   StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.TypID=1"
  End If
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.BranchID = " & val(Me.DcbBranch.BoundText)
End If
If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.AqarID = " & val(Me.dcbAqarType.BoundText)
End If
If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.UnitTypID  = " & val(DcbUnitType.BoundText)

End If



If chkIsLegalAffairs.value = vbChecked Then
        StrWhere = StrWhere & " and  dbo.TblOtheExpensAqar.IsLegalAffairs = 1 "
End If





If chkIsLegalAffairs.value = vbChecked Then
        StrWhere = StrWhere & " and  IsNull(dbo.TblOtheExpensAqar.IsLegalAffairs,0) = 1 "
End If


If chkIsLegalAffairsNo.value = vbChecked Then
        StrWhere = StrWhere & " and  IsNull(dbo.TblOtheExpensAqar.IsLegalAffairs,0) = 0 "
End If



If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.UnitID  = " & val(DcbUnitNo.BoundText)
End If
If val(Me.DcboEmpaqar.BoundText) <> 0 Or Me.DcboEmpaqar.text <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.EmpID  = " & val(DcboEmpaqar.BoundText)
End If
If val(TxtAmount.text) <> 0 Then

If opt2(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.Net  > " & val(TxtAmount.text)
ElseIf opt2(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.Net  < " & val(TxtAmount.text)
ElseIf opt2(3).value = True Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.Net  <= " & val(TxtAmount.text)
ElseIf opt2(4).value = True Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.Net  >= " & val(TxtAmount.text)
ElseIf opt2(2).value = True Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.Net  = " & val(TxtAmount.text)
End If
End If
   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.RecordDate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If
    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblOtheExpensAqar.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
    End If

    StrSQL = StrSQL & StrWhere
 If opt(1).value = True Then

 StrSQL = StrSQL & " order by   dbo.TblOtheExpensAqar.UnitID "
 ElseIf opt(0).value = True Then
 StrSQL = StrSQL & " order by  dbo.TblOtheExpensAqar.RecordDate "
 Else
  StrSQL = StrSQL & " order by  dbo.TblOtheExpensAqar.ID "
  End If
    Set rs = New ADODB.Recordset
    Cn.CommandTimeout = 10000
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL, 1
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
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
   If Ind = 1 Then
          If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtherExpElect.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOtherExpElect.rpt"
            
       End If

Else
If indexx = 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports.rpt"
            
       End If
 ElseIf indexx = 1 Then
     If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports1.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports1.rpt"
            
       End If
  ElseIf indexx = 3 Then
     If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendWithWaverReports.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendWithWaverReportsE.rpt"
            
       End If
  Else
  
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendMonthReports.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendMonthReports.rpt"
            
       End If
  End If
  End If
            

If Ind = 30 Then
Dim X As Integer

X = MsgBox("ÿ»«⁄Â  ⁄—»Ì", vbInformation + vbYesNo)

       If X = vbYes Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports0003A.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarRentsOwendReports0003.rpt"
            
       End If
End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
    Dim MSGType As Integer
    If Ind = 0 And indexx <> 3 Then
     MSGType = MsgBox("Â·  —€» ›Ì   «ŸÂ«— «·«—’œÂ «·’›—ÌÂ", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)

        If MSGType = vbYes Then
         xReport.ParameterFields(18).AddCurrentValue 0
         Else
          xReport.ParameterFields(18).AddCurrentValue 1
        End If
End If
If Ind = 30 Then
xReport.ParameterFields(18).AddCurrentValue 1
End If

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
       If Accredit(0).value = True Then
        xReport.ParameterFields(2).AddCurrentValue "«·⁄ﬁÊœ «·„ÊÀ›… ›ﬁÿ"
       ElseIf Accredit(1).value = True Then
       xReport.ParameterFields(2).AddCurrentValue "«·⁄ﬁÊœ €Ì— «·„ÊÀ…"
       Else
       xReport.ParameterFields(2).AddCurrentValue "ﬂ· «·⁄ﬁÊœ"
       End If
       
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   If FromDate.value <> "" And ToDate.value <> "" Then
    xReport.ParameterFields(14).AddCurrentValue FromDate.value
       xReport.ParameterFields(15).AddCurrentValue FromDateH.value
       xReport.ParameterFields(16).AddCurrentValue ToDate.value
       xReport.ParameterFields(17).AddCurrentValue ToDateH.value
       End If

  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function




Private Sub FromDate_Change()
If Not IsNull(FromDate.value) Then
   FromDateH.value = ToHijriDate(FromDate.value)
   End If
End Sub



Private Sub Fromdateh_LostFocus()

 VBA.Calendar = vbCalGreg
            FromDate.value = ToGregorianDate(FromDateH.value)

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub OptRep_Click(index As Integer)
HideElemint
TypedID(2).value = True
End Sub

Private Sub Text1_Change()
DcboEmpaqar.BoundText = GeTEmpIDByEmpCode(Me.Text1.text, True)
End Sub

Private Sub ToDate_Change()
If Not IsNull(ToDate.value) Then
   ToDateH.value = ToHijriDate(ToDate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()

 VBA.Calendar = vbCalGreg
            ToDate.value = ToGregorianDate(ToDateH.value)

End Sub
Private Sub TxtEmployeeID_Change()
DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.text, True)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
