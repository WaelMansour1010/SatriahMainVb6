VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmUnitInfoReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   Icon            =   "FrmUnitInfoReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10380
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
      Left            =   2760
      TabIndex        =   18
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   9
      Top             =   7080
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
      Caption         =   "»Ì«‰«  "
      Height          =   5565
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   10395
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   " — Ì» »Õ”»"
         Height          =   495
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   480
         Width           =   6135
         Begin XtremeSuiteControls.RadioButton RdSort 
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   48
            Top             =   120
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " — Ì» »Õ”» «· «—ÌŒ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdSort 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   120
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " — Ì» »Õ”» «·⁄ﬁ«—"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.RadioButton RdUnitStatus 
         Height          =   375
         Left            =   3240
         TabIndex        =   45
         Top             =   120
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " ﬁ—Ì— Õ«·… «·ÊÕœ…"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
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
         TabIndex        =   36
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
         Height          =   555
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3240
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
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   3720
         Width           =   6555
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   20
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
            TabIndex        =   21
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   23
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
            Caption         =   "„‰"
            Height          =   315
            Index           =   3
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   6960
         TabIndex        =   16
         Top             =   0
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmUnitInfoReports.frx":038A
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
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Top             =   5520
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Top             =   5520
         Width           =   855
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
         Top             =   1440
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
         TabIndex        =   28
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   2520
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
         TabIndex        =   29
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   2160
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
         TabIndex        =   34
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
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
         TabIndex        =   37
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   2280
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbRentStatus 
         Height          =   315
         Left            =   240
         TabIndex        =   39
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   2880
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranches 
         Height          =   315
         Left            =   240
         TabIndex        =   44
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdUnitTrans 
         Height          =   375
         Left            =   600
         TabIndex        =   46
         Top             =   120
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " ﬁ—Ì— Õ—ﬂ… «·ÊÕœ…"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcaqartypeid 
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   1800
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
         Caption         =   " ÿ»ﬁ« ·‰Ê⁄ «·⁄ﬁ«—"
         Height          =   195
         Index           =   7
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   6
         Left            =   5565
         TabIndex        =   43
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   735
         Left            =   120
         Top             =   4800
         Width           =   6855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ì—ÃÏ «Œ Ì«— «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· Ê«·„œ…"
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
         TabIndex        =   42
         Top             =   4800
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·Ê’› ·ÊÕœ… „ÕœœÂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3360
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·Õ«·Â „ÕœœÂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„Õ’· ⁄ﬁ«— „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   16800
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„”Êﬁ ⁄ﬁœ „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   17040
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
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
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·‰Ê⁄ „Õœœ"
         Height          =   195
         Index           =   15
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   6000
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
         TabIndex        =   26
         Top             =   6240
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5565
         TabIndex        =   8
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
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
      Top             =   6240
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
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   6240
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   2880
      Picture         =   "FrmUnitInfoReports.frx":10A48
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì—  Õ—ﬂ«  «·ÊÕœÂ"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10350
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
Attribute VB_Name = "FrmUnitInfoReports"
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
Private Sub btnClear_Click()
    Cmd_Click (1)
End Sub
Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            If RdUnitStatus.value = True Or RdUnitTrans.value = True Then
                GetData
            Else
                MsgBox "Ì—ÃÏ ≈Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
                Exit Sub
            End If
        Case 1
            clear_all Me
            Fromdate.value = ""
            ToDate.value = ""
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
    Dim str As String
    Dim EmpCode  As String
    
    If val(dcbAqarType.BoundText) = 0 Then: Exit Sub

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
        Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
    End If
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
    
    Dcombos.GetIqar dcbAqarType
    Dcombos.GetRentSatus DcbRentStatus
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    Dcombos.GetBranches DcbBranches
    Dcombos.getAkarType Me.dcaqartypeid
    Fromdate.value = ""
    ToDate.value = ""
    Cmd_Click (1)
    
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

    If RdUnitTrans.value = True Then
        StrSQL = " SELECT     dbo.TblUnitNoInformation.ID, dbo.TblUnitNoInformation.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname, "
        StrSQL = StrSQL & "              dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblUnitNoInformation.UnitStatus, dbo.TblRentStatus.name AS nameStatus,"
        StrSQL = StrSQL & "              dbo.TblRentStatus.namee AS nameStatusE, dbo.TblUnitNoInformation.RecDate, dbo.TblUnitNoInformation.RecDateH, dbo.TblUnitNoInformation.BranchId,"
        StrSQL = StrSQL & "              dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblUnitNoInformation.CusID, dbo.TblCustemers.CusName,"
        StrSQL = StrSQL & "              dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Cus_mobile, dbo.TblUnitNoInformation.Des, dbo.TblUnitNoInformation.NoteID,"
        StrSQL = StrSQL & "              dbo.Notes.renterName, dbo.Notes.CashingType, dbo.TblAqarDetai.ContID, dbo.TblContract.Periods, dbo.TblContract.PeriodsID, dbo.TblAqar.aqartypeid,"
        StrSQL = StrSQL & "              dbo.tblAkarType.name AS AqarType, dbo.tblAkarType.namee AS AqarTypeE"
        StrSQL = StrSQL & "           FROM         dbo.tblAkarType RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAqar ON dbo.tblAkarType.id = dbo.TblAqar.aqartypeid RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAkarUnit RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblContract RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAqarDetai ON dbo.TblContract.ContNo = dbo.TblAqarDetai.ContID ON dbo.TblAkarUnit.id = dbo.TblAqarDetai.unittype RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.Notes RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblUnitNoInformation ON dbo.Notes.NoteID = dbo.TblUnitNoInformation.NoteID LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblCustemers ON dbo.TblUnitNoInformation.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblBranchesData ON dbo.TblUnitNoInformation.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblRentStatus ON dbo.TblUnitNoInformation.UnitStatus = dbo.TblRentStatus.id ON dbo.TblAqarDetai.Id = dbo.TblUnitNoInformation.UnitNo ON"
        StrSQL = StrSQL & "               dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid"
        StrSQL = StrSQL & " WHERE (1=1) "
        
        BolBegine = False
        StrWhere = ""
        
        If Me.TxtDes.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblUnitNoInformation.Des like ' % " & Me.TxtDes.Text & " % '"
        End If
        If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(dcaqartypeid.BoundText)
        End If
               If val(Me.DcbRentStatus.BoundText) <> 0 Or Me.DcbRentStatus.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblUnitNoInformation.UnitStatus = " & val(DcbRentStatus.BoundText)
        End If
        
        If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblUnitNoInformation.UnitNo= " & val(DcbUnitNo.BoundText)
        End If
        If Not IsNull(Me.Fromdate.value) Then
            StrWhere = StrWhere & " AND dbo.TblUnitNoInformation.RecDate >=" & SQLDate(Me.Fromdate.value, True) & ""
        End If
        If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblUnitNoInformation.RecDate <=" & SQLDate(Me.ToDate.value, True) & ""
        End If
    End If

    If RdUnitStatus.value = True Then
        StrSQL = " SELECT     dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, "
        StrSQL = StrSQL & "              dbo.TblRentStatus.name AS nameStatus, dbo.TblRentStatus.namee AS nameStatusE, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
        StrSQL = StrSQL & "              dbo.TblAqar.BranchId, dbo.TblAqarDetai.rentType, dbo.TblRentStatus.id, dbo.TblAqarDetai.Id AS IdDet, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.FilterDateH,"
        StrSQL = StrSQL & "              dbo.TblAqarDetai.FilterDate, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.ACCount,"
        StrSQL = StrSQL & "              dbo.TblAqarDetai.ACCountspleat, dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.length, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.haveFurniture,"
        StrSQL = StrSQL & "              dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ContID, dbo.TblContract.ContDate, dbo.TblContract.Periods, dbo.TblContract.PeriodsID,"
        StrSQL = StrSQL & "              dbo.TblAqar.aqartypeid, dbo.tblAkarType.name AS AqarType, dbo.tblAkarType.namee AS AqarTypeE"
        StrSQL = StrSQL & "    FROM         dbo.TblRentStatus RIGHT OUTER JOIN"
        StrSQL = StrSQL & "               dbo.TblContract RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAqarDetai ON dbo.TblContract.ContNo = dbo.TblAqarDetai.ContID ON dbo.TblRentStatus.id = dbo.TblAqarDetai.Status RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAqar LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id"
        StrSQL = StrSQL & " Where (1 = 1)"
        
        BolBegine = False
        StrWhere = ""
            
        If Not IsNull(Me.Fromdate.value) Then
            StrWhere = StrWhere & " AND dbo.TblAqarDetai.FilterDate >=" & SQLDate(Me.Fromdate.value, True) & ""
        End If
        If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblAqarDetai.FilterDate <=" & SQLDate(Me.ToDate.value, True) & ""
        End If
        If val(Me.DcbRentStatus.BoundText) <> 0 Or Me.DcbRentStatus.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblAqarDetai.Status = " & val(DcbRentStatus.BoundText)
        End If
        If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblAqarDetai.Id= " & val(DcbUnitNo.BoundText)
        End If
        If val(Me.dcaqartypeid.BoundText) <> 0 Or Me.dcaqartypeid.Text <> "" Then
            StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid = " & val(dcaqartypeid.BoundText)
        End If
    End If
    
    If val(Me.DcbBranches.BoundText) <> 0 Or Me.DcbBranches.Text <> "" Then
        StrWhere = StrWhere & " AND dbo.TblBranchesData.branch_id = " & val(Me.DcbBranches.BoundText)
    End If
    If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.Text <> "" Then
        StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)
    End If
    If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.Text <> "" Then
        StrWhere = StrWhere & " AND dbo.TblAqarDetai.unittype = " & val(DcbUnitType.BoundText)
    End If

    StrSQL = StrSQL & StrWhere
 
    If RdUnitTrans.value = True Then
        StrSQL = StrSQL & " order by  dbo.TblUnitNoInformation.ID "
    End If
    If RdUnitStatus.value = True Then
        If RdSort(0).value = True Then
            StrSQL = StrSQL & " order by  dbo.TblAqarDetai.FilterDate "
        Else
            StrSQL = StrSQL & " order by  dbo.TblAqar.Aqarid "
        End If
    End If
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        rs.MoveFirst
        print_report StrSQL
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

    If RdUnitTrans.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaUnitInfoStatus.rpt"
        Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaUnitInfoStatus.rpt"
        End If
    End If
    If RdUnitStatus.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaUnitInformation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqaUnitInformation.rpt"
       End If
    End If

    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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
    If Fromdate.value <> "" Then
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
Private Sub RdUnitStatus_Click()
    Frame8.Enabled = True
    TxtDes.Enabled = False
End Sub
Private Sub RdUnitTrans_Click()
    Frame8.Enabled = True
    TxtDes.Enabled = True
End Sub
Private Sub Text1_Change()
    DcboEmpaqar.BoundText = GeTEmpIDByEmpCode(Me.Text1.Text, True)
End Sub
Private Sub ToDate_Change()
    If ToDate.value <> "" Then
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

    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
