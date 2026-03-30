VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmCustomerAqarReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmCustomerAqarReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   5205
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10395
      Begin XtremeSuiteControls.CheckBox Chck 
         Height          =   375
         Left            =   4440
         TabIndex        =   49
         Top             =   4080
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ﬂ‘› Õ”«» «·„” √Ã—"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox TxtRent 
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
         Left            =   1800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   2040
         Width           =   465
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   43
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   2040
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "«’€— „‰"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox TxtContNo 
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
         Left            =   3960
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtIqama 
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
         Left            =   240
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TxtCusMobil 
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
         Left            =   3960
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   600
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2400
         Width           =   6555
         Begin XtremeSuiteControls.RadioButton Rd1 
            Height          =   375
            Left            =   3600
            TabIndex        =   30
            Top             =   120
            Width           =   2175
            _Version        =   786432
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Ì‰ „Ì «·Ï «·ﬁ«∆„Â «·”Êœ«¡"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd2 
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   2295
            _Version        =   786432
            _ExtentX        =   4048
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "€Ì— „‰ „Ì «·Ï «·ﬁ«∆„Â «·”Êœ«¡"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —… "
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3000
         Width           =   6555
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
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
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   22
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   24
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
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   6960
         TabIndex        =   17
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmCustomerAqarReport.frx":038A
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
            TabIndex        =   18
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   16
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   15
         Top             =   5280
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   240
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
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcCustomer 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbNationality 
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCusType 
         Height          =   315
         Left            =   3960
         TabIndex        =   47
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·„” «Ã—"
         Height          =   195
         Index           =   8
         Left            =   5235
         TabIndex        =   48
         Top             =   2040
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ··«ÌÃ«—"
         Height          =   195
         Index           =   7
         Left            =   2970
         TabIndex        =   45
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ··Ã‰”ÌÂ "
         Height          =   195
         Index           =   5
         Left            =   2925
         TabIndex        =   40
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·⁄ﬁœ"
         Height          =   195
         Index           =   6
         Left            =   5535
         TabIndex        =   39
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·ÂÊÌÂ"
         Height          =   195
         Index           =   4
         Left            =   2685
         TabIndex        =   37
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·ÃÊ«·"
         Height          =   195
         Index           =   3
         Left            =   5430
         TabIndex        =   36
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   4680
         Width           =   6615
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
         TabIndex        =   27
         Top             =   4680
         Width           =   6615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„” «Ã— „Õœœ"
         Height          =   195
         Index           =   2
         Left            =   5265
         TabIndex        =   9
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄ﬁ«— „Õœœ"
         Height          =   195
         Index           =   1
         Left            =   5475
         TabIndex        =   8
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „Õœœ"
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   6000
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
      Top             =   6000
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
      Left            =   600
      Picture         =   "FrmCustomerAqarReport.frx":10A48
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·„” √Ã—Ì‰"
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
      Left            =   -15
      TabIndex        =   6
      Top             =   0
      Width           =   10470
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
Attribute VB_Name = "FrmCustomerAqarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Sub Acount()
If val(dcCustomer.BoundText) = 0 Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„” «Ã—"
dcCustomer.SetFocus
Exit Sub
End If
If IsNull(ToDate.value) Or IsNull(Fromdate.value) Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
Exit Sub
End If
        Dim Account_code As String
        Account_code = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
updateopeningbalanceNewFromsql Fromdate.value, ToDate.value, False, 0, 0, Account_code, 3
        ShowReport Account_code, dcCustomer.Text, Fromdate.value, ToDate.value
End Sub
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       
If Chck.value = vbChecked Then
Acount
Else
 GetData
End If
            
        Case 1
   
            clear_all Me
         Fromdate.value = ""
    ToDate.value = ""
      Rd1.value = False
    Rd2.value = False
    Opt(0).value = False
    Opt(1).value = False
    Opt(2).value = False
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
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
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , dcbAqarType.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
End Sub





Private Sub dcsupplier_Click(Area As Integer)

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
    Dcombos.GETNationality Me.DcbNationality
    Dcombos.GetCustomerType DcbCusType
    Dcombos.GetCustomersSuppliers 56, dcCustomer
    Dcombos.GetBranches DcbBranch
    
 ' Dcombos.GetRentStatus dbcAqarStatus
    
    Fromdate.value = ""
    ToDate.value = ""
    Rd1.value = False
    Rd2.value = False
    Opt(0).value = False
    Opt(1).value = False
    Opt(2).value = False
    
    
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


StrSQL = "SELECT     dbo.TblContract.ContType, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblContract.TotalContract, dbo.TblContract.balanceDes, "
StrSQL = StrSQL & "                      dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Remark, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustomerType.name,"
StrSQL = StrSQL & "                      dbo.TblCustomerType.namee, dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.locked, dbo.TblCustemers.CustGID, dbo.TblCustemers.Company,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CountryID2, dbo.Nationality.name AS nameNationality, dbo.Nationality.namee AS nameNationalityE, dbo.TblContract.Water,"
StrSQL = StrSQL & "                      dbo.TblContract.Electricity, dbo.TblContract.Phone, dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate,"
StrSQL = StrSQL & "                      dbo.TblContract.UnitType, dbo.TblContract.UnitNo, dbo.TblContract.RentType, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue,"
StrSQL = StrSQL & "                      dbo.TblContract.PayAmini, dbo.TblContract.MeterCount, dbo.TblContract.CommiValue, dbo.TblContract.ownerid, dbo.TblContract.InsuranceValue,"
StrSQL = StrSQL & "                      dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate, dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing,"
StrSQL = StrSQL & "                      dbo.TblContract.Remarks, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.TodateH, dbo.TblContract.FirstInstallDateH,"
StrSQL = StrSQL & "                      dbo.TblContract.NoteID, dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblContract.Emp_ID, dbo.TblContract.OutContract,"
StrSQL = StrSQL & "                      dbo.TblContract.OldRent, dbo.TblContract.OldWater, dbo.TblContract.OldElectric, dbo.TblContract.oldCommi, dbo.TblContract.DivWater, dbo.TblContract.DivElectric,"
StrSQL = StrSQL & "                      dbo.TblContract.OldInsurance, dbo.TblContract.balanceDate, dbo.TblContract.balanceDateH, dbo.TblContract.Renew, dbo.TblContract.ContNoOld,"
StrSQL = StrSQL & "                      dbo.TblContract.FromdateHO, dbo.TblContract.FromdateO, dbo.TblContract.EndContract, dbo.TblContract.Employeecontract, dbo.TblContract.Emp_IDContract,"
StrSQL = StrSQL & "                      dbo.TblContract.OutOffice, dbo.TblContract.LegalIssue, dbo.TblContract.NotID, dbo.TblContract.NoteSrial1, dbo.TblContract.NotValue, dbo.TblCustemers.Type,"
StrSQL = StrSQL & "                      dbo.TblContract.ContNo"
StrSQL = StrSQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Nationality.id = dbo.TblCustemers.CountryID2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustomerType ON dbo.TblCustemers.CustomerTypeID = dbo.TblCustomerType.id RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblContract ON dbo.TblCustemers.CusID = dbo.TblContract.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    If Rd1.value = True Then
    StrWhere = StrWhere & " AND dbo.TblCustemers.locked =1"
    End If
     If Rd2.value = True Then
    StrWhere = StrWhere & " AND dbo.TblCustemers.locked =0"
    End If
   If Me.TxtRent.Text <> "" Then
 
      If Opt(0).value = True Then
                   StrWhere = StrWhere & " AND dbo.TblContract.TotalContract < " & val(TxtRent.Text) & ""
      End If
If Opt(1).value = True Then
                   StrWhere = StrWhere & " AND dbo.TblContract.TotalContract = " & val(TxtRent.Text) & ""
      End If
      If Opt(2).value = True Then
                   StrWhere = StrWhere & " AND dbo.TblContract.TotalContract > " & val(TxtRent.Text) & ""
      End If
      
   End If
    If Me.TxtContNo.Text <> "" Then
              StrWhere = StrWhere & " AND dbo.TblContract.NoteSerial1 ='" & TxtContNo.Text & "'"
              End If
              
If TxtCusMobil.Text <> "" Then
              StrWhere = StrWhere & " AND dbo.TblCustemers.Cus_Phone ='" & TxtCusMobil.Text & "'"
              End If

 If TxtIqama.Text <> "" Then
              StrWhere = StrWhere & " AND dbo.TblCustemers.CustGID ='" & TxtIqama.Text & "'"
    End If
    
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND TblContract.Branch_NO = " & val(Me.DcbBranch.BoundText)

End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.Text <> "" Then

StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)

End If


If val(Me.dcCustomer.BoundText) <> 0 Or Me.dcCustomer.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.CusID  = " & val(dcCustomer.BoundText)

End If
If val(Me.DcbNationality.BoundText) <> 0 Or Me.DcbNationality.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblCustemers.CountryID2  = " & val(DcbNationality.BoundText)

End If
If val(Me.DcbCusType.BoundText) <> 0 Or Me.DcbCusType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblCustemers.CustomerTypeID  = " & val(DcbCusType.BoundText)

End If

     If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblContract.StrDate >=" & SQLDate(Me.Fromdate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblContract.StrDate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If




    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL & " order by  dbo.TblContract.ContNo "
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
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
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCustomerAqarReportsh.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCustomerAqarReportsh.rpt"
            
       End If
  
            
    ' If Me.RdDept.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
     '       Else
      '      If Me.RdSuper.value = True Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
        '    Else
         '   If Me.RdFitter.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
          ' Else
             
            '        If Me.RdAll2.value = True Then
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
          '  Else
           '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            
     '      End If
      '      End If
       '     End If
        '     End If
         '   End If
          '  End If
        '    End If
           ' End If
          '  End If
       '      End If
           '
      '  End If



    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        'xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim Total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
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
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
