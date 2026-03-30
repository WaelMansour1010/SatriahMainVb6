VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmCarAutoMSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "««·»ÕÀ ⁄‰ »ÿ«ﬁ… «–‰ «’·«Õ"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16395
   Icon            =   "FromCarAutoMSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   16395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1035
      Index           =   0
      Left            =   3480
      TabIndex        =   73
      Top             =   4590
      Width           =   7125
      Begin VB.TextBox TxtPlateNo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4440
         TabIndex        =   75
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TxtSahseh 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4410
         TabIndex        =   74
         Top             =   600
         Width           =   1725
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   435
         Left            =   30
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   120
         Width           =   4365
         _cx             =   7699
         _cy             =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox txtLetter1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3630
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   0
            Width           =   540
         End
         Begin VB.TextBox txtLetter2 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3210
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   0
            Width           =   450
         End
         Begin VB.TextBox txtLetter3 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2700
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   0
            Width           =   600
         End
         Begin VB.TextBox txtNum1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1500
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   675
         End
         Begin VB.TextBox txtNum2 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   900
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox txtNum3 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   510
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   0
            Width           =   555
         End
         Begin VB.TextBox txtLetter4 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2175
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   0
            Width           =   675
         End
         Begin VB.TextBox txtNum4 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   0
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„  «·‘«’ÌÂ"
         Height          =   195
         Index           =   8
         Left            =   6150
         TabIndex        =   86
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   195
         Index           =   0
         Left            =   6150
         TabIndex        =   85
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— ‘€·"
      Height          =   645
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2640
      Width           =   4515
      Begin VB.TextBox TxtWorkOrder 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         TabIndex        =   70
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtWorkOrderTo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   27
         Left            =   3255
         TabIndex        =   72
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   26
         Left            =   1380
         TabIndex        =   71
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ≈–‰ «·«’·«Õ"
      Height          =   645
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   2640
      Width           =   3915
      Begin VB.TextBox TxtAuthoOrderTo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtAuthoOrder 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         TabIndex        =   63
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   25
         Left            =   1380
         TabIndex        =   66
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   24
         Left            =   3255
         TabIndex        =   65
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ⁄—÷ ”⁄—"
      Height          =   645
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   2640
      Width           =   3915
      Begin VB.TextBox TxtShowPriceOrder 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         TabIndex        =   59
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtShowPriceOrderTo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   23
         Left            =   3255
         TabIndex        =   61
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   21
         Left            =   1380
         TabIndex        =   60
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   5
      Left            =   10680
      TabIndex        =   53
      Top             =   6450
      Width           =   5715
      Begin MSDataListLib.DataCombo DatacomUser 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰Ì «·«” ﬁ»«·"
         Height          =   195
         Index           =   22
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1125
      Index           =   4
      Left            =   0
      TabIndex        =   51
      Top             =   4320
      Width           =   3405
      Begin VB.TextBox TxtComplaint 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtNoteIntial1 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘ﬂÊÏ «·⁄„Ì·"
         Height          =   195
         Index           =   19
         Left            =   2040
         TabIndex        =   55
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «·„·«ÕŸ…"
         Height          =   195
         Index           =   20
         Left            =   2040
         TabIndex        =   52
         Top             =   300
         Width           =   1125
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   765
      Index           =   3
      Left            =   3960
      TabIndex        =   44
      Top             =   6330
      Width           =   6675
      Begin VB.TextBox TxtDoor 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   2235
      End
      Begin VB.TextBox TxtReg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·»«»"
         Height          =   195
         Index           =   16
         Left            =   2460
         TabIndex        =   46
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·”Ã·"
         Height          =   195
         Index           =   15
         Left            =   5760
         TabIndex        =   45
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·„⁄œÂ/«·”Ì«—…"
      Height          =   1215
      Index           =   2
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3360
      Width           =   7125
      Begin VB.ComboBox DcbyearFactor 
         Height          =   315
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   720
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DcbCarType 
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCbCarModel 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbColor 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—«“ «·„⁄œÂ/«·”Ì«—…"
         Height          =   195
         Index           =   12
         Left            =   2040
         TabIndex        =   43
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Ê‰ «·„⁄œÂ/«·”Ì«—…"
         Height          =   195
         Index           =   14
         Left            =   2040
         TabIndex        =   42
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
         Height          =   195
         Index           =   13
         Left            =   5520
         TabIndex        =   41
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—«“"
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   40
         Top             =   240
         Width           =   15
      End
      Begin VB.Label lbltype 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
         Height          =   195
         Left            =   5670
         TabIndex        =   39
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·⁄„Ì·"
      Height          =   2085
      Index           =   1
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3360
      Width           =   5715
      Begin VB.TextBox txtcode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3840
         TabIndex        =   56
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1875
      End
      Begin VB.TextBox TxtAdress 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3120
         TabIndex        =   4
         Top             =   1680
         Width           =   1635
      End
      Begin VB.TextBox TxtFax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1875
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   1635
      End
      Begin VB.TextBox TxtClientName 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   3795
      End
      Begin VB.TextBox txtmobile 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3120
         TabIndex        =   2
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox TxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’‰œÊﬁ »—Ìœ"
         Height          =   195
         Index           =   18
         Left            =   2040
         TabIndex        =   50
         Top             =   1740
         Width           =   885
      End
      Begin VB.Label lbladdress 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   255
         Left            =   4680
         TabIndex        =   49
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›«ﬂ”"
         Height          =   195
         Index           =   17
         Left            =   2040
         TabIndex        =   48
         Top             =   1260
         Width           =   885
      End
      Begin VB.Label lblemail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ì„Ì·"
         Height          =   255
         Left            =   4680
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label LblClientName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   195
         Left            =   5220
         TabIndex        =   37
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«· «·⁄„Ì·"
         Height          =   195
         Index           =   11
         Left            =   4680
         TabIndex        =   36
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â« › «·⁄„Ì·"
         Height          =   195
         Index           =   9
         Left            =   2040
         TabIndex        =   35
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   16920
      TabIndex        =   33
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   120
      Width           =   1035
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   16800
      TabIndex        =   32
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3360
      Width           =   3405
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   155189251
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   155189251
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   29
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   2175
         TabIndex        =   28
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2640
      Width           =   3915
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1380
         TabIndex        =   26
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   3255
         TabIndex        =   25
         Top             =   240
         Width           =   540
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   6780
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
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   20
      Top             =   6780
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   1140
      TabIndex        =   21
      Top             =   6780
      Width           =   735
      _ExtentX        =   1296
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   16395
      _cx             =   28919
      _cy             =   4630
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
      Cols            =   28
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FromCarAutoMSearch.frx":038A
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
   Begin XtremeSuiteControls.CheckBox ChAccept 
      Height          =   495
      Left            =   12660
      TabIndex        =   87
      Top             =   5790
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   " „ „Ê«›ﬁ…  «·⁄„Ì· "
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   495
      Left            =   11460
      TabIndex        =   88
      Top             =   5790
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   " Õ  «·«‰ Ÿ«—"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox CheckBox2 
      Height          =   495
      Left            =   9900
      TabIndex        =   89
      Top             =   5790
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "⁄œ„ „Ê«›ﬁ…«·⁄„Ì·"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      TabIndex        =   31
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   30
      Top             =   6180
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   22
      Top             =   6240
      Width           =   1785
   End
End
Attribute VB_Name = "FrmCarAutoMSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch



Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub Cal_Board()
    TxtPlateNO.Text = Trim(txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text)
End Sub
Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub
Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board

End Sub
Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 
 GetData
      
        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub DcbCarType_Change()
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DcbCarType.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DcbCarType.BoundText)
   End If
End Sub


Private Sub Fg_Click()

If FrmCarAuthontication.bo = True Then
  FrmCarAuthontication.retrive1 (val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id"))))
  Else
  FrmCarAuthontication.Retrive3 (val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id"))))
  

End If


End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
   '  Dcombos.GetClientName Me.DCEmp_Name
      If SystemOptions.UserInterface = EnglishInterface Then
     
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"

             Else
  
 DcbOrderStatus.AddItem "ÃœÌœ"
DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


    End If
    Set DCboSearch = New clsDCboSearch
    
    Dcombos.GetUsers Me.DatacomUser
    Set DCboSearch.Client = Me.DatacomUser
    
   Dim year As Integer

  Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetTblColor Me.DcbColor
    Dim i As Integer
      For i = 1995 To 2100
      Me.DcbyearFactor.AddItem (i)
      Next i
      
   Dcombos.GetTblCarModels Me.DcbCarModel
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

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
StrSQL = "SELECT     dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReform.RecordDate, "
If SystemOptions.UserInterface = EnglishInterface Then

    StrSQL = StrSQL & "                    isNull(TblCustemers.CusNamee,    dbo.TblCardAuthorizationReform.ClientName) as ClientName, "
    StrSQL = StrSQL & "                      dbo.TblCarModels.ModelE as Model , dbo.TblCarModels.ModelE, "
    StrSQL = StrSQL & "                      dbo.TBLCarTypes.namee as name, dbo.TBLCarTypes.namee,"
    StrSQL = StrSQL & "                      dbo.TblColor.namee AS ColnameE, dbo.TblColor.namee AS Colname,"
Else
    StrSQL = StrSQL & "                   isNull(TblCustemers.CusName,    dbo.TblCardAuthorizationReform.ClientName) as ClientName  , "
    StrSQL = StrSQL & "                      dbo.TblCarModels.Model , dbo.TblCarModels.ModelE, "
    StrSQL = StrSQL & "                      dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
    StrSQL = StrSQL & "                      dbo.TblColor.name Colname, dbo.TblColor.namee AS ColnameE,"
End If
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.CarTypeID, "
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.CarModelID, "


StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.PlateNo, TblCardAuthorizationReform.Accept,TblCardAuthorizationReform.wait,TblCardAuthorizationReform.notAcepted,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.ColorID, dbo.TblCardAuthorizationReform.YearFact,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept, dbo.TblCardAuthorizationReform.EndDate,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Month_Day, dbo.TblCardAuthorizationReform.Granty, dbo.TblCardAuthorizationReform.DateStartG,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.DateEndG, dbo.TblCardAuthorizationReform.CarMeter, dbo.TblCardAuthorizationReform.LongGranty,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.PayFirst, dbo.TblCardAuthorizationReform.AmountAccept, dbo.TblCardAuthorizationReform.Complaint,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Noteinitial, dbo.TblCardAuthorizationReform.Shaseh, dbo.TblCardAuthorizationReform.NotAccept,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.RecordeTime, dbo.TblCardAuthorizationReform.typerequest, dbo.TblCardAuthorizationReform.FitterID,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.mobile, dbo.TblCardAuthorizationReform.Cash, dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.box, dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.boxzip, dbo.TblCardAuthorizationReform.codereg, dbo.TblCardAuthorizationReform.typereg,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.DateEnter, dbo.TblCardAuthorizationReform.DateExit,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies, dbo.TblCardAuthorizationReform.driver,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit, dbo.TblCardAuthorizationReform.TimeAcutExite,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.ClientCode, dbo.TblCardAuthorizationReform.ShowPriceOrder,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.WorkOrder, dbo.TblCardAuthorizationReform.AuthoOrder, dbo.TblCardAuthorizationReform.CarMetarOut,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.LastWorOrder"
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReform LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblColor ON dbo.TblCardAuthorizationReform.ColorID = dbo.TblColor.Id "

StrSQL = StrSQL & "                      LEFT OUTER JOIN "
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblCardAuthorizationReform.CusID = dbo.TblCustemers.CusID "

    BolBegine = False
    StrWhere = ""
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ID>=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    ''///////////////
        If val(Me.TxtAuthoOrder.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.AuthoOrder >=" & val(Me.TxtAuthoOrder.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.AuthoOrder>=" & val(Me.TxtAuthoOrder.Text) & ""
        End If
    End If
    
     
     
     If ChAccept.value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Accept =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Accept =1"
        End If
    
     End If
      
     If CheckBox1.value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.wait =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.wait =1"
        End If
    
     End If
            
          If CheckBox2.value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.notAcepted =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.notAcepted =1"
        End If
    
     End If
     
    If val(Me.TxtAuthoOrderTo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.AuthoOrder <=" & val(Me.TxtAuthoOrderTo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.AuthoOrder <=" & val(Me.TxtAuthoOrderTo.Text) & ""
        End If
    End If
    ''//////////////////
     If val(Me.TxtWorkOrder.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.WorkOrder >=" & val(Me.TxtWorkOrder.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.WorkOrder>=" & val(Me.TxtWorkOrder.Text) & ""
        End If
    End If
    If val(Me.TxtWorkOrderTo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.WorkOrder <=" & val(Me.TxtWorkOrderTo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.WorkOrder <=" & val(Me.TxtWorkOrderTo.Text) & ""
        End If
    End If
    ''/////////////////
        If val(Me.TxtShowPriceOrder.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ShowPriceOrder >=" & val(Me.TxtShowPriceOrder.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ShowPriceOrder>=" & val(Me.TxtShowPriceOrder.Text) & ""
        End If
    End If
    If val(Me.TxtShowPriceOrderTo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ShowPriceOrder <=" & val(Me.TxtShowPriceOrderTo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ShowPriceOrder <=" & val(Me.TxtShowPriceOrderTo.Text) & ""
        End If
    End If
    '///////////////////
         If TxtClientName.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND (dbo.TblCardAuthorizationReform.ClientName like '%" & Me.TxtClientName.Text & "%' Or TblCustemers.CusNamee like '%" & Trim(Me.TxtClientName.Text) & "%' Or TblCustemers.CusName like '%" & Trim(Me.TxtClientName.Text) & "%' )"
        Else
            BolBegine = True
            StrWhere = " Where (dbo.TblCardAuthorizationReform.ClientName like '%" & Me.TxtClientName.Text & "%' Or TblCustemers.CusNamee like '%" & Trim(Me.TxtClientName.Text) & "%' Or TblCustemers.CusName like '%" & Trim(Me.TxtClientName.Text) & "%' )"
        End If
    End If
     If Me.TXTCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ClientCode like '%" & Me.TXTCode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ClientCode like '%" & Me.TXTCode.Text & "%'"
        End If
    End If
    '''////////////////
     If TxtPhone.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Telephone like '%" & Me.TxtPhone.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Telephone like '%" & Me.TxtPhone.Text & "%'"
        End If
    End If
    ''''''''''''''/
    
         If TxtMobile.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.mobile like '%" & Me.TxtMobile.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.mobile like '%" & Me.TxtMobile.Text & "%'"
        End If
    End If
    '////////////////////
    
          If TxtEmail.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.email like '%" & Me.TxtEmail.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.email like '%" & Me.TxtEmail.Text & "%'"
        End If
    End If
    '''''''''''''''''/
          If TxtFax.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.fax like '%" & Me.TxtFax.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.fax like '%" & Me.TxtFax.Text & "%'"
        End If
    End If
    ''''''''/////////////////
          If Me.TxtAdress.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.address like '%" & Me.TxtAdress.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.address like '%" & Me.TxtAdress.Text & "%'"
        End If
    End If
    ''''''''''''///////
          If TxtBox.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.box like '%" & Me.TxtBox.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.box like '%" & Me.TxtBox.Text & "%'"
        End If
    End If
'////////////////////////Data of cars

 If Trim(TxtPlateNO.Text) <> "" Then
 
    Dim myCharArry() As String
    'Dim i As Integer
    ReDim myCharArry(Len(TxtPlateNO.Text) - 1)
    
   Dim mTxtPlateNo As String

    For i = 1 To Len(TxtPlateNO.Text)

        myCharArry(i - 1) = mId(TxtPlateNO.Text, i, 1)
        If IsNumeric(myCharArry(i - 1)) Then
            mTxtPlateNo = mTxtPlateNo & "%"
        End If
        mTxtPlateNo = mTxtPlateNo & myCharArry(i - 1)
    Next i

        If BolBegine = True Then
           ' StrWhere = StrWhere & " AND (dbo.TblCardAuthorizationReform.PlateNo like N'%" & Trim(TxtPlateNo.Text) & "%'"
           ' StrWhere = StrWhere & " Or dbo.TblCardAuthorizationReform.PlateNo like N'%" & Trim(TxtPlateNo.Text) & "%' )"
          '  StrWhere = StrWhere & "  pLATEnO like N'%" & Trim(mTxtPlateNo) & "%'"
            StrWhere = StrWhere & "  AND TblCardAuthorizationReform.ClientCode In ( Select  Code FROM TblCustemers AS tc Where"
            StrWhere = StrWhere & " tc.CusID In (Select CustomerID From TblCusCar Where "
            StrWhere = StrWhere & " (BoardNO like N'%" & Trim(mTxtPlateNo) & "%' Or  nBoardNo like N'%" & Trim(mTxtPlateNo) & "%') "
            StrWhere = StrWhere & "      AND ("
            StrWhere = StrWhere & "           TblCardAuthorizationReform.pLATEnO ="
            StrWhere = StrWhere & "          TblCusCar.nBoardNo"
            StrWhere = StrWhere & "         OR TblCardAuthorizationReform.pLATEnO"
            StrWhere = StrWhere & "        = TblCusCar.BoardNO"
            StrWhere = StrWhere & "       )))"
            
            
        Else
            BolBegine = True
           ' StrWhere = StrWhere & " Where (dbo.TblCardAuthorizationReform.PlateNo like N'%" & Trim(TxtPlateNo.Text) & "%'"
           ' StrWhere = StrWhere & " Or dbo.TblCardAuthorizationReform.PlateNo like N'%" & Trim(TxtPlateNo.Text) & "%' )"
            
            StrWhere = StrWhere & " where TblCardAuthorizationReform.ClientCode In ( Select  Code FROM TblCustemers AS tc Where"
            StrWhere = StrWhere & " tc.CusID In (Select CustomerID From TblCusCar Where "
            StrWhere = StrWhere & " (BoardNO like N'%" & Trim(mTxtPlateNo) & "%' Or  nBoardNo like N'%" & Trim(mTxtPlateNo) & "%') "
            StrWhere = StrWhere & "      AND ("
            StrWhere = StrWhere & "           TblCardAuthorizationReform.pLATEnO ="
            StrWhere = StrWhere & "          TblCusCar.nBoardNo"
            StrWhere = StrWhere & "         OR TblCardAuthorizationReform.pLATEnO"
            StrWhere = StrWhere & "        = TblCusCar.BoardNO"
            StrWhere = StrWhere & "       )))"
            

        End If
    End If
    ''''''''''''''''//////////////
     If TxtSahseh.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Shaseh like '%" & Me.TxtSahseh.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Shaseh like '%" & Me.TxtSahseh.Text & "%'"
        End If
    End If
    ''''''''''''''''//////////////
     If TxtReg.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.codereg like '%" & Me.TxtReg.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.codereg like '%" & Me.TxtReg.Text & "%'"
        End If
    End If
    ''''''''''''''''//////////////
     If TxtDoor.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.codedoor like '%" & Me.TxtDoor.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.codedoor like '%" & Me.TxtDoor.Text & "%'"
        End If
    End If
                            ''''''''''''''''///////////////////////‘ﬂÊÏ «·⁄„Ì·
     If TxtComplaint.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Complaint like '%" & Me.TxtComplaint.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Complaint like '%" & Me.TxtComplaint.Text & "%'"
        End If
    End If
    ''''''''''''''''//////////////
    If TxtNoteIntial1.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Noteinitial like '%" & Me.TxtNoteIntial1.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Noteinitial like '%" & Me.TxtNoteIntial1.Text & "%'"
        End If
    End If
    ''''''''''''''''//////////////
   If Me.DatacomUser.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.FitterID=" & Me.DatacomUser.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.FitterID=" & Me.DatacomUser.BoundText & ""
        End If
    End If
    '''''///////////////

   If Me.DcbCarType.BoundText <> "" Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.CarTypeID=" & Me.DcbCarType.BoundText & ""
        Else
            BolBegine = True
           StrWhere = " Where    dbo.TblCardAuthorizationReform.CarTypeID=" & Me.DcbCarType.BoundText & ""
        End If
    End If
  ''''''''''  /////////////////////
     If Me.DcbCarModel.BoundText <> "" Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.CarModelID=" & Me.DcbCarModel.BoundText & ""
        Else
            BolBegine = True
           StrWhere = " Where    dbo.TblCardAuthorizationReform.CarModelID=" & Me.DcbCarModel.BoundText & ""
        End If
    End If
    ''''''''''/////////
     If Me.DcbyearFactor.Text <> "" Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.YearFact= '" & Me.DcbyearFactor.Text & "'"
        Else
            BolBegine = True
           StrWhere = " Where    dbo.TblCardAuthorizationReform.YearFact='" & Me.DcbyearFactor.Text & "'"
        End If
    End If
     If Me.DcbColor.BoundText <> "" Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.ColorID=" & Me.DcbColor.BoundText & ""
        Else
            BolBegine = True
           StrWhere = " Where    dbo.TblCardAuthorizationReform.ColorID=" & Me.DcbColor.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReform.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
              .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("ClientCode").value), "", rs("ClientCode").value)
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                .TextMatrix(i, .ColIndex("mobile")) = IIf(IsNull(rs("mobile").value), "", rs("mobile").value)
                .TextMatrix(i, .ColIndex("box")) = IIf(IsNull(rs("box").value), "", rs("box").value)
                .TextMatrix(i, .ColIndex("fax")) = IIf(IsNull(rs("fax").value), "", rs("fax").value)
                .TextMatrix(i, .ColIndex("email")) = IIf(IsNull(rs("email").value), "", rs("email").value)
                .TextMatrix(i, .ColIndex("address")) = IIf(IsNull(rs("address").value), "", rs("address").value)
                
                .TextMatrix(i, .ColIndex("shasehno")) = IIf(IsNull(rs("Shaseh").value), "", rs("Shaseh").value)
                .TextMatrix(i, .ColIndex("codereg")) = IIf(IsNull(rs("codereg").value), "", rs("codereg").value)
                .TextMatrix(i, .ColIndex("codedoor")) = IIf(IsNull(rs("codedoor").value), "", rs("codedoor").value)
               .TextMatrix(i, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
                .TextMatrix(i, .ColIndex("fitter")) = IIf(IsNull(rs("address").value), "", rs("address").value)
                .TextMatrix(i, .ColIndex("Complaint")) = IIf(IsNull(rs("Complaint").value), "", rs("Complaint").value)
               .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("Noteinitial").value), "", rs("Noteinitial").value)
              .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(i, .ColIndex("model")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
                .TextMatrix(i, .ColIndex("color")) = IIf(IsNull(rs("Colname").value), "", rs("Colname").value)
               .TextMatrix(i, .ColIndex("yearfact")) = IIf(IsNull(rs("YearFact").value), "", rs("YearFact").value)
               '''
               .TextMatrix(i, .ColIndex("ShowPriceOrder")) = IIf(IsNull(rs("ShowPriceOrder").value), "", rs("ShowPriceOrder").value)
               .TextMatrix(i, .ColIndex("WorkOrder")) = IIf(IsNull(rs("WorkOrder").value), "", rs("WorkOrder").value)
               .TextMatrix(i, .ColIndex("AuthoOrder")) = IIf(IsNull(rs("AuthoOrder").value), "", rs("AuthoOrder").value)
               
                .TextMatrix(i, .ColIndex("Accept")) = IIf(IsNull(rs("Accept").value), "", rs("Accept").value)
                 .TextMatrix(i, .ColIndex("wait")) = IIf(IsNull(rs("wait").value), "", rs("wait").value)
                  .TextMatrix(i, .ColIndex("notAcepted")) = IIf(IsNull(rs("notAcepted").value), "", rs("notAcepted").value)
               
               
               
               
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Search CardAuthorizationReform"

Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(9).Caption = "Telephone"
lbl(11).Caption = "Mobile"
Me.lblemail.Caption = "Email"
lbl(17).Caption = "Fax"
Me.lbladdress.Caption = "Address"
lbl(18).Caption = "Mailbox"
lbl(22).Caption = "Fitter"
Me.lblType.Caption = "Type"
lbl(13).Caption = "Year"
lbl(12).Caption = "Model"
lbl(14).Caption = "Color"
lbl(8).Caption = "Chassis No"
lbl(0).Caption = "PlateNo"
lbl(2).Caption = "Total"
lbl(15).Caption = "Record No"
lbl(16).Caption = "Door No"
lbl(20).Caption = "Remark"
lbl(19).Caption = "Customer complaint"
Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
Fra(2).Caption = "Data of Car"
Fra(1).Caption = "Data of Client"
lbl(25).Caption = "To"
lbl(24).Caption = "From"

lbl(26).Caption = "To"
lbl(27).Caption = "From"

lbl(6).Caption = "To"
lbl(5).Caption = "From"

lbl(21).Caption = "To"
lbl(23).Caption = "From"
Frame1.Caption = "Offer price"
Frame2.Caption = "Authorization Reform"
Frame1.Caption = "Offer price"
Frame3.Caption = "Job order"
     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("ClientName")) = "ClientName"
        .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
       .TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
       .TextMatrix(0, .ColIndex("ShowPriceOrder")) = "Show Price Order"
       .TextMatrix(0, .ColIndex("AuthoOrder")) = "Autho Order"
       .TextMatrix(0, .ColIndex("WorkOrder")) = "Job Order"
       
       '''''''/////////////////////
          
                  .TextMatrix(0, .ColIndex("mobile")) = "Mobile"
                .TextMatrix(0, .ColIndex("box")) = "Mailbox"
                .TextMatrix(0, .ColIndex("fax")) = "Fax"
                .TextMatrix(0, .ColIndex("email")) = "Email"
                .TextMatrix(0, .ColIndex("address")) = "Address"
                .TextMatrix(0, .ColIndex("code")) = "Customer Code"
                
                .TextMatrix(0, .ColIndex("shasehno")) = "Chassis No."
                .TextMatrix(0, .ColIndex("codereg")) = "Record No"
                .TextMatrix(0, .ColIndex("codedoor")) = "Door No"
                .TextMatrix(0, .ColIndex("fitter")) = "Fitter"
                .TextMatrix(0, .ColIndex("Complaint")) = "Customer complaint"
               .TextMatrix(0, .ColIndex("remark")) = "Remarks"
              .TextMatrix(0, .ColIndex("type")) = "Car Type"
                .TextMatrix(0, .ColIndex("model")) = "Model"
                .TextMatrix(0, .ColIndex("color")) = "Color"
               .TextMatrix(0, .ColIndex("yearfact")) = "Year"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub



