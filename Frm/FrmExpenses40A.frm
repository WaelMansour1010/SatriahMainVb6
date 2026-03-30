VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses40A 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‘«‘… «·≈÷«›«  ··≈’Ê·"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   HelpContextID   =   280
   Icon            =   "FrmExpenses40A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "FrmExpenses40A.frx":6852
   RightToLeft     =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   15270
   Begin VDSCOMBOLibCtl.SmartCombo CboDes 
      Height          =   135
      Left            =   3600
      TabIndex        =   160
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1980958603
      _cy             =   1980956910
      Alignment       =   0
      Appearance      =   1
      AutoSearch      =   0   'False
      BackColor       =   16777215
      BackgroundColor =   14871017
      BorderColor     =   0
      BorderVisible   =   -1  'True
      Caption         =   "SmartCombo1"
      CaptionAlignment=   4
      CaptionBackColor=   14871017
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
      PictureAlignment=   5
      PictureBackColor=   16777215
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
      Visible0        =   -1  'True
      Width0          =   90
      Enabled1        =   -1  'True
      Position1       =   1
      Tip1            =   "Picture"
      Visible1        =   0   'False
      Width1          =   32
      Enabled2        =   -1  'True
      Position2       =   2
      Tip2            =   "Check Box (Space, Ctrl + Space)"
      Visible2        =   0   'False
      Width2          =   16
      Enabled3        =   -1  'True
      Position3       =   3
      Tip3            =   "Edit Box"
      Visible3        =   -1  'True
      Width3          =   8
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
      Tip9            =   "Down Arrow (Alt + Down, F4)"
      Visible9        =   -1  'True
      Width9          =   16
      Enabled10       =   -1  'True
      Position10      =   10
      Tip10           =   "Right Arrow (Alt + >)"
      Visible10       =   0   'False
      Width10         =   16
   End
   Begin VB.ComboBox CboPaymentType1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmExpenses40A.frx":835C9
      Left            =   15960
      List            =   "FrmExpenses40A.frx":835CB
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   158
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   153
      Top             =   840
      Width           =   15135
      Begin VB.TextBox TxtVouSerial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   5160
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   197
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox TxtOderNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7620
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   196
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox DcbBasedOn 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "FrmExpenses40A.frx":835CD
         Left            =   6600
         List            =   "FrmExpenses40A.frx":835CF
         RightToLeft     =   -1  'True
         TabIndex        =   194
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   12480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox CboType 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "FrmExpenses40A.frx":835D1
         Left            =   10080
         List            =   "FrmExpenses40A.frx":835D3
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "CboType"
         Top             =   480
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   10080
         TabIndex        =   2
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   129826817
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmExpenses40A.frx":835D5
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSDataListLib.DataCombo DCbAccount 
         Bindings        =   "FrmExpenses40A.frx":835EA
         Height          =   315
         Left            =   240
         TabIndex        =   206
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õ”«»"
         Height          =   255
         Index           =   1
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   207
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   285
         Index           =   56
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   195
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lbltoday 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÌÊ„"
         Height          =   285
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   159
         Top             =   120
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·”‰œ"
         Height          =   285
         Index           =   23
         Left            =   14280
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   285
         Index           =   4
         Left            =   14400
         RightToLeft     =   -1  'True
         TabIndex        =   155
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   147
      Top             =   8040
      Width           =   15375
      Begin VB.Frame Frame11 
         Caption         =   "»Ì«‰«  „Õ«”»Ì…"
         Height          =   735
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   201
         Top             =   120
         Width           =   7335
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   198
         Top             =   1200
         Width           =   2505
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   11760
         TabIndex        =   30
         Top             =   240
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdAttach 
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   240
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
         ButtonImage     =   "FrmExpenses40A.frx":835FF
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
         CausesValidation=   0   'False
         Height          =   375
         Index           =   10
         Left            =   4080
         TabIndex        =   199
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
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
         ButtonImage     =   "FrmExpenses40A.frx":89E61
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         Height          =   255
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   200
         Top             =   1200
         Width           =   735
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
         Height          =   255
         Index           =   7
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "/"
         Height          =   315
         Index           =   6
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   151
         Top             =   240
         Width           =   165
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   240
         Width           =   405
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   255
         Index           =   8
         Left            =   14040
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   146
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TXT_A_NoteID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   16560
      RightToLeft     =   -1  'True
      TabIndex        =   144
      Text            =   "Text8"
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   17040
      RightToLeft     =   -1  'True
      TabIndex        =   143
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   142
      Top             =   7320
      Width           =   15375
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   11400
         TabIndex        =   21
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   10200
         TabIndex        =   22
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   9000
         TabIndex        =   23
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         CausesValidation=   0   'False
         Height          =   375
         Index           =   3
         Left            =   7800
         TabIndex        =   24
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   6720
         TabIndex        =   25
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         CausesValidation=   0   'False
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   29
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton CmdHelp 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   5640
         TabIndex        =   26
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         CausesValidation=   0   'False
         Height          =   375
         Index           =   8
         Left            =   4440
         TabIndex        =   27
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
   End
   Begin VB.TextBox TxtLoseProfitValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   141
      Text            =   "Text20"
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtFASalesPrice 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   140
      Text            =   "Text20"
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox XPTxtValView 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   6915
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   6135
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   840
      Width           =   15375
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2E9E9&
         Height          =   1815
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   4200
         Width           =   15255
         Begin VB.TextBox txt_general_des 
            Alignment       =   1  'Right Justify
            Height          =   1125
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   192
            Top             =   480
            Width           =   7395
         End
         Begin VB.TextBox TxtFixeNewValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   191
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox TxtQstNewValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   190
            Top             =   840
            Width           =   1635
         End
         Begin VB.TextBox TxtQstNewNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   189
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox TxtFixeIncValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9840
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   188
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox TxtQstIncValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9840
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   187
            Top             =   840
            Width           =   1635
         End
         Begin VB.TextBox TxtQstIncNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9840
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   186
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox TxtFixeCurValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   185
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox TxtQstCurValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   184
            Top             =   840
            Width           =   1635
         End
         Begin VB.TextBox TxtQstCurNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   183
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ  "
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   20
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·«’·"
            Height          =   285
            Index           =   55
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   1200
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
            Height          =   285
            Index           =   54
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ"
            Height          =   285
            Index           =   53
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÃœÌœ"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   52
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·“Ì«œ…"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   51
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   50
            Left            =   11520
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   240
            Width           =   1755
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2E9E9&
         Caption         =   "  »Ì«‰«  «·«÷«›…"
         Height          =   3255
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   161
         Top             =   960
         Width           =   3495
         Begin VB.TextBox TxtQstNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   840
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   172
            Top             =   1800
            Width           =   1155
         End
         Begin MSComCtl2.DTPicker SatrtDate 
            Height          =   315
            Left            =   360
            TabIndex        =   169
            Top             =   2640
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Format          =   150405121
            CurrentDate     =   38784
         End
         Begin VB.TextBox TxtQstValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   167
            Top             =   2280
            Visible         =   0   'False
            Width           =   1635
         End
         Begin XtremeSuiteControls.RadioButton Distrbute 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   165
            Top             =   1080
            Width           =   3255
            _Version        =   786432
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Ê“Ì⁄ ﬁÌ„… «·«÷«›… ⁄·Ï «·«ﬁ”«ÿ «·„ »ﬁÌ…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtAddValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   164
            Top             =   600
            Width           =   1635
         End
         Begin MSComCtl2.DTPicker DateAdd 
            Height          =   315
            Left            =   360
            TabIndex        =   163
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Format          =   150405121
            CurrentDate     =   38784
         End
         Begin XtremeSuiteControls.RadioButton Distrbute 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   166
            Top             =   1440
            Width           =   3255
            _Version        =   786432
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "“Ì«œ… ⁄„— «·«’· Ê«Õ ”«» ﬁ”ÿ «·«Â·«ﬂ ⁄·Ï"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ”ÿ"
            Height          =   285
            Index           =   49
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
            Height          =   285
            Index           =   48
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   2280
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·«÷«›…"
            Height          =   285
            Index           =   45
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì»œ√ „‰  «—ÌŒ"
            Height          =   285
            Index           =   47
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   2640
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«÷«›…"
            Height          =   285
            Index           =   46
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   240
            Width           =   1755
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·ﬁÌ„Â «·„÷«›Â"
         Height          =   735
         Left            =   -4560
         RightToLeft     =   -1  'True
         TabIndex        =   137
         Top             =   360
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   139
            Top             =   240
            Width           =   2475
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„Â «·„÷«›…"
            Height          =   285
            Index           =   44
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Caption         =   " »Ì«‰«  «·√’· «·„÷«›"
         ForeColor       =   &H00C00000&
         Height          =   3255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   960
         Width           =   6135
         Begin VB.TextBox TxtAssesetCode1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   174
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox TxtNuminstallmCurr2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   2760
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmTotal2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1680
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmRemin2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   2400
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmExcu2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   2040
            Width           =   3555
         End
         Begin VB.TextBox TxtCurrentValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1320
            Width           =   3555
         End
         Begin VB.TextBox TxtAccDepre2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   960
            Width           =   3555
         End
         Begin VB.TextBox TxtPurchasePrice2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   600
            Width           =   3555
         End
         Begin MSDataListLib.DataCombo DCAccounts 
            Height          =   315
            Left            =   6120
            TabIndex        =   135
            Top             =   720
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSandAdd 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12640511
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·«’· «·„÷«›"
            Height          =   285
            Index           =   42
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ﬁ”ÿ «·Õ«·Ì…"
            Height          =   285
            Index           =   41
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   2760
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·«Ã„«·Ì…"
            Height          =   285
            Index           =   40
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   1680
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„ »ﬁÌ…"
            Height          =   285
            Index           =   39
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   2400
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„‰›–…"
            Height          =   285
            Index           =   34
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   2040
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
            Height          =   285
            Index           =   32
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   1320
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
            Height          =   285
            Index           =   31
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·‘—«¡"
            Height          =   285
            Index           =   26
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   600
            Width           =   2115
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "  »Ì«‰«  «·√’· «·√”«”Ì"
         Height          =   3255
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   960
         Width           =   5415
         Begin VB.TextBox TxtAssesetCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   175
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox TxtPurchasePrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   600
            Width           =   3555
         End
         Begin VB.TextBox TxtAccDepre 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   960
            Width           =   3555
         End
         Begin VB.TextBox TxtCurrentValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1320
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmExcu 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   2040
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmRemin 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   2400
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1680
            Width           =   3555
         End
         Begin VB.TextBox TxtNuminstallmCurr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2760
            Width           =   3555
         End
         Begin MSDataListLib.DataCombo DcFixedAssets 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12640511
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·«’· «·«”«”Ì"
            Height          =   285
            Index           =   27
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·‘—«¡"
            Height          =   285
            Index           =   28
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
            Height          =   285
            Index           =   29
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   960
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
            Height          =   285
            Index           =   30
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   1320
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„‰›–…"
            Height          =   285
            Index           =   35
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   2040
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„ »ﬁÌ…"
            Height          =   285
            Index           =   36
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   2400
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·«Ã„«·Ì…"
            Height          =   285
            Index           =   37
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1680
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ﬁ”ÿ «·Õ«·Ì…"
            Height          =   285
            Index           =   38
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   2760
            Width           =   1755
         End
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   16080
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   960
         Width           =   3315
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   2925
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   1320
         Width           =   4635
         Begin VB.TextBox TXTBankName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   960
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1320
            Width           =   3285
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   600
            Width           =   705
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   960
            Width           =   705
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   30
            TabIndex        =   91
            Top             =   2100
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   150142977
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   30
            TabIndex        =   92
            Top             =   960
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   0
            TabIndex        =   93
            Top             =   600
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCVendor 
            Height          =   315
            Left            =   0
            TabIndex        =   94
            Top             =   240
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccounts1 
            Height          =   315
            Left            =   0
            TabIndex        =   113
            Top             =   2520
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcChequeBox 
            Height          =   315
            Left            =   0
            TabIndex        =   115
            Top             =   1680
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
            Height          =   285
            Index           =   43
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   285
            Index           =   33
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·Œ“Ì‰…"
            Height          =   285
            Index           =   16
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   22
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   16560
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   4560
         Width           =   2715
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   16920
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   1590
         Width           =   2655
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   82
         Text            =   "Text2"
         Top             =   1110
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   16320
         TabIndex        =   101
         Top             =   1140
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmExpenses40A.frx":906C3
         Height          =   315
         Left            =   16320
         TabIndex        =   102
         Top             =   780
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   14880
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "FrmExpenses40A.frx":906D8
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
         Left            =   14400
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·»Ì⁄"
         Height          =   255
         Index           =   15
         Left            =   15960
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
         Height          =   285
         Index           =   0
         Left            =   15840
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   810
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   285
         Index           =   21
         Left            =   15720
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   735
         Left            =   10920
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ… Â«„…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   24
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   6840
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   540
         Index           =   25
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   6600
         Width           =   1695
      End
   End
   Begin VB.OptionButton OptSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   16920
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   15840
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2340
      Left            =   16680
      TabIndex        =   62
      Top             =   4440
      Visible         =   0   'False
      Width           =   10755
      _cx             =   18971
      _cy             =   4128
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40A.frx":90C62
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   2550
         RightToLeft     =   -1  'True
         ScaleHeight     =   3915
         ScaleWidth      =   9405
         TabIndex        =   67
         Top             =   810
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CommandButton Command3 
            Caption         =   "Call des"
            Height          =   255
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add des"
            Height          =   255
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   3600
            Width           =   1350
         End
         Begin VB.TextBox txtcodesub 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox TxtDese 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   68
            Top             =   2040
            Width           =   8955
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3900
            Left            =   120
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   0
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
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   73
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
               TabIndex        =   74
               Top             =   0
               Width           =   2445
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   495
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   3480
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
         Height          =   1215
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   3720
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "‰”Œ"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16080
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   720
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
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   9780
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   54
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   47
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15720
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   16440
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   7140
      Width           =   2145
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Left            =   0
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   15255
      _cx             =   26908
      _cy             =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      Caption         =   "‘«‘… «·≈÷«›«  ··≈’Ê·            "
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   33
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
         ButtonImage     =   "FrmExpenses40A.frx":90F3E
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
         TabIndex        =   35
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
         ButtonImage     =   "FrmExpenses40A.frx":912D8
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
         TabIndex        =   32
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
         ButtonImage     =   "FrmExpenses40A.frx":91672
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
         TabIndex        =   34
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
         ButtonImage     =   "FrmExpenses40A.frx":91A0C
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
         Left            =   12840
         Top             =   480
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   4080
         Picture         =   "FrmExpenses40A.frx":91DA6
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image ImgFavoritesdd 
         Height          =   615
         Left            =   6360
         Picture         =   "FrmExpenses40A.frx":95A0E
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label LblShortcutKeys 
         Alignment       =   2  'Center
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
         TabIndex        =   45
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   16200
      TabIndex        =   36
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
      Height          =   2340
      Left            =   16680
      TabIndex        =   55
      Top             =   4440
      Visible         =   0   'False
      Width           =   10800
      _cx             =   19050
      _cy             =   4128
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40A.frx":9ADC3
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
         Left            =   240
         RightToLeft     =   -1  'True
         ScaleHeight     =   1635
         ScaleWidth      =   2925
         TabIndex        =   57
         Top             =   960
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
            TabIndex        =   58
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
            TabIndex        =   59
            Top             =   0
            Width           =   2445
         End
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   9600
      TabIndex        =   56
      Top             =   9360
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
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
      MICON           =   "FrmExpenses40A.frx":9AF29
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
      Index           =   9
      Left            =   5640
      TabIndex        =   60
      Top             =   9240
      Visible         =   0   'False
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
      Left            =   9600
      TabIndex        =   61
      Tag             =   "Delete Row"
      Top             =   9360
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
      MICON           =   "FrmExpenses40A.frx":9AF45
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   2340
      Left            =   480
      TabIndex        =   111
      Top             =   9480
      Width           =   10800
      _cx             =   19050
      _cy             =   4128
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   2
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40A.frx":9AF61
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
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   145
      Top             =   0
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc detect_no 
      Height          =   585
      Left            =   0
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
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   6900
      Width           =   12015
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   6960
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   2520
      Width           =   1515
   End
End
Attribute VB_Name = "FrmExpenses40A"
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
Dim BolEditOnMainAccounts As Boolean
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim group_id As Integer
Dim ProfitOrLose As Integer
Dim ProfitOrLoseValue As Double
Dim line_no As Integer
Dim LoseProfitValue As Double
Dim txtmyDes As String
Dim txtmyDesE As String
'    Function saveChequeBoxContents(NoteID As Double)
'    Dim i As Integer
'    Dim rs As New ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
'    Cn.Execute StrSQL, , adExecuteNoRecords
'    If val(DcChequeBox.BoundText) = 0 Then Exit Function
'
'   ' rs.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
'   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
''
 '   rs.AddNew
 '   rs("noteid").value = NoteID
 '   rs("ChequeBoxID").value = val(DcChequeBox.BoundText)
 ''
  '  rs("RecordDate").value = XPDtbTrans.value
  '  rs("DueDate").value = DtpChequeDueDate.value
  '  rs("BankName").value = TXTBankName.text
  '  rs("ChequeNo").value = TxtChequeNumber.text
  '  rs("ChequeValue").value = val(XPTxtVal.text)
  '
  '  rs("Remarks").value = DcboCreditSide.text
  '  rs("Deposited").value = 0
  '  rs("Collected").value = 0
  '  rs("CreditAccount").value = (DcboCreditSide.BoundText)
  '  rs.update
  '
  '  rs.Close
'End Function
'
'Function saveChequeBoxContents1(NoteID As Double)
'
'    If SystemOptions.banks_Accounts3 = False Then Exit Function
'    Dim i As Integer
'    Dim rs As New ADODB.Recordset
'
'    Dim StrSQL As String
'
'    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
'    Cn.Execute StrSQL, , adExecuteNoRecords
'
'    rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    If CboPaymentType.ListIndex = 1 Then
'        rs.AddNew
'        rs("noteid").value = NoteID
'
'        rs("RecordDate").value = XPDtbTrans.value
'        rs("DueDate").value = DtpChequeDueDate.value
'        rs("BankID").value = val(DcboBankName.BoundText)
'        rs("BankName").value = DcboBankName.text
'
'        rs("ChequeNo").value = TxtChequeNumber.text
'        rs("ChequeValue").value = val(XPTxtVal.text)
'
'        rs("Remarks").value = Me.DcboDebitSide.text
'        rs("Payed").value = 0
'
'        rs("DepitAccount").value = (DcboDebitSide.BoundText)
'        rs("notes_all").value = NoteID
'
'        rs.update
'    End If
'
'    rs.Close
'End Function

Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Else
        MsgBox "It can not be the cost of distribution centers because you chose in distribution", vbCritical
        End If
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.Text)
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
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If

    marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
    marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
    marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
    marakes_taklefa_tawze3.Adodc3.Refresh
    '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_AfterAutoCloseUp()
   
    CboDes.Visible = False
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1, "0604201601"

End Sub



Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbBasedOn_Change()
Me.DcbAccount.Visible = False
Label1(1).Visible = False
TxtVouSerial.Visible = False
TxtAddValue.Enabled = False
If val(DcbBasedOn.ListIndex) = 3 Then
Me.DcbAccount.Visible = True
Label1(1).Visible = True
TxtAddValue.Enabled = True
Else
TxtVouSerial.Visible = True
End If
TxtVouSerial.Text = ""
End Sub

Private Sub DcbBasedOn_Click()
DcbBasedOn_Change
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbSandAdd_Change()
    If val(DcbSandAdd.BoundText) = 0 Then Exit Sub
  
    DcbSandAdd_Click (0)
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
Dim i As Integer
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset

    Dim LngDevID As Long
    Dim Msg As String
   Dim NotValue As Double
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    Dim DebitAccount As String
    Dim DebitAccount2 As String
    Dim Account_code As String
    Dim GroupID As Integer
        Msg = Ele.Caption & " —ﬁ„ «·”‰œ " & TxtSerial1 & " » «—ÌŒ" & XPDtbTrans.value
 
 
      '  If SystemOptions.FAAddtionCreateAccount = True Then
    '    DebitAccount = GetFixedAssetAddAccount(val(DcFixedAssets.BoundText))
       
      '   Else
     '   GetAllDataAboutFixedAsset val(DcFixedAssets.BoundText), , , , , , , , , , , , , , , , , , , , , , , , DebitAccount
'   End If
        GetAllDataAboutFixedAsset val(DcbSandAdd.BoundText), , GroupID
        GetFixedAssetsGroupAccount GroupID, , , , , , , , , , DebitAccount2
        GetFixedAssetsGroupAccount GroupID, , , , , , , , Account_code
        
        GetAllDataAboutFixedAsset val(DcFixedAssets.BoundText), , GroupID
        GetFixedAssetsGroupAccount GroupID, , , , , , , , DebitAccount
       ' Account_Code = GetFixedAssetAddAccount(val(DcbSandAdd.BoundText))
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = val(DcBranch.BoundText)
    If val(CboType.ListIndex) = 1 Then
''/////////////œ„Ã «’·
NotValue = val(TxtCurrentValue2.Text)
line_no = 1

If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, DebitAccount, Round(NotValue, 2), 0, Msg & "œ„Ã  ··«’·   " & DcFixedAssets.Text & " " & "„⁄ «·«’·" & " " & DcbSandAdd.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
NotValue = val(TxtAccDepre2.Text)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, DebitAccount2, Round(NotValue, 2), 0, Msg & "œ„Ã  ··«’·   " & DcFixedAssets.Text & " " & "„⁄ «·«’·" & " " & DcbSandAdd.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
NotValue = val(TxtAccDepre2.Text) + val(TxtCurrentValue2.Text)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code, Round(NotValue, 2), 1, Msg & "œ„Ã  ··«’·   " & DcFixedAssets.Text & " " & "„⁄ «·«’·" & " " & DcbSandAdd.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
Else
''/////////////////
Dim total_value As Double
Dim TempValue As Double
Dim Balance As String
NotValue = val(TxtAddValue.Text)
If NotValue > 0 Then
           If ModAccounts.AddNewDev(LngDevID, line_no, DebitAccount, Round(NotValue, 2), 0, Msg & "«÷«›… ﬁÌ„…  ··«’·   " & DcFixedAssets.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
sql = ""
If val(DcbBasedOn.ListIndex) = 0 Then
''''«÷«›… ﬁÌ„…

sql = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
sql = sql & " FROM         dbo.notes_all LEFT OUTER JOIN"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.A_NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
sql = sql & " WHERE     (dbo.notes_all.NoteType = 80) AND (dbo.notes_all.bill_type <> 2) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
sql = sql & "                      (dbo.notes_all.NoteSerial1 = " & val(TxtVouSerial.Text) & ")"
ElseIf val(DcbBasedOn.ListIndex) = 1 Then
sql = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
sql = sql & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS RIGHT OUTER JOIN"
sql = sql & "                      dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
sql = sql & " Where (dbo.Notes.NoteType = 5) And (dbo.Notes.NoteSerial1 = " & val(TxtVouSerial.Text) & ") And (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
End If
If sql <> "" Then
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Rs9.MoveFirst
Account_code = IIf(IsNull(Rs9("Account_Code").value), "", Rs9("Account_Code").value)

If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code, Round(NotValue, 2), 1, Msg & "«÷«›… ﬁÌ„…  ··«’·   " & DcFixedAssets.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If



End If
End If
End If
If val(DcbBasedOn.ListIndex) = 2 Then
sql = "SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value]"
sql = sql & " FROM         dbo.Notes INNER JOIN"
sql = sql & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
sql = sql & " Where (dbo.Notes.NoteSerial1 = " & val(TxtVouSerial.Text) & ") And (dbo.Notes.NoteType = 3) And (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
If sql <> "" Then
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Rs9.MoveFirst
For i = 1 To Rs9.RecordCount
Account_code = IIf(IsNull(Rs9("Account_Code").value), "", Rs9("Account_Code").value)
NotValue = IIf(IsNull(Rs9("Value").value), "", Rs9("Value").value)
If NotValue > 0 Then

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code, Round(NotValue, 2), 1, Msg & "«÷«›… ﬁÌ„…  ··«’·   " & DcFixedAssets.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
Rs9.MoveNext
Next i
End If
End If
End If
If val(DcbBasedOn.ListIndex) = 3 Then
If NotValue > 0 Then
Account_code = Me.DcbAccount.BoundText

           If ModAccounts.AddNewDev(LngDevID, line_no, Account_code, Round(NotValue, 2), 1, Msg & "«÷«›… ﬁÌ„…  ··«’·   " & DcFixedAssets.Text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
End If
End If

updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "«÷«›«  «·«’Ê· »—ﬁ„ " & TxtSerial1 & " ‰Ê⁄ «·”‰œ  " & CboType.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblAdditionsAssest"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0

 notytype = 8073
 If val(CboType.ListIndex) = 1 Then
Notevalue = val(TxtPurchasePrice2.Text) + val(TxtAccDepre2.Text)
Else
Notevalue = val(TxtAddValue.Text)
End If
 

 BranchID = val(DcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
 

     End If

End Function
Private Sub DcbSandAdd_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
    Dim AccDepreciation As Double
    Dim RemianInstallments As Double
    Dim CurrentInstalmentNo As Double
    Dim Installmentvalue As Double
    Dim NewAccDepreciation As Double
    Dim FixedAsssetid As Integer
    Dim purchaseprice As Double
    Dim FixedAssetName As String
    Dim Fullcode As String
    Dim KhordaPrice As Double

    If val(DcbSandAdd.BoundText) = 0 Then Exit Sub
    FixedAsssetid = val(DcbSandAdd.BoundText)
    Me.TxtFASalesPrice = 0

    GetFixedAssetHistory FixedAsssetid, AccDepreciation, RemianInstallments, CurrentInstalmentNo, Installmentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, , Fullcode, KhordaPrice, group_id
 
    TxtPurchasePrice2.Text = purchaseprice
    TxtAccDepre2.Text = AccDepreciation
    ' TxtCurrentValue = TxtPurchasePrice.text - (TxtAccDepre.text + KhordaPrice)
    TxtCurrentValue2 = val(TxtPurchasePrice2.Text) - val(TxtAccDepre2.Text)
     TxtNuminstallmCurr2.Text = Installmentvalue
 TxtNuminstallmRemin2.Text = RemianInstallments
 TxtNuminstallmExcu2.Text = CurrentInstalmentNo
 TxtNuminstallmTotal2.Text = RemianInstallments + CurrentInstalmentNo
 End If
    Dim AsseCode1 As String
If val(DcbSandAdd.BoundText) <> 0 Then
GetAsseteCode_ID val(DcbSandAdd.BoundText), AsseCode1, 0
TxtAssesetCode1.Text = AsseCode1
End If
End Sub

Private Sub DcChequeBox_Change()

    If DcChequeBox.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxID", val(Me.DcChequeBox.BoundText), "Account_Code1")
    End If

End Sub

Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.Text = "E" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""
        DCVendor.Text = ""
        DCAccounts1.Text = ""
        DcChequeBox.Text = ""

    End If

    If Me.CboPaymentType.ListIndex = 0 Then
        DcChequeBox.Enabled = False
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        DCAccounts1.Enabled = False
        DCAccounts1.Text = ""
    ElseIf Me.CboPaymentType.ListIndex = 1 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DcChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
        End If

        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Me.DCVendor.Enabled = False
        DCAccounts1.Enabled = False
        DCAccounts1.Text = ""

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
            lbl(19).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
        Else
            lbl(18).Caption = "Cheque No"
            lbl(19).Caption = "Due Date"
        End If
    
    ElseIf Me.CboPaymentType.ListIndex = 2 Then '⁄„Ì·
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        DcChequeBox.Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.Text = ""
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.Text = ""
        Me.DCVendor.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·ÕÊ«·… "
            lbl(19).Caption = " «—ÌŒÂ«"
        Else
            lbl(18).Caption = "Transfer  No"
            lbl(19).Caption = "Date"
        End If
      
    ElseIf Me.CboPaymentType.ListIndex = 4 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        DCVendor.BoundText = ""
        DcboBox.BoundText = ""
        DcboBankName.BoundText = ""
        DCAccounts1.Enabled = True
        DcChequeBox.Enabled = False
        '        DCAccounts1.text = ""
 
    ElseIf Me.CboPaymentType.ListIndex = 5 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        TXTBankName.Visible = False
        DcChequeBox.Enabled = False
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.Text = ""
        Me.DCVendor.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
            lbl(19).Caption = " «—ÌŒÂ  "
        Else
            lbl(18).Caption = "Cheque No"
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
        Me.DCVendor.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

'Function setfoxy()
'    Text1.text = CStr(new_id("foxy", "id", "", True))
'
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    rs("id").value = Text1.text
'
'    rs.update
'
'End Function
Private Sub CboPaymentType1_Change()
    If Me.CboPaymentType1.ListIndex = 0 Then
        Fg_Journal.Visible = True
        VSFlexGrid1.Visible = False
    ElseIf Me.CboPaymentType1.ListIndex = 1 Then
        Fg_Journal.Visible = False
        VSFlexGrid1.Visible = True
    End If
End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub CboType_Change()
    CboType_Click
End Sub

Private Sub CboType_Click()
TxtVouSerial.Visible = False
lbl(56).Visible = False
DcbBasedOn.Visible = False
TxtOderNo.Visible = False
Frame9.Visible = False
Me.DcbAccount.Visible = False
If val(Me.CboType.ListIndex) = 0 Then
Frame9.Visible = True
TxtVouSerial.Visible = True
        Frame4.Visible = False
        Frame5.Visible = True
        lbl(56).Visible = True
        DcbBasedOn.Visible = True
        Me.DcbAccount.Visible = True
      ' TxtOderNo.Visible = True
   
    Else

        Frame4.Visible = True
        Frame5.Visible = False
    
    End If

End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
       
            TxtModFlg.Text = "N"
            clear_all Me
           ' DcCostCenter.text = ""
           ' CboPaymentType1.ListIndex = 2
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
          
           ' DtpChequeDueDate.value = Date
           ' setfoxy
            Me.DcBranch.BoundText = branch_id
           ' CuurentLogdata

        Case 1
                                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            Dim Msg  As String
                    
            '                  If SystemOptions.banks_Accounts3 = True Then
            '     If ChequeBoxOperations1(Val(Me.XPTxtID)) = False Then
            '         Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
            '         Msg = Msg & Chr(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         Exit Sub
            '     End If
            ' End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
        
          '  If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         '
         '       If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
         '           Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
         '           Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
         '           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         '           Exit Sub
         '       End If
    '
    '        End If
      
            TxtModFlg.Text = "E"
        
        Case 2
                                 If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If Trim(DcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify branch"
                Else
                    Msg = "Õœœ «·›—⁄"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.DcBranch.BoundText
    If val(CboType.ListIndex) = 0 Then
    If val(TxtAddValue.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· ﬁÌ„… «·“Ì«œ…"
    Else
    MsgBox "Enter Value"
    End If
    ' TxtAddValue.SetFocus
    Exit Sub
    End If
    End If
           ' DcboBox_Change
           ' DcboBankName_Change
           ' DCVendor_Change
           '' DCAccounts1_Change
            'DcChequeBox_Change
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
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If SystemOptions.ChequeBox = True And CboPaymentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

             Load FrmSearshExpens40A
             FrmSearshExpens40A.show vbModal

        Case 6
            Unload Me

        Case 7
         '   ViewDataList

        Case 8
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report (TxtSerial.Text)

        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

          '  print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.text

        Case 10
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc TxtSerial.Text, , 200
    
    End Select

    Exit Sub
ErrTrap:
End Sub

'Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
'    hide_logo = True
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'
'    MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
'    Else
'        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
'    End If
'
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
'    'MsgBox ToHijriDate(Date)
'
'    xReport.ParameterFields(5).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 1, 2)
'    xReport.ParameterFields(6).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 4, 2)
'    xReport.ParameterFields(7).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 9, 2)
'
'    xReport.ParameterFields(8).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
'    xReport.ParameterFields(9).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
'    xReport.ParameterFields(10).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
'    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.text)
'    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
'    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.text)
'    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
'
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.Title
'    xReport.ReportAuthor = App.Title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, ""
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function
'
Function print_report(Optional NoteSerial As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

     
MySQL = MySQL & "  SELECT     dbo.TblAdditionsAssest.NoteSerial1 as  NoteSerial, dbo.TblAdditionsAssest.RecordDate, dbo.TblAdditionsAssest.UserID, dbo.TblAdditionsAssest.BranchID,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAdditionsAssest.TypeSand, dbo.TblAdditionsAssest.SandAdd,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.FixedID, dbo.FixedAssets.Name, dbo.FixedAssets.namee, FixedAssets_1.Name AS FixAddName, FixedAssets_1.namee AS FixAddNameE,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.PurchasePrice, dbo.TblAdditionsAssest.PurchasePrice2, dbo.TblAdditionsAssest.AccDepre, dbo.TblAdditionsAssest.AccDepre2,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.CurrentValue, dbo.TblAdditionsAssest.CurrentValue2, dbo.TblAdditionsAssest.NuminstallmTotal, dbo.TblAdditionsAssest.NuminstallmTotal2,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.NuminstallmExcu, dbo.TblAdditionsAssest.NuminstallmExcu2, dbo.TblAdditionsAssest.NuminstallmRemin,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.NuminstallmRemin2, dbo.TblAdditionsAssest.NuminstallmCurr, dbo.TblAdditionsAssest.NuminstallmCurr2,"
MySQL = MySQL & "                      dbo.TblAdditionsAssest.general_des "
MySQL = MySQL & " FROM         dbo.TblAdditionsAssest LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblAdditionsAssest.SandAdd = FixedAssets_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblAdditionsAssest.FixedID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblAdditionsAssest.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblAdditionsAssest.id =" & val(XPTxtID.Text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "RepAdditionFixedAssest.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "RepAdditionFixedAssest.rpt"
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
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
   ' CViewer.FireReport xReport, WindowTarget, ""
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Sub
    Dim sql As String

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If CboPaymentType1.ListIndex = 0 Then
        If Fg_Journal.Rows > 1 Then
            If Fg_Journal.Rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.Rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                        Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                    End If
                End If
            End If
        End If
            
        With Fg_Journal
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    ElseIf CboPaymentType1.ListIndex = 1 Then

        If VSFlexGrid1.Rows > 1 Then
            If VSFlexGrid1.Rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.Rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With
             
    ElseIf CboPaymentType1.ListIndex = 2 Then

        If VSFlexGrid2.Rows > 1 Then
            If VSFlexGrid2.Rows = 2 Then
                Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid2.Rows > 1 Then
                    If Me.VSFlexGrid2.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        Me.VSFlexGrid2.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid2
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With
             
    Else
 
        Exit Sub
    End If

End Sub

Private Sub DCAccounts1_Change()

    If DCAccounts1.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = DCAccounts1.BoundText
    End If

End Sub

Private Sub DCAccounts1_Click(Area As Integer)
    DCAccounts1_Change
End Sub

Private Sub DCAccounts1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 194
            
    End If

End Sub

Private Sub DcboBankName_Change()

    'On Error Resume Next
    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        If CboPaymentType.ListIndex = 3 Or CboPaymentType.ListIndex = 5 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
TxtNoteSerial.Text = ""
    TxtSerial.Text = ""
  '  TxtSerial1.Text = ""
End Sub

Private Sub DcChequeBox_Click(Area As Integer)
    DcChequeBox_Change
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 3
    End If

End Sub

Private Sub DcFixedAssets_Change()
    If val(DcFixedAssets.BoundText) = 0 Then Exit Sub
    DcFixedAssets_Click (0)
End Sub
Private Sub DcFixedAssets_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
    Dim AccDepreciation As Double
    Dim RemianInstallments As Double
    Dim CurrentInstalmentNo As Double
    Dim Installmentvalue As Double
    Dim NewAccDepreciation As Double
    Dim FixedAsssetid As Integer
    Dim purchaseprice As Double
    Dim FixedAssetName As String
    Dim Fullcode As String
    Dim KhordaPrice As Double

    If val(DcFixedAssets.BoundText) = 0 Then Exit Sub
    FixedAsssetid = val(DcFixedAssets.BoundText)
    Me.TxtFASalesPrice = 0

    GetFixedAssetHistory FixedAsssetid, AccDepreciation, RemianInstallments, CurrentInstalmentNo, Installmentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, , Fullcode, KhordaPrice, group_id
 
    TxtPurchasePrice.Text = purchaseprice
    TxtAccDepre.Text = AccDepreciation
    ' TxtCurrentValue = TxtPurchasePrice.text - (TxtAccDepre.text + KhordaPrice)
    TxtCurrentValue.Text = TxtPurchasePrice.Text - TxtAccDepre.Text
    TxtCurrentValue.Text = val(TxtCurrentValue.Text) + SmAddValue(val(XPTxtID.Text), val(Me.DcFixedAssets.BoundText))
    TxtFixeCurValue.Text = TxtCurrentValue.Text
 TxtNuminstallmCurr.Text = Installmentvalue + SmQstValue(val(XPTxtID.Text), val(Me.DcFixedAssets.BoundText))
 TxtNuminstallmRemin.Text = RemianInstallments + SmQstNo(val(XPTxtID.Text), val(Me.DcFixedAssets.BoundText))
 TxtNuminstallmExcu.Text = CurrentInstalmentNo
 TxtQstCurNo.Text = val(TxtNuminstallmRemin.Text)
 TxtNuminstallmTotal.Text = RemianInstallments + CurrentInstalmentNo
 TxtQstCurValue.Text = val(TxtNuminstallmCurr.Text)
 TxtQstValue.Text = val(TxtNuminstallmCurr.Text)
    TxtFASalesPrice_Change
   ' WriteDev
End If
   Dim AsseCode1 As String
If val(DcFixedAssets.BoundText) <> 0 Then
GetAsseteCode_ID val(DcFixedAssets.BoundText), AsseCode1, 0
TxtAssesetCode.Text = AsseCode1
End If
End Sub
Function SmQstNo(Optional ID As Double, Optional Fixed As Integer)
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = " SELECT     SUM(QstIncNo) AS SmQstIncNo"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & " Where (FixedID = " & Fixed & ") And (Effct = 1) And (id <> " & ID & ")"
sql = sql & " GROUP BY FixedID"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
SmQstNo = IIf(IsNull(Rs5("SmQstIncNo").value), 0, Rs5("SmQstIncNo").value)
Else
SmQstNo = 0
End If
End Function

Function SmQstValue(Optional ID As Double, Optional Fixed As Integer)
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = " SELECT     SUM(QstIncValue) AS SmQstIncValue"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & " Where (FixedID = " & Fixed & ") And (Effct = 1) And (id <> " & ID & ")"
sql = sql & " GROUP BY FixedID"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
SmQstValue = IIf(IsNull(Rs5("SmQstIncValue").value), 0, Rs5("SmQstIncValue").value)
Else
SmQstValue = 0
End If
End Function
Function SmAddValue(Optional ID As Double, Optional Fixed As Integer)
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = " SELECT     SUM(AddValue) AS SmQstIncValue"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & " Where (FixedID = " & Fixed & ") And (Effct = 1) And (id <> " & ID & ")"
sql = sql & " GROUP BY FixedID"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
SmAddValue = IIf(IsNull(Rs5("SmQstIncValue").value), 0, Rs5("SmQstIncValue").value)
Else
SmAddValue = 0
End If
End Function

Sub GetAsseteCode_ID(Optional ByRef ID As Double = 0, Optional ByRef Fullcode As String = "", Optional Typ As Integer = 0)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If Typ = 0 Then
sql = "select Fullcode  from FixedAssets where id=" & ID & " "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Fullcode = IIf(IsNull(Rs7("Fullcode").value), "", Rs7("Fullcode").value)
Else
Fullcode = ""
End If
Else
sql = "select ID  from FixedAssets where Fullcode='" & Fullcode & "' "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ID = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
Else
ID = 0
End If
End If
End Sub
'Function WriteDev()
'
''    If Me.TxtModFlg <> "R" Then
'
'        If SystemOptions.AssetAccount1 = True Then
'            If val(TxtFASalesPrice.text) > val(TxtCurrentValue.text) Then
'                DCAccounts.BoundText = get_FixedAsset_Account(group_id, branch_id, "Account_Code3")
'            Else
'                DCAccounts.BoundText = get_FixedAsset_Account(group_id, branch_id, "Account_Code4")
'            End If
'
'        Else
'
'            If val(TxtFASalesPrice.text) > val(TxtCurrentValue.text) Then
'                Account_Code_dynamic3 = get_account_code_branch(66, my_branch)
'
''                If Account_Code_dynamic3 = "NO branch" Then
 '                   MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
 '                   GoTo ErrTrap
 '               Else
'
'                    If Account_Code_dynamic3 = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ     Õ”«» «—»«Õ »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'                DCAccounts.BoundText = Account_Code_dynamic3
'            Else
'                Account_Code_dynamic4 = get_account_code_branch(67, my_branch)
'
'                If Account_Code_dynamic4 = "NO branch" Then
'                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                    GoTo ErrTrap
'                Else
'
'                    If Account_Code_dynamic4 = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ  Õ”«» Œ”«—… »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'                DCAccounts.BoundText = Account_Code_dynamic4
'            End If
'
'        End If
'
'    End If
'
'ErrTrap:
'End Function

Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FixedAssetsSearch.RetrunType = 1
        FixedAssetsSearch.show vbModal
  
    End If

End Sub

Private Sub DCVendor_Change()

    If DCVendor.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Text2.Text = Me.DCVendor.BoundText
End Sub

Private Sub DCVendor_Click(Area As Integer)
    DCVendor_Change
End Sub

Private Sub Distrbute_Click(Index As Integer)
TxtQstNo.Enabled = False
TxtQstValue.Enabled = False
SatrtDate.Enabled = False
If Distrbute(1).value = True Then
TxtQstNo.Enabled = True
TxtQstValue.Enabled = True
SatrtDate.Enabled = True
End If
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)

            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

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
                Dim sgl As String
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
               
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

     
End Sub

'Function calcnets()
'
'    If Me.CboPaymentType1.ListIndex = 0 Then
'
'        With Fg_Journal
'            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
'        End With
'
'    ElseIf Me.CboPaymentType1.ListIndex = 1 Then
'
'        With Me.VSFlexGrid1
'            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
'        End With
'
'    ElseIf Me.CboPaymentType1.ListIndex = 2 Then
'
'        With Me.VSFlexGrid2
'            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
'        End With

'    End If

'End Function

'Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
'                                  ByVal Col As Long, _
'                                  Cancel As Boolean)
'
'    With Fg_Journal
'
'        If Row > .FixedRows Then
'            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
'            '      Cancel = True
'            '  End If
'        End If
'
'        Select Case .ColKey(Col)
'
'            Case "value"
'                .ComboList = ""
'
'            Case "des"
'                .ComboList = ""
'                '  Cancel = True
'
'            Case "Order_No"
'                .ComboList = ""
'        End Select
'
'    End With
'
'End Sub
'
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
            CboDes.Visible = False
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
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

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
                    FrmExpensesSearch.RetrunType = 2
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
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
                StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")

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

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String

    On Error GoTo ErrTrap

  ' StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
  ' fill_combo Me.DcCostCenter, StrSQL
    
    ScreenNameArabic = "‘«‘…  «·« ÷«›… ··«’Ê·"
    ScreenNameEnglish = "Additions to Assets"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
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
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetFixedAssets Me.DcFixedAssets
    Dcombos.GetFixedAssets Me.DcbSandAdd
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.DcBranch
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
If SystemOptions.UserInterface = ArabicInterface Then
    With Me.CboType
        .Clear
        .AddItem "«÷«›… ﬁÌ„Â ·√’·"
        .AddItem "œ„Ã «’·"
         End With
      
    With DcbBasedOn
    .Clear
    .AddItem "›« Ê—… „«·Ì…"
    .AddItem "”‰œ ’—› «·„œ›Ê⁄« "
    .AddItem "”‰œ ’—›  Õ·Ì·Ì "
    .AddItem "Õ”«»"
    End With
Else

    With DcbBasedOn
    .Clear
    .AddItem "Financial Bill"
    .AddItem "Payments Voucher"
    .AddItem "Analytical  Payments Voucher"
    .AddItem "Account"
    End With
    
  With Me.CboType
        .Clear
        .AddItem "Assets Additions"
        .AddItem "Assets Merge"
    End With
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblAdditionsAssest  where 1=1 "
     StrSQL = StrSQL & "  AND      BranchID in(" & Current_branchSql & ")"
     
          If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                               
                  If SystemOptions.usertype <> UserAdminAll Then
      '  StrSQL = StrSQL & " AND   BranchID=" & Current_branch
    End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
     Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
   

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    hide_logo = False
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

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
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 1
    End If

End Sub

Private Sub TxtAddValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtFixeIncValue.Text = val(TxtAddValue.Text)
End If
End Sub

Private Sub TxtAddValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAddValue.Text, 0)
End Sub

Private Sub TxtAssesetCode_KeyPress(KeyAscii As Integer)
Dim AsseID As Double
If TxtAssesetCode.Text <> "" Then
GetAsseteCode_ID AsseID, TxtAssesetCode.Text, 1
DcFixedAssets.BoundText = AsseID
End If
End Sub

Private Sub TxtAssesetCode1_KeyPress(KeyAscii As Integer)
Dim AsseID As Double
If TxtAssesetCode1.Text <> "" Then
GetAsseteCode_ID AsseID, TxtAssesetCode1.Text, 1
DcbSandAdd.BoundText = AsseID
End If
End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
         
        CboDes.CloseUp
    End If

End Sub

Private Sub TxtFASalesPrice_Change()
 
    'DcFixedAssets_Click (0)

    LoseProfitValue = val(TxtFASalesPrice) - val(TxtCurrentValue)
    TxtLoseProfitValue.Text = Abs(LoseProfitValue)

    If LoseProfitValue > 0 Then
        TxtLoseProfitValue.ForeColor = vbGreen
    ElseIf LoseProfitValue < 0 Then
        TxtLoseProfitValue.ForeColor = vbRed
    Else
        TxtLoseProfitValue.ForeColor = vbBlack
    End If

   ' WriteDev
End Sub



Private Sub TxtFixeCurValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtFixeNewValue.Text = val(TxtFixeCurValue.Text) + val(TxtFixeIncValue.Text)
End If
End Sub

Private Sub TxtFixeIncValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtFixeNewValue.Text = val(TxtFixeCurValue.Text) + val(TxtFixeIncValue.Text)
End If
End Sub

Private Sub TxtFixeIncValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtFixeIncValue.Text, 0)
End Sub

Private Sub TxtFixeNewValue_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtQstNewNo.Text) <> 0 Then
TxtQstNewValue.Text = val(TxtFixeNewValue.Text) / val(TxtQstNewNo.Text)
TxtQstNewValue.Text = Round(val(TxtQstNewValue.Text), 2)
End If
End If
End Sub

Private Sub TxtModFlg_Change()

    'On Error GoTo ErrTrap
    Select Case Me.TxtModFlg.Text

        Case "R"
            DcFixedAssets.Enabled = False
        
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
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
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
        
            DcFixedAssets.Enabled = True
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
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
 
            DcFixedAssets.Enabled = False
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
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtQstCurNo_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtQstNewNo.Text = val(TxtQstCurNo.Text) + val(TxtQstIncNo.Text)
TxtQstNewNo.Text = Round(val(TxtQstNewNo.Text), 2)
End If
End Sub

Private Sub TxtQstCurValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtQstIncValue.Text = val(TxtQstNewValue.Text) - val(TxtQstCurValue.Text)
TxtQstIncValue.Text = Round(val(TxtQstIncValue.Text), 2)
End If
End Sub

Private Sub TxtQstIncNo_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtQstNewNo.Text = val(TxtQstCurNo.Text) + val(TxtQstIncNo.Text)
TxtQstNewNo.Text = Round(val(TxtQstNewNo.Text), 2)
End If
End Sub

Private Sub TxtQstIncNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQstIncNo.Text, 0)
End Sub



Private Sub TxtQstNewNo_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtQstNewNo.Text) <> 0 Then
TxtQstNewValue.Text = val(TxtFixeNewValue.Text) / val(TxtQstNewNo.Text)
TxtQstNewValue.Text = Round(val(TxtQstNewValue.Text), 2)
End If
End If
End Sub

Private Sub TxtQstNewValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtQstIncValue.Text = val(TxtQstNewValue.Text) - val(TxtQstCurValue.Text)
TxtQstIncValue.Text = Round(val(TxtQstIncValue.Text), 2)
End If
End Sub

Private Sub TxtQstNo_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtQstIncNo.Text = val(TxtQstNo.Text)
End If
End Sub

Private Sub TxtQstNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQstNo.Text, 0)
End Sub

Private Sub TxtQstValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQstValue.Text, 0)
End Sub
Sub RetriveOrderPayment(Optional NoteSerial1 As Double = 0, Optional NoteType As Integer = -1)
If NoteSerial1 <> 0 Then
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = " SELECT     NoteType, Note_Value, NoteSerial1, NoteID"
sql = sql & " From dbo.Notes"
sql = sql & " Where (NoteType = " & NoteType & ") And (NoteSerial1 = " & NoteSerial1 & ")AND (AssestPayd IS NULL) "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
TxtOderNo.Text = IIf(IsNull(Rs5("NoteID").value), 0, Rs5("NoteID").value)
'TxtVouSerial.text = IIf(IsNull(Rs5("NoteSerial1").value), "", Rs5("NoteSerial1").value)
TxtAddValue.Text = IIf(IsNull(Rs5("Note_Value").value), 0, Rs5("Note_Value").value)
Else
TxtOderNo.Text = 0
'TxtVouSerial.text = ""
TxtAddValue.Text = 0
End If
End If
End Sub
Function GetAccountDeptBillAcounting(Optional NoteSerial1 As Double) As String
Dim i As Integer
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
sql = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
sql = sql & " FROM         dbo.notes_all LEFT OUTER JOIN"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.A_NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
sql = sql & " WHERE     (dbo.notes_all.NoteType = 80) AND (dbo.notes_all.bill_type <> 2) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
sql = sql & "                      (dbo.notes_all.NoteSerial1 = " & NoteSerial1 & ")"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Rs9.MoveFirst
For i = 1 To Rs9.RecordCount
Rs9.MoveNext
Next i

End If
End Function
Sub RetriveOrder(Optional NoteSerial1 As Double = 0, Optional NoteType As Integer = -1)
If NoteSerial1 <> 0 Then
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = " SELECT     NoteType, Note_Value, NoteSerial1, NoteID"
sql = sql & " From dbo.notes_all"
sql = sql & " Where (NoteType = " & NoteType & ") And (NoteSerial1 = " & NoteSerial1 & ")AND (AssestPayd IS NULL) "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
TxtOderNo.Text = IIf(IsNull(Rs5("NoteID").value), 0, Rs5("NoteID").value)
'TxtVouSerial.text = IIf(IsNull(Rs5("NoteSerial1").value), "", Rs5("NoteSerial1").value)
TxtAddValue.Text = IIf(IsNull(Rs5("Note_Value").value), 0, Rs5("Note_Value").value)
Else
TxtOderNo.Text = 0
'TxtVouSerial.text = ""
TxtAddValue.Text = 0
End If
End If
End Sub

Private Sub TxtVouSerial_Change()
If Me.TxtModFlg.Text <> "R" Then
Dim NoteType As Integer
If val(TxtVouSerial.Text) <> 0 Then
If val(DcbBasedOn.ListIndex) = 0 Then
NoteType = 80
RetriveOrder val(TxtVouSerial.Text), NoteType
ElseIf val(DcbBasedOn.ListIndex) = 2 Then
NoteType = 3
RetriveOrder val(TxtVouSerial.Text), NoteType
End If
If val(DcbBasedOn.ListIndex) = 1 Then
NoteType = 5
RetriveOrderPayment val(TxtVouSerial.Text), NoteType
End If
End If
End If
End Sub

Private Sub TxtVouSerial_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
If val(DcbBasedOn.ListIndex) = 0 Then
   Load FrmNotesSearch
   FrmNotesSearch.SearchType = 8008
   FrmNotesSearch.show vbModal
 ElseIf val(DcbBasedOn.ListIndex) = 1 Then
Load FrmNotesSearch
   FrmNotesSearch.SearchType = 5005
   FrmNotesSearch.show vbModal
   
ElseIf val(DcbBasedOn.ListIndex) = 2 Then
Load FrmNotesSearch
   FrmNotesSearch.SearchType = 3003
   FrmNotesSearch.show vbModal
   End If
End If
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With VSFlexGrid1

        Select Case .ColKey(Col)
    
            Case "Value"
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
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
                    If BolEditOnMainAccounts = False Then
                        'If LastAccount(rs("Account_Code").value) = False Then
                        '    .TextMatrix(Row, Col) = ""
                        '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                        '    Exit Sub
                        'End If
                    End If

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

                If BolEditOnMainAccounts = False Then
                    'If LastAccount(StrAccountCode) = False Then
                    '    .TextMatrix(Row, Col) = ""
                    '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Else

                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    'End If
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

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
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

         

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

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
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "AccountName"
                
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
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
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

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid2

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                Dim GroupID As Integer
                Dim branch_id As Integer
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
            
                StrSQL = "select * from FixedAssets where id=" & val(StrAccountCode)
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    GroupID = IIf(IsNull(rs("group_id").value), "", rs("group_id").value)
                    .TextMatrix(Row, .ColIndex("groupid")) = GroupID
                    branch_id = IIf(IsNull(rs("Branch_NO").value), "", rs("Branch_NO").value)
                    .TextMatrix(Row, .ColIndex("branch_id")) = branch_id
              
                Else
                    .TextMatrix(Row, .ColIndex("groupid")) = 0
                    GroupID = 0
                    branch_id = 0
                    .TextMatrix(Row, .ColIndex("branch_id")) = 0
                End If
              
                .TextMatrix(Row, .ColIndex("AccountCode")) = get_FixedAsset_Account(GroupID, branch_id)
               
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

     

End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2
 
        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
    
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
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
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from FixedAssets where New_or_opening=0 and PurchasePrice=0 order by Name"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Name", "Id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
         
        End Select

    End With

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

Public Sub Retrive(Optional Lngid As String = "")
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid = "-1" Then
        Cmd_Click (0)
    End If

    If Lngid <> "" Then
        '  If XPTxtID.text <> 0 Then
        rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "›« Ê—… €Ì— „”Ã·… ", vbInformation
            Else
                MsgBox " Un Refistered Bill ", vbInformation
            End If

            Exit Sub
        End If

        '  End If
    End If

   ' If Not IsNull(rs("general_cost_center").value) Then
    '    Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
   '' Else
    '    Me.DcCostCenter.BoundText = ""
    'End If
   XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
   
   'TxtSerial1.Text = val(XPTxtID.Text)
    Me.TxtSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
   XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
   Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), 0, rs("UserID").value)
   Me.DcBranch.BoundText = IIf(IsNull(rs("BranchID").value), 0, rs("BranchID").value)
   Me.CboType.ListIndex = IIf(IsNull(rs("TypeSand").value), -1, rs("TypeSand").value)
   DcbSandAdd.BoundText = IIf(IsNull(rs("SandAdd").value), 0, rs("SandAdd").value)
   DcFixedAssets.BoundText = IIf(IsNull(rs("FixedID").value), 0, rs("FixedID").value)
   TxtPurchasePrice.Text = IIf(IsNull(rs("PurchasePrice").value), 0, val(rs("PurchasePrice").value))
   TxtPurchasePrice2.Text = IIf(IsNull(rs("PurchasePrice2").value), 0, val(rs("PurchasePrice2").value))
   TxtAccDepre.Text = IIf(IsNull(rs("AccDepre").value), 0, val(rs("AccDepre").value))
   TxtAccDepre2.Text = IIf(IsNull(rs("AccDepre2").value), 0, val(rs("AccDepre2").value))
   TxtCurrentValue.Text = IIf(IsNull(rs("CurrentValue").value), 0, val(rs("CurrentValue").value))
   TxtCurrentValue2.Text = IIf(IsNull(rs("CurrentValue2").value), 0, val(rs("CurrentValue2").value))
   TxtNuminstallmTotal.Text = IIf(IsNull(rs("NuminstallmTotal").value), 0, val(rs("NuminstallmTotal").value))
   TxtNuminstallmTotal2.Text = IIf(IsNull(rs("NuminstallmTotal2").value), 0, val(rs("NuminstallmTotal2").value))
   TxtNuminstallmExcu.Text = IIf(IsNull(rs("NuminstallmExcu").value), 0, val(rs("NuminstallmExcu").value))
   TxtNuminstallmExcu2.Text = IIf(IsNull(rs("NuminstallmExcu2").value), 0, val(rs("NuminstallmExcu2").value))
   TxtNuminstallmCurr.Text = IIf(IsNull(rs("NuminstallmCurr").value), 0, val(rs("NuminstallmCurr").value))
   TxtNuminstallmCurr2.Text = IIf(IsNull(rs("NuminstallmCurr2").value), 0, val(rs("NuminstallmCurr2").value))
   txt_general_des.Text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
   TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
   '''''//////////
   DateAdd.value = IIf(IsNull(rs("DateAdd").value), Date, rs("DateAdd").value)
   SatrtDate.value = IIf(IsNull(rs("SatrtDate").value), Date, rs("SatrtDate").value)
   TxtAddValue.Text = IIf(IsNull(rs("AddValue").value), 0, rs("AddValue").value)
   TxtQstValue.Text = IIf(IsNull(rs("QstValue").value), 0, rs("QstValue").value)
   TxtQstNo.Text = IIf(IsNull(rs("QstNo").value), 0, rs("QstNo").value)
If Not (IsNull(rs("Distrbute").value)) Then
If val(rs("Distrbute").value) = 0 Then
Distrbute(0).value = True
ElseIf val(rs("Distrbute").value) = 1 Then
Distrbute(1).value = True
End If
End If
TxtQstCurNo.Text = IIf(IsNull(rs("QstCurNo").value), 0, rs("QstCurNo").value)
TxtQstIncNo.Text = IIf(IsNull(rs("QstIncNo").value), 0, rs("QstIncNo").value)
TxtQstNewNo.Text = IIf(IsNull(rs("QstNewNo").value), 0, rs("QstNewNo").value)
TxtQstCurValue.Text = IIf(IsNull(rs("QstCurValue").value), 0, rs("QstCurValue").value)
TxtQstIncValue.Text = IIf(IsNull(rs("QstIncValue").value), 0, rs("QstIncValue").value)
TxtQstNewValue.Text = IIf(IsNull(rs("QstNewValue").value), 0, rs("QstNewValue").value)
TxtFixeCurValue.Text = IIf(IsNull(rs("FixeCurValue").value), 0, rs("FixeCurValue").value)
TxtFixeIncValue.Text = IIf(IsNull(rs("FixeIncValue").value), 0, rs("FixeIncValue").value)
TxtFixeNewValue.Text = IIf(IsNull(rs("FixeNewValue").value), 0, rs("FixeNewValue").value)
Me.DcbBasedOn.ListIndex = IIf(IsNull(rs("BasedOn").value), -1, rs("BasedOn").value)
TxtOderNo.Text = IIf(IsNull(rs("OderNo").value), 0, rs("OderNo").value)
TxtVouSerial.Text = IIf(IsNull(rs("VouSerial").value), "", rs("VouSerial").value)

    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)

Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
Me.DcbAccount.BoundText = IIf(IsNull(rs("Account").value), "", rs("Account").value)



'    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)
'
'    If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 
'
'        StrSQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value],dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
'        StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
''        StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
 '       StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
 '       StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
'
'        Set RsDev = New ADODB.Recordset
'        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'        If RsDev.RecordCount > 0 Then
'            RsDev.MoveFirst
'        End If
'
'        With Me.VSFlexGrid1
'
'            .Rows = .FixedRows + RsDev.RecordCount
'
'            For i = .FixedRows To .Rows
'                .TextMatrix(i, .ColIndex("LineNo")) = i
'                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
'
'                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'
'                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
'                Else
'                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
'                End If
'
'                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
'
'                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
'
'                RsDev.MoveNext
'            Next i
'
'        End With
'
'        Exit Sub
'    End If
'
'    '-----------------------------------------------------------------------------
'    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·«’Ê·
'        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
'        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
'        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"
'
'        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
'        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
'        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
'        StrSQL = "SELECT  dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetbranch_id , dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetgroupid, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetID ,  dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name,"
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
'        StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
'        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
'        StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
'        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
'        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
'
'        Set RsDev = New ADODB.Recordset
'        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'        If Not (RsDev.BOF Or rs.EOF) Then
'            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
'            Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
'            RsDev.MoveFirst
'
'            For i = 1 To RsDev.RecordCount
'
'                If RsDev("Credit_Or_Debit").value = 0 Then
'                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
'                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
'                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
'                End If
'
'                RsDev.MoveNext
'            Next i
'
'            RsDev.MoveFirst
'
'            With Me.VSFlexGrid2
'
'                If Me.dcproject.BoundText = "" Then
 '                   .Rows = .FixedRows + RsDev.RecordCount
''                Else
 '                   .Rows = .FixedRows + RsDev.RecordCount - 1
 '               End If
'
'                For i = .FixedRows To .Rows - 1
'                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
'
'                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
'
'                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("FixedAssetId").value), "", RsDev("FixedAssetId").value)
'
'                    .TextMatrix(i, .ColIndex("AccountName")) = getFixedAsstName(val(.TextMatrix(i, .ColIndex("id"))), "name")
'
'                    .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(RsDev("FixedAssetgroupid").value), "", RsDev("FixedAssetgroupid").value)
'
'                    .TextMatrix(i, .ColIndex("branch_id")) = IIf(IsNull(RsDev("FixedAssetbranch_id").value), "", RsDev("FixedAssetbranch_id").value)
'
'                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
'
'                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
'
'                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
'
'                    RsDev.MoveNext
'                Next i
'
'                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
'                ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
'                '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
'                '  .Rows - 1, .ColIndex("CreditValue"))
'                '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
'                '  .Rows - 1, .ColIndex("DebitValue"))
'            End With
'
'        End If
'
'    End If
'
'    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
'    ReLineGrid
    Me.TxtModFlg = "R"
'
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim sql As String
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then

        If Trim(Me.DcFixedAssets.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·«’·..!!"
            Else
                Msg = "Select Asset..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcFixedAssets.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        If val(DcbBasedOn.ListIndex) = 3 Then
    If Me.DcbAccount.BoundText = "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
       Else
       MsgBox "Please Select Account"
       End If
       DcbAccount.SetFocus
       Exit Sub
       End If
     End If
       ' If CheckFixedAssetsDipre(val(DcFixedAssets.BoundText)) = True And Me.TxtModFlg = "N" Then
       '     If SystemOptions.UserInterface = ArabicInterface Then
       '         Msg = "     „ «· Œ·’ „‰ Â–« «·«’· ”«»ﬁ«..!!"
       ''     Else
        '        Msg = "  Asset was disposed..!!"
        '    End If
'
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            '                                DcFixedAssets.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
'
'        End If

      '  If Me.CboPaymentType1.ListIndex = -1 Then
      '      If SystemOptions.UserInterface = ArabicInterface Then
      '          Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·›« Ê—… ...!!!"
      '      Else
      '          Msg = "Select Bill Type ...!!!"
      '      End If
'
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            CboPaymentType.SetFocus
'            Exit Sub
'        End If
    
        If Me.CboType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·”‰œ ...!!!"
            Else
                Msg = "Select   Type ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboType.SetFocus
            Exit Sub
        End If
    
'        If Me.CboType.ListIndex = 0 Then '»Ì⁄ «’·
'
'            If val(TxtFASalesPrice.text) = 0 Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ»   «œŒ«· ﬁÌ„… «·»Ì⁄ ...!!!"
'                Else
'                    Msg = "    Enter Price ...!!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                CboType.SetFocus
'                Exit Sub
'            End If
'
'        ElseIf Me.CboType.ListIndex = 1 Then ' Œ—Ìœ «’·
'            TxtFASalesPrice.text = 0
'        End If
'
   '     If Me.CboPaymentType.ListIndex = -1 And Me.CboType.ListIndex = 0 Then
   '         If SystemOptions.UserInterface = ArabicInterface Then
   '             Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
   ''         Else
    ''            Msg = "Select Payment method ...!!!"
     '       End If
'
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            CboPaymentType.SetFocus
'            Exit Sub
'        End If
    
'        If Me.CboPaymentType.ListIndex = 2 Then
'            If Trim(Me.DCVendor.BoundText) = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
'                Else
'                    Msg = "Select vendor..!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DCVendor.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'        End If
        
'        If Me.CboPaymentType.ListIndex = 4 Then
'            If Trim(Me.DCAccounts1.BoundText) = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·Õ”«»..!!"
'                Else
'                    Msg = "Select Account..!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DCAccounts1.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'        End If
    
'        If Me.CboPaymentType.ListIndex = 0 Then
'            If Trim(Me.DcboBox.BoundText) = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
'                Else
'                    Msg = "Select Box..!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DcboBox.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'        ElseIf Me.CboPaymentType.ListIndex = 1 Then
'            '                                                             If Me.DcboBankName.BoundText = "" Then
            '                                                                         If SystemOptions.UserInterface = ArabicInterface Then
            '                                                                             Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            '                                                                         Else
            '                                                                         Msg = "Select Bank...!!"
            '
            '                                                                        End If
            '                                                                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            ' '                                                                       DcboBankName.SetFocus
            ' '                                                                       SendKeys "{F4}"
            '                                                                        Exit Sub
            '                                                            End If
            '                If Trim$(Me.TxtChequeNumber.text) = "" Then
            '                                      If SystemOptions.UserInterface = ArabicInterface Then
            '                                          Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            '                                      Else
            '                                      Msg = "Enter Cheque No:...!!"
            '                                      End If
            '                  MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '                  TxtChequeNumber.SetFocus
            '                  Exit Sub
            '              End If
            '
      
'            If SystemOptions.ChequeBox = True Then
'
'                If DcChequeBox.BoundText = "" Then
'
'                    If SystemOptions.UserInterface = ArabicInterface Then
''                        Msg = "Õœœ Õ«›Ÿ… «·‘Ìﬂ«  ...!!"
 '                   Else
 '                       Msg = "Select Cheque Box ...!!"
 '                   End If
'
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    DcChequeBox.SetFocus
'                    SendKeys "{F4}"
'                    Exit Sub
'
'                End If
'
''                If TXTBankName.text = "" Then
 '
 '                   If SystemOptions.UserInterface = ArabicInterface Then
 '                       Msg = "«ﬂ » «”„ »‰ﬂ «·‘Ìﬂ    « ...!!"
 '                   Else
 '                       Msg = " Enter Bank Name For Cheque  ...!!"
 '                   End If
'
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    TXTBankName.SetFocus
'                    SendKeys "{F4}"
'                    Exit Sub
'
'                End If
'
'                If Trim$(Me.TxtChequeNumber.text) = "" Then
'                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    TxtChequeNumber.SetFocus
'                    Exit Sub
'                End If
'
'            Else
'
'                If Me.DcboBankName.BoundText = "" Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    DcboBankName.SetFocus
'                    SendKeys "{F4}"
'                    Exit Sub
'                End If
'
'                If Trim$(Me.TxtChequeNumber.text) = "" Then
'                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
'                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    TxtChequeNumber.SetFocus
'                    Exit Sub
'                End If
'            End If
'
'        ElseIf Me.CboPaymentType.ListIndex = 3 Then
'
'            If Me.DcboBankName.BoundText = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
'                Else
'                    Msg = "Select Bank...!!"
'
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DcboBankName.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'            If Trim$(Me.TxtChequeNumber.text) = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
'                Else
'                    Msg = "Enter Transfer No:...!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                TxtChequeNumber.SetFocus
'                Exit Sub
'
'            End If
'
'        ElseIf Me.CboPaymentType.ListIndex = 5 Then
'
'            If Me.DcboBankName.BoundText = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
'                Else
'                    Msg = "Select Bank...!!"
'
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DcboBankName.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'            If Trim$(Me.TxtChequeNumber.text) = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
'                Else
'                    Msg = "Enter Cheque No:...!!"
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                TxtChequeNumber.SetFocus
'                Exit Sub
'
'            End If
'
'        End If
'
'        If Me.TxtModFlg.text = "N" Then
'            If Me.CboPaymentType.ListIndex = 0 Then
'                If val(Me.DcboBox.BoundText) <> 0 Then
'                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
'                        Exit Sub
'                    End If
'                End If
'            End If

'        ElseIf Me.TxtModFlg.text = "E" Then

'            If Me.CboPaymentType.ListIndex = 0 Then
'                If val(Me.DcboBox.BoundText) <> 0 Then
'                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value, , , val(Me.XPTxtID.text)) = False Then
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'
'        Dim xrow As Integer
'
'        Dim i As Integer
'
'        ' calcnets
'
'        '-------------------------------------------------------------------------------------------
'        Dim notes_result As String
'        Dim Vchr_result As String
'
'        '-------------------------------------------------------------------------------------------
'        If TxtSerial1.text = "" Then
'            Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 28, 8028)
'
'            If Vchr_result = "error" Then
''                If SystemOptions.UserInterface = ArabicInterface Then
 '                   MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   Œ·’ „‰ «’· ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
 '               Else
 '                   MsgBox " Cant't Create  Disposal Of FA  Voucher to this Process no You exceed the maximum number ": Exit Sub
 '               End If
'
'            Else
'
'                If Vchr_result = "" Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'                    Else
'                        MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
'                    End If
'
'                Else
'                    TxtSerial1.text = Vchr_result
'                End If
'            End If
'        End If
'
'        If TxtSerial.text = "" Then
'            notes_result = Notes_coding(val(my_branch), XPDtbTrans.value)
'
'            If notes_result = "error" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
'                Else
'                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
'                End If
'
'            Else
'
'                If notes_result = "" Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
'                    Else
'                        MsgBox "You must Define JE Coding ": Exit Sub
'                    End If
'
'                Else
'                    TxtSerial.text = notes_result
'                End If
'            End If
'        End If
'


      If TxtSerial1.Text = "" Then
                If Voucher_coding(val(DcBranch.BoundText), XPDtbTrans.value, 78, 8073) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(DcBranch.BoundText), XPDtbTrans.value, 78, 8073) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtSerial1.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtSerial1.Text = Voucher_coding(val(DcBranch.BoundText), XPDtbTrans.value, 78, 8073)
                    End If
                End If
            End If
        Cn.BeginTrans
        BeginTrans = True
'
        '///////////////NOTESALL
        Dim A_NoteID As Long

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
            rs.AddNew
      
       ElseIf Me.TxtModFlg.Text = "E" Then
           StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


  '          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.text)
  '          Cn.Execute StrSQL, , adExecuteNoRecords
  '
  '          StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
  '          Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
    
       
        'TxtSerial1.Text = val(XPTxtID.Text)
        rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.Text)
        rs("UserID").value = val(Me.DCboUserName.BoundText)
        rs("BranchID").value = val(Me.DcBranch.BoundText)
        rs("TypeSand").value = (val(CboType.ListIndex))
        rs("SandAdd").value = val(Me.DcbSandAdd.BoundText)
        rs("FixedID").value = val(Me.DcFixedAssets.BoundText)
        rs("PurchasePrice").value = val(TxtPurchasePrice.Text)
        rs("PurchasePrice2").value = val(TxtPurchasePrice2.Text)
        rs("AccDepre").value = val(TxtAccDepre.Text)
        rs("AccDepre2").value = val(TxtAccDepre2.Text)
        rs("CurrentValue").value = val(TxtCurrentValue.Text)
        rs("CurrentValue2").value = val(TxtCurrentValue2.Text)
        rs("NuminstallmTotal").value = val(TxtNuminstallmTotal.Text)
        rs("NuminstallmTotal2").value = val(TxtNuminstallmTotal2.Text)
        rs("NuminstallmExcu").value = val(TxtNuminstallmExcu.Text)
        rs("NuminstallmExcu2").value = val(TxtNuminstallmExcu2.Text)
        rs("NuminstallmRemin").value = val(TxtNuminstallmRemin.Text)
        rs("NuminstallmRemin2").value = val(TxtNuminstallmRemin2.Text)
        rs("NuminstallmCurr").value = val(TxtNuminstallmCurr.Text)
        rs("NuminstallmCurr2").value = val(TxtNuminstallmCurr2.Text)
        rs("general_des").value = txt_general_des.Text
        rs("NoteSerial").value = TxtSerial.Text
        ''''/////////
        rs("Effct").value = 1
        rs("DateAdd").value = DateAdd.value
        rs("SatrtDate").value = SatrtDate.value
        rs("AddValue").value = val(TxtAddValue.Text)
        rs("QstValue").value = val(Me.TxtQstValue.Text)
        rs("QstNo").value = val(TxtQstNo.Text)
        If Distrbute(0).value = True Then
        rs("Distrbute").value = 0
        ElseIf Distrbute(1).value = True Then
        rs("Distrbute").value = 1
        Else
        rs("Distrbute").value = Null
        End If
        rs("QstCurNo").value = val(TxtQstCurNo.Text)
        rs("QstIncNo").value = val(TxtQstIncNo.Text)
        rs("QstNewNo").value = val(TxtQstNewNo.Text)
        rs("QstCurValue").value = val(TxtQstCurValue.Text)
        rs("QstIncValue").value = val(TxtQstIncValue.Text)
        rs("QstNewValue").value = val(TxtQstNewValue.Text)
        rs("FixeCurValue").value = val(TxtFixeCurValue.Text)
        rs("FixeIncValue").value = val(TxtFixeIncValue.Text)
        rs("FixeNewValue").value = val(TxtFixeNewValue.Text)
        rs("OderNo").value = val(TxtOderNo.Text)
        rs("BasedOn").value = val(DcbBasedOn.ListIndex)
        rs("VouSerial").value = (TxtVouSerial.Text)
        rs("Account").value = Me.DcbAccount.BoundText
        If val(DcbBasedOn.ListIndex) = 0 Or val(DcbBasedOn.ListIndex) = 2 Then
        sql = "Update notes_all set AssestPayd=1 where NoteID=" & val(TxtOderNo.Text) & ""
        Cn.Execute sql
        ElseIf val(DcbBasedOn.ListIndex) = 1 Then
        sql = "Update Notes set AssestPayd=1 where NoteID=" & val(TxtOderNo.Text) & ""
        Cn.Execute sql
        End If
rs.update
End If



'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
'            Else
'                bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
'
'            End If
'
'        ElseIf Me.CboPaymentType.ListIndex = 2 Then
'            rs("NoteCashingType").value = 2
'            rs("CusID").value = val(Me.DCVendor.BoundText)
'        ElseIf Me.CboPaymentType.ListIndex = 3 Then
'            rs("BoxID").value = Null
'            rs("BankID").value = val(Me.DcboBankName.BoundText)
'            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
'            rs("DueDate").value = Me.DtpChequeDueDate.value
'            rs("NoteCashingType").value = 3
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                bankDes = "  ’—› »ÕÊ«·…  —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
'            Else
'                bankDes = "  Bank Transfere No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
'            End If
'
'        ElseIf Me.CboPaymentType.ListIndex = 4 Then
'            rs("BoxID").value = Null
'            rs("BankID").value = Null
'            rs("ChqueNum").value = Null
'            rs("DueDate").value = Null
'            rs("NoteCashingType").value = 4
'
'            rs("AccountCode").value = (Me.DCAccounts1.BoundText)
'
'        ElseIf Me.CboPaymentType.ListIndex = 5 Then
'            rs("BoxID").value = Null
'            rs("BankID").value = val(Me.DcboBankName.BoundText)
'            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
'            rs("DueDate").value = Me.DtpChequeDueDate.value
'            rs("NoteCashingType").value = 5
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                bankDes = "  Õ’·  »‘Ìﬂ   —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
'            Else
'                bankDes = "  Cheque   No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
'            End If
'
'            '
'        End If
'
'        If CboType.ListIndex = 1 Then
'            rs("BoxID").value = Null
'            rs("BankID").value = Null
'            rs("ChqueNum").value = Null
'            rs("DueDate").value = Null
'            rs("NoteCashingType").value = -1
'
'        End If
'
'        rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
'        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
'        rs("Buy").value = "0"
'        rs("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
'        rs("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
'        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”·   ›« Ê—…
'        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'        rs("numbering_type1").value = sand_numbering_type(28) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
'
'        rs("sanad_year").value = year(XPDtbTrans.value)
'        rs("sanad_month").value = Month(XPDtbTrans.value)
'
'        If dcproject.BoundText <> "" Then
'            ' rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
'        Else
'            ' rs("note_value_by_characters").value = WriteNo(Format(Val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
'        End If
'
'        If Me.TxtModFlg.text = "N" Then
'            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
'            TXT_A_NoteID.text = A_NoteID
'        Else
'            A_NoteID = val(TXT_A_NoteID.text)
'        End If
'
'        rs("A_NoteID").value = val(A_NoteID)
'
'        rs.update
'
'        Dim ExpensesID As Double
'
'        Dim NoteID As String
'
'        '  «·«’Ê· „œÌ‰
'
'        '//////////////////////////////////////Notes////////////////////////////////////
'        Set RsNotes = New ADODB.Recordset
'        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'        Set RsDev = New ADODB.Recordset
'        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'        '«·ÿ—› «·„œÌ‰
'
'        line_no = 1
'
'        ' «·«’Ê· «·ÿ—› «·„œÌÊ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
'
'        RsNotes.AddNew
'        NoteID = CStr(new_id("Notes", "NoteID", "", True))
'        RsNotes("NoteID").value = CStr(NoteID)
'        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
'
'        '    RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
'        RsNotes("Remark").value = Me.txt_general_des
'        RsNotes("foxy_no").value = val(Text1.text)
'
'        If Me.CboPaymentType.ListIndex = 0 Then
'            RsNotes("BoxID").value = val(DcboBox.BoundText)
'            RsNotes("BankID").value = Null
'            RsNotes("ChqueNum").value = Null
'            RsNotes("DueDate").value = Null
'            RsNotes("NoteCashingType").value = 0
'        ElseIf Me.CboPaymentType.ListIndex = 1 Then
'            RsNotes("BoxID").value = Null
''
 '           ' RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
 '           If SystemOptions.ChequeBox = False Then
 '
 '               rs("BankID").value = val(Me.DcboBankName.BoundText)
 '           Else
 '               rs("BankID").value = Null
 '           End If
 '
 '           RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
 '           RsNotes("DueDate").value = Me.DtpChequeDueDate.value
 '           RsNotes("NoteCashingType").value = 1
 '       ElseIf Me.CboPaymentType.ListIndex = 2 Then
 '           RsNotes("CusID").value = val(DCVendor.BoundText)
 '
 '       ElseIf Me.CboPaymentType.ListIndex = 3 Then
 '           RsNotes("BoxID").value = Null
 '           RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
 '           RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
 '           RsNotes("DueDate").value = Me.DtpChequeDueDate.value
 '           RsNotes("NoteCashingType").value = 3
 '
 '       ElseIf Me.CboPaymentType.ListIndex = 4 Then
 '           RsNotes("BoxID").value = Null
 '           RsNotes("BankID").value = Null
 '           RsNotes("ChqueNum").value = Null
 '           RsNotes("DueDate").value = Null
 '           RsNotes("NoteCashingType").value = Null
 '
 '       ElseIf Me.CboPaymentType.ListIndex = 5 Then
 '           RsNotes("BoxID").value = Null
 '           RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
 '           RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
 '           RsNotes("DueDate").value = Me.DtpChequeDueDate.value
 '           RsNotes("NoteCashingType").value = 5
                            
 '       End If
 '
 '       RsNotes("NoteType").value = 8028
 '       RsNotes("NoteDate").value = XPDtbTrans.value
 '       RsNotes("UserID").value = user_id
 '
 '       RsNotes("notes_all").value = Me.XPTxtID.text
 '       RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
 '       RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
 '       RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
 '       RsNotes("numbering_type1").value = sand_numbering_type(28) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
 '       RsNotes("sanad_year").value = year(XPDtbTrans.value)
 '       RsNotes("sanad_month").value = Month(XPDtbTrans.value)
 '       RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
 '
 '       RsNotes.update
 '
 '       XPTxtVal = 0
 '
 '       txtmyDes = txtmyDes & " " & Me.txt_general_des
 '       txtmyDesE = txtmyDesE & " " & Me.txt_general_des
'
'        '«·ÿ—› «·„œÌÊ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
'        If val(TxtFASalesPrice.text) > 0 Then
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("DEV_ID_Line_No").value = line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = DcboCreditSide.BoundText
'            RsDev("Value").value = IIf(IsNumeric(TxtFASalesPrice.text), TxtFASalesPrice.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
'            RsDev("Credit_Or_Debit").value = 0
'            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
'            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
'            RsDev("RecordDate").value = Me.XPDtbTrans.value
'            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
'            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'            RsDev("UserID").value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'            RsDev("notes_all").value = Me.XPTxtID.text
'
'            XPTxtVal = val(XPTxtVal.text) + val(TxtFASalesPrice.text)
'            RsDev.update
'            line_no = line_no + 1
'        End If
'
'        If val(TxtAccDepre.text) > 0 Then
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("DEV_ID_Line_No").value = line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = get_FixedAsset_Account(group_id, branch_id, "Account_Code2")
'            RsDev("Value").value = IIf(IsNumeric(TxtAccDepre.text), TxtAccDepre.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
'            RsDev("Credit_Or_Debit").value = 0
'            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
'            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
'            RsDev("RecordDate").value = Me.XPDtbTrans.value
'            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
'            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'            RsDev("UserID").value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'            RsDev("notes_all").value = Me.XPTxtID.text
'
'            RsDev.update
'            XPTxtVal = val(XPTxtVal.text) + val(TxtAccDepre.text)
'            line_no = line_no + 1
'        End If
'
'        If val(TxtFASalesPrice) > val(TxtCurrentValue.text) Then
'            ProfitOrLose = 1
'            ProfitOrLoseValue = val(TxtFASalesPrice) - val(TxtCurrentValue.text)
'        ElseIf val(TxtFASalesPrice) < val(TxtCurrentValue.text) Then
'            ProfitOrLose = 0
'            ProfitOrLoseValue = Abs(val(TxtCurrentValue.text) - val(TxtFASalesPrice))
'            XPTxtVal = val(XPTxtVal.text) + (val(TxtCurrentValue.text) - val(TxtFASalesPrice))
'        Else
'            ProfitOrLose = -1
'            ProfitOrLoseValue = 0
'        End If
'
'        If val(ProfitOrLoseValue) > 0 Then
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("DEV_ID_Line_No").value = line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = DCAccounts.BoundText
'            RsDev("Value").value = IIf(IsNumeric(ProfitOrLoseValue), ProfitOrLoseValue, 0) '.TextMatrix(I, .ColIndex("VALUE"))
'            RsDev("Credit_Or_Debit").value = ProfitOrLose
'            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
'            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
'            RsDev("RecordDate").value = Me.XPDtbTrans.value
'            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
'            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'            RsDev("UserID").value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'            RsDev("notes_all").value = Me.XPTxtID.text
'
'            RsDev.update
'
'            line_no = line_no + 1
'        End If
'
'        If val(TxtPurchasePrice.text) > 0 Then
'            RsDev.AddNew
'            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'            RsDev("DEV_ID_Line_No").value = line_no
'            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'            RsDev("Account_Code").value = get_FixedAsset_Account(group_id, branch_id, "Account_Code")
'            RsDev("Value").value = IIf(IsNumeric(TxtPurchasePrice.text), TxtPurchasePrice.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
'            RsDev("Credit_Or_Debit").value = 1
'            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
'            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
'            RsDev("RecordDate").value = Me.XPDtbTrans.value
'            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
'            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'            RsDev("UserID").value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'            RsDev("notes_all").value = Me.XPTxtID.text
'
'            RsDev.update
'            line_no = line_no + 1
'        End If
'
'    End If
'
'    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'    '   UpdateFixedAssetPurchaseInformations ' ÕœÌÀ »Ì«‰«  «·«’· «
'
'    LblDevID.Caption = LngDevID
'    lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
'
'll:
    Cn.CommitTrans
    BeginTrans = False
    createVoucher
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
'    CuurentLogdata
'
    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
        
            End If

           ' Fg_Journal.Enabled = False

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            End If

           ' Fg_Journal.Enabled = False
    End Select

    'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
    'saveChequeBoxContents (val(Me.XPTxtID.text))
      
    TxtModFlg.Text = "R"
    'Dim sql As String
    'sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text)
    'Cn.Execute sql
    'sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text)
    'Cn.Execute sql
  '
  '  sql = "Update   FixedAssets  set Status_id='" & CboType.ListIndex + 2 & "' where id=" & val(DcFixedAssets.BoundText)
  '  Cn.Execute sql
  ''  sql = "  update FixedAssets  set   KhordaPrice =0 ,  saleprice=" & val(TxtFASalesPrice.text) & " where id=" & val(DcFixedAssets.BoundText)
   ' Cn.Execute sql

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

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

'Function UpdateFixedAssetPurchaseInformations(Optional delete As Boolean)
'    Dim sql As String
'    Dim i As Integer
'    Dim KhordaPrice As Double
'    Dim currentvalue As Double
'    Dim PurcahsePrice As Double
'    Dim Installmentvalue As Double
'
'    With Me.VSFlexGrid2
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
'
'                sql = "update FixedAssets set PurchaseDate=CONVERT(DATETIME, '" & XPDtbTrans.value & " 00:00:00', 103), PurchaseBillId=" & TxtSerial1.text & ",PurchasePrice="
'
'                PurcahsePrice = val(.TextMatrix(i, .ColIndex("value")))
'                sql = sql & PurcahsePrice
'
'                Dim noofinstllments As Double
'
'                GetAllDataAboutFixedAsset val(.TextMatrix(i, .ColIndex("id"))), , , , , , , , , , , , , noofinstllments, , , , , , KhordaPrice
'                currentvalue = PurcahsePrice - KhordaPrice
'                sql = sql & ",CurrentValue= " & currentvalue
'
'                If noofinstllments = 0 Then
'                    noofinstllments = 0
'                Else
'                    Installmentvalue = Round(currentvalue / noofinstllments, 2)
'                End If
'
'                sql = sql & ",Installmentvalue= " & Installmentvalue
'                sql = sql & ",NoteSerial=' " & Me.TxtNoteSerial.text & "'"
'                sql = sql & "  where id=" & val(.TextMatrix(i, .ColIndex("id")))
'                Cn.Execute sql
'
'                If noofinstllments <> 0 Then
'                    updateFixedAsseTInstallmentInformations val(.TextMatrix(i, .ColIndex("id"))), , , , XPDtbTrans.value, , , , True, True ' ÕœÌÀ »Ì«‰«  «·«ﬁ”«ÿ
'                End If
'
'                If delete = True Then
'                    '  sql = "update FixedAssets NoteSerial=0,  PurchaseBillId=" & "" & ",PurchasePrice=0,Installmentvalue=0,CurrentValue=0"
'                End If
'
'            End If
'
'        Next i
'
'    End With
'
'End Function

'Public Function save_General_cost_center(cost_center_id As String, _
'                                         cost_center, _
'                                         opr_type As String, _
'                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
'    Dim i As Integer
'    Dim rs As New ADODB.Recordset
'    Dim StrSQL As String
'
'    StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
'    Cn.Execute StrSQL, , adExecuteNoRecords
'
'    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    With Fg_Journal
'
'        .Rows = .Rows + 1
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
'
'                rs.AddNew
''                rs("cost_center_id").value = cost_center_id
 '               rs("cost_center").value = cost_center
 '               rs("value").value = .TextMatrix(i, .ColIndex("value"))
 '               rs("depit_or_credit").value = "„œÌ‰"
 '               rs("opr_id").value = Me.Text1.text
 '               rs("kedno").value = Me.Text1.text
 '               rs("opr_type").value = opr_type
 '               rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
 '               rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
 '               rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
 '               rs("record_date").value = record_date
 '               rs.update
 '
 '           End If
'
'        Next i
'
'    End With
'
'    rs.Close
'End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim sql As String
    'On Error GoTo ErrTrap

   ' If SystemOptions.banks_Accounts3 = True Then
   '     If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
   '         Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
   '         Msg = Msg & Chr(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
   '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
   '         Exit Sub
   '     End If
   ' End If
    
  '  Dim noOfInstallments As Integer 'Â–« «·Ã“¡ Ì √ﬂœ „‰  ‰›Ì– «ﬁ”«ÿ «Â·«ﬂ
  '  Dim msgstr As String
  '  Dim i As Integer

    '    UpdateFixedAssetPurchaseInformations True
    
    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    Else
    Msg = "Confirm Delete?"
   End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
              If Not rs.RecordCount < 1 Then
        If val(DcbBasedOn.ListIndex) = 0 Or val(DcbBasedOn.ListIndex) = 2 Then
        sql = "Update notes_all set AssestPayd=Null where NoteID=" & val(TxtOderNo.Text) & ""
        Cn.Execute sql
        ElseIf val(DcbBasedOn.ListIndex) = 1 Then
        sql = "Update Notes set AssestPayd=Null where NoteID=" & val(TxtOderNo.Text) & ""
        Cn.Execute sql
        End If
   
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

   
   
      
               ' CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
            '        Fg_Journal.Clear flexClearScrollable, flexClearEverything
            '        Fg_Journal.Rows = 3
            '        Fg_Journal.Enabled = False
            '
            '        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            '        VSFlexGrid1.Rows = 2
            '        VSFlexGrid1.Enabled = False
            '
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This is Process unavailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error douring delete " & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



'Private Sub ReLineGrid()
'    Dim i As Integer
'    Dim IntCounter As Integer
'
'    With Fg_Journal
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
'                IntCounter = IntCounter + 1
'                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
'
'            End If
'
'        Next i
'
'    End With
'
'    IntCounter = 0
'
'    With Me.VSFlexGrid1
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
'                IntCounter = IntCounter + 1
'                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
'
'            End If
'
'        Next i
'
'    End With
'
'    IntCounter = 0
'
'    With Me.VSFlexGrid2
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
'                IntCounter = IntCounter + 1
'                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    .TextMatrix(i, .ColIndex("des")) = " ﬁÌ„… ‘—«¡ «·«’· " & .TextMatrix(i, .ColIndex("AccountName"))
'
'                Else
'                    .TextMatrix(i, .ColIndex("des")) = "PURCHASE Value Of Asset " & .TextMatrix(i, .ColIndex("AccountName"))
'                End If
'
'            End If
'
'        Next i
'
'    End With
'
'End Sub

'Function UPDATEStatusToNewAsset()
'    Dim StrSQL As String
'    Dim i As Integer
'
'    With Me.VSFlexGrid2
'
'        For i = .FixedRows To .Rows - 1
'
'            If .TextMatrix(i, .ColIndex("id")) <> "" Then
'                StrSQL = "UPDATE FixedAssets SET CurrentValue = 0,PurchaseBillId='',Installmentvalue = 0,NoteSerial='', New_or_opening=0 ,PurchasePrice=0 where  id=" & val(.TextMatrix(i, .ColIndex("id")))
'
'                Cn.Execute StrSQL
'            End If
'
'        Next i
'
'    End With
'
'End Function

'Private Sub PutData()
'
'    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
'    With Fg_Journal
'
'        If Len(TxtDes.text) > 0 Then
'            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
'            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
'            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
'        Else
'            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
'            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
'            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
'        End If
'
'    End With

'End Sub

'Function sand_numbering() As String
'    On Error Resume Next
'    Dim start_at As Integer
'    Dim end_at As Integer
''    Dim auto_sanad_no As String
 '   Dim NO As String
 '   auto_sanad_no = ""
 '   departement_name = 1
 '   Branch_NO = 1
 ''   connection_string = Cn.ConnectionString
  '  numbering.ConnectionString = connection_string
  '  numbering.CommandType = adCmdText
  '  numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
  '  numbering.Refresh
'
'    If numbering.Recordset.RecordCount = 0 Then
'        numbering_type = 0
'    Else
'        numbering_type = numbering.Recordset.Fields!numbering_id
''        start_at = numbering.Recordset.Fields!start_at
 '       end_at = numbering.Recordset.Fields!end_at
'
'    End If
'
'    If numbering_type = 1 Then
'        detect_no.ConnectionString = connection_string
'        detect_no.CommandType = adCmdText
'        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
'        detect_no.Refresh
'
'        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
'
'            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
'
'            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
'                sand_numbering = "error"
'                Exit Function
'            End If
'        End If
'
'    Else
'
'        If numbering_type = 2 Then
'
'            detect_no.ConnectionString = connection_string
'            detect_no.CommandType = adCmdText
'            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
'            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
'            detect_no.Refresh
'
'            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
'                NO = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
'
'                If end_at = 0 Then end_at = NO + 1
'                If NO >= end_at Then
'                    sand_numbering = "error"
'                    Exit Function
'                End If
'            End If
'
'        Else
'
'            If numbering_type = 3 Then
''
 '               detect_no.ConnectionString = connection_string
 ''               detect_no.CommandType = adCmdText
  ''              detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
   ''             'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
    '            detect_no.Refresh
'
'                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
'                    NO = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
'
'                    If end_at = 0 Then end_at = NO + 1
'                    If NO >= end_at Then
'                        sand_numbering = "error"
'                        Exit Function
'                    End If
'                End If
'
'            End If
'
'        End If
'    End If
'
'    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then
'
'        If numbering_type = 0 Then
'            ' auto_sanad_no = 1
'        Else
'
'            If numbering_type = 1 Then
'                auto_sanad_no = start_at
'            Else
'
'                If numbering_type = 2 Then
'                    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at
'
'                Else
'
'                    If numbering_type = 3 Then
'                        auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at
'
'                    End If
'                End If
'            End If
'        End If
'
'    Else
'
'        If numbering_type = 0 Then
'            'auto_sanad_no = x + 1
'        Else
'
'            If numbering_type = 1 Then
'                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
'            Else
'
'                If numbering_type = 2 Then
'                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
'                    ' no = 1
'                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
'                    '  Else
'                    NO = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
'                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
'                    '  End If
'
'                Else
'
'                    If numbering_type = 3 Then
'                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
'                        'no = 1
'                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
'                        '    Else
'                        NO = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
'                        auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

'                        '    End If
'
'                    End If
'                End If
'            End If
'        End If

'    End If

'    sand_numbering = auto_sanad_no
'
'    'MsgBox auto_sanad_no
'
'End Function
'
'Function setfoxy_Line() As Double
'
'    Dim X As Double
'    X = CStr(new_id("foxy", "id1", "", True))
'    setfoxy_Line = X
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    rs("id1").value = X ' last_line_id
'
'    rs.update
'
'End Function

'Function CuurentLogdata(Optional Currentmode As String)
'     LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & "—ﬁ„ «·”‰œ  " & TxtSerial1.text & Chr(13) & "   «· «—ÌŒ  " & XPDtbTrans & Chr(13) & "   «·›—⁄ " & Dcbranch & Chr(13) & "   ‰Ê⁄ «·”‰œ " & CboType & Chr(13) & "     «·«’·  " & DcFixedAssets & Chr(13) & "   ÿ—Ìﬁ… «·»Ì⁄  " & CboPaymentType & Chr(13) & "   ﬁÌ„…  «·‘—«¡  " & TxtPurchasePrice & Chr(13) & "„Ã„⁄ «·«Â·«ﬂ " & TxtAccDepre & Chr(13) & "      «·ﬁÌ„… «·œ› —Ì…  " & TxtCurrentValue & Chr(13) & "   ﬁÌ„…  «·»Ì⁄  " & "" & Chr(13) & "     «·—»Õ «Ê «·Œ”«—…  " & "" & Chr(13) & "   «·Œ“Ì‰… " & DcboBox & Chr(13) & "   «·»‰ﬂ  " & DcboBankName & Chr(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & Chr(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & Chr(13) & "   «·⁄„Ì·  " & DCVendor & Chr(13) & " «·Õ”«»  " & DCAccounts1 & Chr(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & Chr(13) & "   «Ã„«·Ì «·”‰œ    " & XPTxtValView
'        LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Bill . No " & TxtSerial1.text & Chr(13) & "   Date  " & XPDtbTrans & Chr(13) & "   Branch " & Dcbranch & Chr(13) & "    Type   " & CboType & Chr(13) & "     F.A. Name  " & DcFixedAssets & Chr(13) & "  Salle Type  " & CboPaymentType & Chr(13) & "Purchase Price " & TxtPurchasePrice & Chr(13) & "Acc Depre " & TxtAccDepre & Chr(13) & "Current Value " & TxtCurrentValue & Chr(13) & "  Sales Price " & "" & Chr(13) & "Lose /Profit Value " & "" & Chr(13) & "   Box " & DcboBox & Chr(13) & "   Bank  " & DcboBankName & Chr(13) & "   Cheque No:   " & TxtChequeNumber & Chr(13) & "   Supplier  " & DCVendor & Chr(13) & " Account  " & DCAccounts1 & Chr(13) & "  Remarks  " & txt_general_des & Chr(13) & "   Vchr Total   " & XPTxtValView
'       If Currentmode <> "D" Then
'        AddToLogFile CInt(user_id), 8028, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, , , val(TxtSerial), val(TxtSerial1)
'    Else
'        AddToLogFile CInt(user_id), 8028, Date, Time, LogTextA, LogTextE, Me.name, "D", , , TxtSerial, TxtSerial1
'    End If
    
'End Function

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
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·«÷«›«  ··«’Ê·", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Additions to Assets", 1, 15204351, -2147483630, BolRtl
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

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()
    TxtSerial.Text = ""
    TxtSerial1.Text = ""
    TxtNoteSerial.Text = ""
    lbltoday.Caption = WeekdayName(Weekday(XPDtbTrans.value))
End Sub

Private Sub XPTxtVal_Change()
    XPTxtValView.Text = Format(val(XPTxtVal.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 1)

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

'Private Sub ViewDataList()
'    Dim FrmView As FrmViewList
'    Dim Fg As VSFlex8UCtl.vsFlexGrid
'    Dim StrSQL As String
'    Dim rs As ADODB.Recordset
'    Dim StrComboList As String
'    Dim GrdBack As ClsBackGroundPic
'    'Dim cProgress As ClsProgress
'    Dim BolFrmLoaded As Boolean
'    Set FrmView = New FrmViewList
'    Set Fg = FrmView.vsfGroup1.vsFlexGrid
'
'    With Fg
'        .Cols = 18
'        .RowHeightMin = 320
'        .ExplorerBar = flexExSortShowAndMove
'        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
'        .ColKey(0) = "NoteID"
'        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
'        .ColKey(1) = "NoteSerial"
'        .TextMatrix(0, 2) = "«· «—ÌŒ"
'        .ColKey(2) = "NoteDate"
'        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
'        .ColKey(3) = "Name"
'        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
'        .ColKey(4) = "Note_Value"
'        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
'        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
'        .ColKey(5) = "BoxName"
'        .TextMatrix(0, 6) = "„·«ÕŸ« "
'        .ColKey(6) = "Remark"
'        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
'        .ColKey(7) = "UserName"
'
'        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
'        StrSQL = StrSQL + " Order By NoteID"
'        Set rs = New ADODB.Recordset
'        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
'        '------------------------------------
'        '
'        '
'        '
        '
'
'        '------------------------------------
'        Set .DataSource = rs
'        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
'        .ColKey(0) = "NoteID"
'        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
'        .ColKey(1) = "NoteSerial"
'        .TextMatrix(0, 2) = "«· «—ÌŒ"
'        .ColKey(2) = "NoteDate"
'        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
'        .ColKey(3) = "Name"
'        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
'        .ColKey(4) = "Note_Value"
'        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
'        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
'        .ColKey(5) = "BoxName"
'        .TextMatrix(0, 6) = "„·«ÕŸ« "
'        .ColKey(6) = "Remark"
'        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
'        .ColKey(7) = "UserName"
'
'        'Rs.Close
'        'Set Rs = Nothing
'        .AutoSize 0, .Cols - 1, False
'    End With
'
'    Set GrdBack = New ClsBackGroundPic
'    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
'    FrmView.vsfGroup1.SetRTL = True
'    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
'    FrmView.vsfGroup1.sql = StrSQL
'    FrmView.vsfGroup1.ShowTreeGroups = True
'    FrmView.vsfGroup1.update
'    FrmView.SetDblClickRetrun Me, "NoteID"
'    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
'    FrmView.Show
'End Sub
'
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    'LblValue.Visible = False
     Label1(35).Caption = "No.GL"
   ' Command8.Caption = "Acc.Statement"
    Frame11.Caption = "Accounting"
    Command9.Caption = "Print GL"
    
 Label1(0).Caption = "Branch"
'mdAttach.Caption = "Attachments"
lbl(50).Caption = "Current"
lbl(51).Caption = "Increase"
lbl(52).Caption = "New"
lbl(56).Caption = "Based On"
lbl(53).Caption = "Installment No"
lbl(54).Caption = "Installment Value"
lbl(55).Caption = "Asset Value"
    lbl(24).Caption = "Hint."
    lbl(25).Caption = "This Window Allow Added  Of Fixed Assets"
lbltoday.Caption = "Day"
    lbl(23).Caption = " Type"
    Me.lbl(4).Caption = "Vchr ID"
    Me.lbl(1).Caption = "Date"
    'Label1.Caption = "Branch"
    Frame2.Caption = "Basic Data"
    Frame4.Caption = "New Data"
    lbl(27).Caption = "Select Asset Basic"
    lbl(42).Caption = "Select Asset Add"
    lbl(28).Caption = "Purch. Price"
    lbl(26).Caption = "Purch. Price"
    lbl(29).Caption = "Acc Dep"
    lbl(31).Caption = "Acc Dep"
    lbl(30).Caption = "Current Value"
    lbl(32).Caption = "Current Value"
    lbl(37).Caption = "Total Inst."
    lbl(40).Caption = "Total Inst."
    lbl(35).Caption = "EXE Inst."
    lbl(34).Caption = "EXE Inst."
    lbl(36).Caption = "Remains Inst."
    lbl(39).Caption = "Remains Inst."
    lbl(41).Caption = "Current Value"
    lbl(38).Caption = "Current Value"
    lbl(20).Caption = "Description"
    Me.lbl(2).Caption = "Total"
    Label3.Caption = "GL No."
    CmdAttach.Caption = "Attach"
    Frame9.Caption = "Added Data"
    lbl(46).Caption = "Added Date"
     lbl(47).Caption = "Start Date"
    lbl(45).Caption = "Added Value"
    lbl(49).Caption = "Installment"
    lbl(48).Caption = "Inst.Value"
    Distrbute(0).RightToLeft = False
    Distrbute(1).RightToLeft = False
    Distrbute(1).Caption = "Age Increase Fixed Asset"
    Distrbute(0).Caption = "Distribut.Value To Installments"
    'Label1.Caption = "Manual #"
   'Me.ALLButton1.Caption = "Cost Center"
   ' lbl(15).Caption = "Sales Method"
   ' lbl(16).Caption = "Box Name"
   ' lbl(20).Caption = "General Des"
   ' lbl(21).Caption = "Order No:"
   '
   ' lbl(26).Caption = "Account"
    
    
    
   ' lbl(31).Caption = "Sales Value"
   ' lbl(32).Caption = "Profit Or Loss"
'
'    lbl(26).Caption = "ACC."
'
'    Label8.Caption = "General C. C."

'    With Me.CboPaymentType
'        .Clear
'        .AddItem "Cash"
'        .AddItem "Cheque"
'        .AddItem "Credit"
'        .AddItem "Transfer"
'        .AddItem "Account"
'        .AddItem "Collected Cheque"
'    End With

'    With Me.CboPaymentType1
'        .Clear
'        .AddItem "Expenses"
'        .AddItem "Accounts"
'        .AddItem "Fixed Asset Purchase"
'    End With

  
    
    
    
    
    
'        lbl(40).Caption = "Total Inst."
'    lbl(34).Caption = "EXE Inst."
'    lbl(39).Caption = "Remains Inst."

    

'    lbl(41).Caption = "Current Value"
'    lbl(42).Caption = "Addition Assets"


'    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Asset Additions"
    Me.Ele.Caption = Me.Caption

'    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    
'    Me.lbl(3).Caption = "Expenses Type"
'
'    Me.lbl(0).Caption = "Vendor Bill#"
'    Me.lbl(5).Caption = "Remarks"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."
    
'    Fra.Caption = "GL"
'    lbl(11).Caption = "GL#"
'    lbl(13).Caption = "interval"
'    lbl(9).Caption = "Depit"
'    lbl(10).Caption = "Credit"
'    lbl(17).Caption = "Bank"
'    lbl(18).Caption = "Cheque#"
'    lbl(19).Caption = "Due Date"
'    lbl(22).Caption = "Vendor"
'
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

   ' With Me.Fg_Journal
   '     .TextMatrix(0, .ColIndex("LineNo")) = "Index"
   '     .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
   '     .TextMatrix(0, .ColIndex("value")) = "value"
   '     .TextMatrix(0, .ColIndex("des")) = "description"
   '     .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
'
'    End With

End Sub
