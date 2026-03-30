VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChiqueRelease 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "⁄„·Ì«  «·‘Ìﬂ« "
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   DrawWidth       =   10
   Icon            =   "FrmChiqueRelease.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
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
      Height          =   885
      Index           =   1
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   4440
      Width           =   6495
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   54
         Top             =   180
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
         TabIndex        =   55
         Top             =   510
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   32
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   31
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   30
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   29
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   33
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   510
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtNoteID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   1020
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtChekNum 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   2595
   End
   Begin VB.ComboBox CboDealerType 
      Height          =   315
      Left            =   4530
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   885
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ⁄‰ «·‘Ìﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1875
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   2490
      Width           =   2715
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·‘Ìﬂ"
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
         Height          =   285
         Index           =   21
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1470
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’«œ— ⁄‰"
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
         Height          =   285
         Index           =   16
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
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
         Height          =   285
         Index           =   15
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·√” Õﬁ«ﬁ"
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
         Height          =   285
         Index           =   14
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· Õ—Ì—"
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
         Height          =   285
         Index           =   13
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   2790
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1350
      Width           =   2625
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   630
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   3810
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   630
      Width           =   1635
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   1185
      Left            =   2820
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3180
      Width           =   2625
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   2850
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2445
      Width           =   2595
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Index           =   0
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6645
      _cx             =   11721
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
      Caption         =   "⁄„·Ì«  «·‘Ìﬂ« "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
         Left            =   1185
         TabIndex        =   12
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
         ButtonImage     =   "FrmChiqueRelease.frx":038A
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
         TabIndex        =   13
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
         ButtonImage     =   "FrmChiqueRelease.frx":0724
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
         TabIndex        =   14
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
         ButtonImage     =   "FrmChiqueRelease.frx":0ABE
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
         TabIndex        =   15
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
         ButtonImage     =   "FrmChiqueRelease.frx":0E58
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   990
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   1650
      TabIndex        =   4
      Top             =   1680
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   3930
      TabIndex        =   16
      Top             =   5400
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   2820
      TabIndex        =   8
      Top             =   2820
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   540
      Index           =   1
      Left            =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5820
      Width           =   6585
      _cx             =   11615
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   5820
         TabIndex        =   18
         Top             =   105
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   4920
         TabIndex        =   19
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   4110
         TabIndex        =   20
         Top             =   105
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
         Left            =   3315
         TabIndex        =   21
         Top             =   105
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
         Left            =   2520
         TabIndex        =   22
         Top             =   105
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
         Left            =   30
         TabIndex        =   23
         Top             =   105
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
         Left            =   825
         TabIndex        =   24
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   1710
         TabIndex        =   25
         Top             =   105
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
   End
   Begin ImpulseButton.ISButton CmdSearchTrans 
      Height          =   345
      Left            =   2310
      TabIndex        =   6
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonPositionImage=   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmChiqueRelease.frx":11F2
   End
   Begin MSDataListLib.DataCombo dcbranch 
      Height          =   315
      Left            =   240
      TabIndex        =   62
      Top             =   960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   315
      Index           =   23
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»… —ﬁ„ «·‘Ìﬂ À„ «·÷€ÿ ⁄·Ï «‰ —"
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
      Height          =   405
      Index           =   12
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2040
      Width           =   2205
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·‘Ìﬂ"
      Height          =   315
      Index           =   10
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   2070
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   285
      Index           =   9
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2850
      Width           =   1125
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   5400
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   5400
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   8
      Left            =   5685
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5385
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   285
      Index           =   0
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1350
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1005
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·„»·€"
      Height          =   285
      Index           =   2
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2490
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„ ⁄«„·"
      Height          =   285
      Index           =   3
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1380
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   660
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   5
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3180
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ ⁄«„·"
      Height          =   315
      Index           =   11
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   1125
   End
End
Attribute VB_Name = "FrmChiqueRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch

Private Sub CboDealerType_Change()
    On Error Resume Next
    Dim Dcombos As New ClsDataCombos

    Select Case CboDealerType.ListIndex

        Case 0
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False

        Case 1
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, False
    End Select

    cSearchDcbo.Refresh
End Sub

Private Sub CboDealerType_Click()
    CboDealerType_Change
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            XPTxtID.text = CStr(new_id("TblCheckRelease", "OperaID", "", True))
            Me.DCboUserName.BoundText = user_id
            XPDtbTrans.SetFocus
            Me.Dcbranch.BoundText = branch_id

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText
 
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmChiqueReleaseShearch
            FrmChiqueReleaseShearch.Show vbModal

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdSearchTrans_Click()
    FrmCheckSearch.Show vbModal
End Sub

Private Sub DBCboClientName_Change()
    WriteDev
End Sub

Private Sub DcboBox_Change()
    WriteDev
End Sub

Private Sub DCboCashType_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim intDef As Integer
    Dim Dcombos As New ClsDataCombos

    Select Case DCboCashType.ListIndex

        Case 0
            StrSQL = "SELECT * From TblCustemers where Type=1"
            StrSQL = StrSQL + " and CusID <>2 Order By CusName"
            'fill_combo Me.DBCboClientName, StrSQL
            'intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            'DBCboClientName.BoundText = intDef
            Me.DcboBox.Enabled = True
            Me.lbl(9).Enabled = True

        Case 1
            StrSQL = "SELECT * From TblCustemers where Type=2"
            StrSQL = StrSQL + " and CusID<>1 Order by CusName"
            Me.DcboBox.Enabled = True
            Me.lbl(9).Enabled = True

            'fill_combo Me.DBCboClientName, StrSQL
            'intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            'DBCboClientName.BoundText = intDef
        Case 2
            Me.DcboBox.BoundText = ""
            Me.DcboBox.Enabled = False
            Me.lbl(9).Enabled = False

        Case 3
            Me.DcboBox.BoundText = ""
            Me.DcboBox.Enabled = False
            Me.lbl(9).Enabled = False
    End Select

    cSearchDcbo.Refresh
    WriteDev
    Exit Sub
ErrTrap:
End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "cheque Release"
    Ele(0).Caption = Me.Caption

    lbl(4).Caption = "ID"
    lbl(1).Caption = "Date"
    lbl(0).Caption = "Type"
    lbl(11).Caption = "Name "
 
    lbl(10).Caption = " cheque#"
    lbl(2).Caption = "Amount"
    lbl(9).Caption = "Bank"

    lbl(5).Caption = "Remarks"
    lbl(12).Caption = "enter Chique# then press Enter"
 
    Fra(0).Caption = "Chique Information"
    lbl(13).Caption = "Date"
    lbl(14).Caption = "Due Date"
    lbl(15).Caption = "Bank"
    lbl(16).Caption = "From"
    lbl(21).Caption = "Type"
    Fra(1).Caption = "GL"
    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"
 
    lbl(32).Caption = "Depit"
    lbl(31).Caption = "Credit"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
 
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Dim Msg As String

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    AddTip

    If SystemOptions.UserInterface = EnglishInterface Then
        
        With Me.DCboCashType
            .Clear
            .AddItem "Collecting TO Company"
            .AddItem "Pay"
            '.AddItem "‘Ìﬂ „— œ ⁄·Ï ⁄„Ì· «Ê „Ê—œ"
            '.AddItem "‘Ìﬂ „— œ ⁄·Ï «·‘—ﬂ…"
        End With

        With Me.CboDealerType
            .Clear
            .AddItem "Customer"
            .AddItem "Vendor"
        End With

    Else
        
        With Me.DCboCashType
            .Clear
            .AddItem " Õ’Ì· ‘Ìﬂ ··‘—ﬂ…"
            .AddItem "”œ«œ ‘Ìﬂ ⁄·Ï «·‘—ﬂ…"
            '.AddItem "‘Ìﬂ „— œ ⁄·Ï ⁄„Ì· «Ê „Ê—œ"
            '.AddItem "‘Ìﬂ „— œ ⁄·Ï «·‘—ﬂ…"
        End With

        With Me.CboDealerType
            .Clear
            .AddItem "⁄„Ì·"
            .AddItem "„Ê—œ"
        End With

    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName

    'Dcombos.GetBoxes Me.DcboBox
    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo DcboBox, StrSQL

    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DBCboClientName

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = False
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "TblCheckRelease"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdTable

    SetDtpickerDate XPDtbTrans
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

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
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtChekNum_KeyDown(KeyCode As Integer, _
                               Shift As Integer)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim LngNoteID As Long

    If KeyCode = vbKeyReturn Then
        If Me.TxtChekNum.text = "" Then
            Me.TXTNoteID.text = ""
            Me.lbl(17).Caption = ""
            Me.lbl(18).Caption = ""
            Me.lbl(19).Caption = ""
            Me.lbl(20).Caption = ""
        Else
            Set rs = New ADODB.Recordset
            StrSQL = "Select Notes.NoteID  From NOTES where ChqueNum='" & Trim(Me.TxtChekNum.text) & "'"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                LngNoteID = rs("NoteID").value
                rs.Close
                Set rs = Nothing
                Me.TXTNoteID.text = LngNoteID
                GetCheckInfo LngNoteID
            Else
                Me.TXTNoteID.text = ""
                Me.lbl(17).Caption = ""
                Me.lbl(18).Caption = ""
                Me.lbl(19).Caption = ""
                Me.lbl(20).Caption = ""
            End If
        End If
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '       Me.Caption = " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
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
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            XPMTxtRemarks.locked = True
            DCboCashType.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            CmdSearchTrans.Enabled = False
            CboDealerType.locked = True
            Me.DBCboClientName.locked = True

        Case "N"
            '       Me.Caption = " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False

            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DCboCashType.ListIndex = 0
            CmdSearchTrans.Enabled = True
            CboDealerType.locked = False
            Me.DBCboClientName.locked = False
        
            lbl(17).Caption = ""
            lbl(18).Caption = ""
            lbl(19).Caption = ""
            lbl(20).Caption = ""
        
        Case "E"
            '       Me.Caption = " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« (  ⁄œÌ· )"
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
            XPDtbTrans.Enabled = True
            XPMTxtRemarks.locked = False
            DCboCashType.locked = False
            CmdSearchTrans.Enabled = True
            CboDealerType.locked = False
            Me.DBCboClientName.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNoteID_Change()

    If Me.TXTNoteID.text <> "" Then
        GetCheckInfo val(Me.TXTNoteID.text)
    Else
    End If

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
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "OperaID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("OperaID").value), "", val(rs("OperaID").value))
    XPDtbTrans.value = IIf(IsNull(rs("OperaDate").value), Date, rs("OperaDate").value)
    XPTxtVal.text = IIf(IsNull(rs("NoteValue").value), "", Trim(rs("NoteValue").value))

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", Trim(rs("NoteID").value))
    GetCheckInfo val(Me.TXTNoteID.text)

    XPMTxtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)

    If rs("OperaType").value = 0 Then
        DCboCashType.ListIndex = 0
    ElseIf rs("OperaType").value = 1 Then
        DCboCashType.ListIndex = 1
    ElseIf rs("OperaType").value = 2 Then
        DCboCashType.ListIndex = 2
    ElseIf rs("OperaType").value = 3 Then
        DCboCashType.ListIndex = 3
    End If

    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    '-----------------------------------------------------------------------------
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where OperaID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(33).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    End If

    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim BeginTrans As Boolean

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        If DCboCashType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄  Õ’Ì· Ê”œ«œ «·‘Ìﬂ«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboCashType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        ElseIf Me.DCboCashType.ListIndex = 0 Then

            If val(Me.lbl(22).Tag) = 13 Then
                Msg = "‰Ê⁄ «·‘Ìﬂ «·„Œ «— ( ‘Ìﬂ ⁄·Ï «·‘—ﬂ… -‘Ìﬂ ’«œ— „‰ «·‘—ﬂ…) ..!!"
                Msg = Msg & Chr(13) & "ÌÃ» «‰  Œ «— «‰  ﬂÊ‰ «·⁄„·Ì… ”œ«œ ‘Ìﬂ ⁄·Ï «·‘—ﬂ…."
                Msg = Msg & Chr(13) & "«Ê  ”ÃÌ· ‘Ìﬂ „— œ ⁄·Ï «·‘—ﬂ…"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCboCashType.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.DCboCashType.ListIndex = 1 Then

            If val(Me.lbl(22).Tag) = 2 Then
                Msg = "‰Ê⁄ «·‘Ìﬂ «·„Œ «— ( ‘Ìﬂ ··‘—ﬂ… ) ..!!"
                Msg = Msg & Chr(13) & "ÌÃ» «‰  Œ «— «‰  ﬂÊ‰ «·⁄„·Ì…  Õ’Ì· ‘Ìﬂ ··‘—ﬂ…."
                Msg = Msg & Chr(13) & "«Ê  ”ÃÌ· ‘Ìﬂ „— œ ⁄·Ï ⁄„Ì· «Ê „Ê—œ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCboCashType.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If
    
        If DBCboClientName.text = "" Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If XPTxtVal.text = "" Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
            Msg = "ﬁÌ„…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.DCboCashType.ListCount = 0 Or Me.DCboCashType.ListIndex = 1 Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "›Ï Õ«·…  Õ’Ì· «Ê ”œ«œ «·‘Ìﬂ« "
                Msg = Msg & Chr(13) & "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        End If

        If CheckCheckState = False Then
            Exit Sub
        End If

        If Me.DCboCashType.ListIndex = 1 Then
            '”œ«œ ‘Ìﬂ ÌÃ» «·ﬂ‘› ⁄‰ —’Ìœ «·Œ“‰…
            '  If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), Me.XPDtbTrans.value, True) = False Then
            '      Exit Sub
            '  End If
        End If

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            rs.AddNew
            rs("OperaID").value = val(Me.XPTxtID.text)
        End If

        rs("OperaDate").value = Me.XPDtbTrans.value

        If DCboCashType.ListIndex = 0 Then
            rs("OperaType").value = 0
        ElseIf DCboCashType.ListIndex = 1 Then
            rs("OperaType").value = 1
        ElseIf DCboCashType.ListIndex = 2 Then
            rs("OperaType").value = 2
        ElseIf DCboCashType.ListIndex = 3 Then
            rs("OperaType").value = 3
        End If

        rs("NoteID").value = IIf(Me.TXTNoteID.text = "", Null, val(TXTNoteID.text))
        '   rs("branch_no").value = Val(Me.dcbranch.BoundText)
        '    Rs("BankID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
        'Rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
        rs("NoteValue").value = val(XPTxtVal.text)
        rs("Remarks").value = IIf(Trim(Me.XPMTxtRemarks.text) = "", Null, Trim(Me.XPMTxtRemarks.text))
        rs("UserID").value = user_id
                
        rs.update

        '==========================================================================
        ' ”ÃÌ· ﬁÌÊœ
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
            RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            '«·ÿ—› «·„œÌ‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("Notes_ID").value = Null
            RsDev("OperaID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("Notes_ID").value = Null
            RsDev("OperaID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            LblDevID.Caption = LngDevID
            lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

        '==========================================================================
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If

        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "OperaID=" & val(XPTxtID.text) & "", , adSearchForward, adBookmarkFirst

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

    If XPTxtID.text <> "" Then
        If Me.DCboCashType.ListIndex = 0 Then
            If CheckBoxAccount(val(Me.DcboBox.BoundText), val(Me.XPTxtVal.text), Date, False) = False Then
                Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If

        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where OperaID=" & val(XPTxtID.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
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
            'Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« ", 1, 15204351, -2147483630
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

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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
ErrTrap:         End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
    'Dim Msg As String
    'Dim RsTemp As ADODB.Recordset
    'Dim DblCreditNoteValue As Double
    'Dim LngDebitNoteID As Long
    'Dim StrSQL As String
    '
    'CheckDebitTrans = False
    'If LngTransID = 0 Then
    '    Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
    '    Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '
    '    Exit Function
    'ElseIf LngTransID <> 0 Then
    '    Set RsTemp = New ADODB.Recordset
    '    StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
    '    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    '    If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '        If RsTemp("PaymentType").Value = 0 Then
    '            'Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '            Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            'TxtTransSerial.SetFocus
    '            Exit Function
    '        End If
    '        If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").Value), "", RsTemp("CusID").Value) Then
    '            'Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    ''            TxtTransSerial.SetFocus
    '            Exit Function
    '        End If
    '        If LngTransID <> Val(Me.TxtTransID.Text) Then
    '            Me.TxtTransID.Text = LngTransID
    '        End If
    '
    '        DblCreditNoteValue = 0
    '        StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
    '        "Transactions.Transaction_Type, Transactions.PaymentType, " & _
    '        "Notes.Note_Value, Notes.OperaID "
    '        StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & _
    '        "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
    '        Set RsTemp = New ADODB.Recordset
    '        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    '        If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '            LngDebitNoteID = RsTemp("OperaID").Value
    '            DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").Value), 0, RsTemp("Note_Value").Value)
    '            '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
    '            'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
    '            StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
    '            Set RsTemp = New ADODB.Recordset
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    '            If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '                If RsTemp.RecordCount > 0 Then
    '                    Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
    '                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
    '                    Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
    '                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                    Exit Function
    '                End If
    '            End If
    '        Else
    '        'LngDebitNoteID
    '            Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Exit Function
    '        End If
    '        If DblCreditNoteValue < Val(Me.XPTxtVal.Text) Then
    '            Msg = "⁄›Ê« ..."
    '            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
    '            Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
    '            Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
    '            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Me.XPTxtVal.SetFocus
    '            Exit Function
    '        End If
    '        Set RsTemp = New ADODB.Recordset
    '        StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
    '        "Transactions.Transaction_Type, Transactions.PaymentType," & _
    '        "Sum(Notes.Note_Value) AS SumNote_Value "
    '        StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & _
    '        "Notes.Transaction_ID " & _
    '        " Where ((Notes.NoteType = 4 OR Notes.NoteType = 9) And Transactions.Transaction_ID = " & LngTransID & ")"
    '        If Me.TxtModFlg.Text = "E" Then
    '            StrSQL = StrSQL + " And Notes.OperaID <>" & Me.XPTxtID.Text & ""
    '        End If
    '        StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
    '        "Transactions.Transaction_Type, Transactions.PaymentType "
    '        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    '        If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '            If DblCreditNoteValue = RsTemp("SumNote_Value").Value Then
    '                Msg = "⁄›Ê« ...!!!!!" & Chr(13)
    '                Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
    '                Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Function
    '            ElseIf RsTemp("SumNote_Value").Value + Val(Me.XPTxtVal.Text) > _
    '                DblCreditNoteValue Then
    '                Msg = "⁄›Ê« ..."
    '                Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) „”»ﬁ« ·Â–Â «·›« Ê—…"
    '                Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
    '                Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
    '                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
    '                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
    '                Msg = Msg & Chr(13) & "ﬁÌ„…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").Value
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Function
    '            End If
    '        End If
    '    Else
    '        Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '        Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
    '        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        TxtTransSerial.SetFocus
    '        Exit Function
    '    End If
    'End If
    'CheckDebitTrans = True
    'Exit Function
    'ErrTrap:
End Function

Private Function CheckDebitMaintaince(LngTransID As Long) As Boolean
    'Dim Msg As String
    'Dim RsTemp As ADODB.Recordset
    'Dim DblCreditNoteValue As Double
    'Dim LngDebitNoteID As Long
    'Dim StrSQL As String
    '
    'CheckDebitMaintaince = False
    'If LngTransID = 0 Then
    '    Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
    '    Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    'TxtTransSerial.SetFocus
    '    Exit Function
    'ElseIf LngTransID <> 0 Then
    '    Set RsTemp = New ADODB.Recordset
    '    StrSQL = "Select CusID,PaymentType From TblMaintenece where MaintananceID=" & LngTransID & ""
    '    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    '    If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '        If RsTemp("PaymentType").Value = 0 Then
    '            'Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '            Msg = Msg & Chr(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            TxtTransSerial.SetFocus
    '            Exit Function
    '        End If
    '        If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").Value), "", RsTemp("CusID").Value) Then
    '            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '            Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            TxtTransSerial.SetFocus
    '            Exit Function
    '        End If
    '        If LngTransID <> Val(Me.TxtTransID.Text) Then
    '            Me.TxtTransID.Text = LngTransID
    '        End If
    '
    '        DblCreditNoteValue = 0
    '        StrSQL = "SELECT Notes.Note_Value, Notes.OperaID, TblMaintenece.MaintananceID," & _
    '        "TblMaintenece.PaymentType, TblMaintenece.MType "
    '        StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON " & _
    '        "TblMaintenece.MaintananceID = Notes.MaintananceID " & _
    '        " WHERE (((Notes.NoteType)=1)) AND TblMaintenece.MaintananceID=" & LngTransID & ""
    '        Set RsTemp = New ADODB.Recordset
    '        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '        If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '            LngDebitNoteID = RsTemp("OperaID").Value
    '            DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").Value), 0, RsTemp("Note_Value").Value)
    '            '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
    '            'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
    '            StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
    '            Set RsTemp = New ADODB.Recordset
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    '            If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '                If RsTemp.RecordCount > 0 Then
    '                    Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
    '                    Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
    '                    Msg = Msg & Chr(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
    '                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                    Exit Function
    '                End If
    '            End If
    '        Else
    '        'LngDebitNoteID
    '            Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Exit Function
    '        End If
    '        If DblCreditNoteValue < Val(Me.XPTxtVal.Text) Then
    '            Msg = "⁄›Ê« ..."
    '            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
    '            Msg = Msg & Chr(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
    '            Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
    '            Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Me.XPTxtVal.SetFocus
    '            Exit Function
    '        End If
    '        Set RsTemp = New ADODB.Recordset
    '
    '        StrSQL = "SELECT  TblMaintenece.MaintananceID," & _
    '        "TblMaintenece.MType, TblMaintenece.PaymentType," & _
    '        "Sum(Notes.Note_Value) AS SumNote_Value "
    '        StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON TblMaintenece.MaintananceID =" & _
    '        "Notes.MaintananceID " & _
    '        " Where ((Notes.NoteType = 4) And TblMaintenece.MaintananceID = " & LngTransID & ")"
    '        If Me.TxtModFlg.Text = "E" Then
    '            StrSQL = StrSQL + " And Notes.OperaID <>" & Me.XPTxtID.Text & ""
    '        End If
    '        StrSQL = StrSQL + " GROUP BY TblMaintenece.MaintananceID," & _
    '        "TblMaintenece.MType, TblMaintenece.PaymentType"
    '
    '        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    '        If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '            If DblCreditNoteValue = RsTemp("SumNote_Value").Value Then
    '                Msg = "⁄›Ê« ...!!!!!"
    '                Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
    '                Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Function
    '            ElseIf RsTemp("SumNote_Value").Value + Val(Me.XPTxtVal.Text) > _
    '                DblCreditNoteValue Then
    '                Msg = "⁄›Ê« ..."
    '                Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  „”»ﬁ« ·Â–Â «·›« Ê—…"
    '                Msg = Msg & Chr(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
    '                Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
    '                Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
    '                Msg = Msg & Chr(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
    '                Msg = Msg & Chr(13) & "ﬁÌ„…  Õ’Ì· Ê”œ«œ «·‘Ìﬂ«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").Value
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Function
    '            End If
    '        End If
    '    Else
    '        Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
    '        Msg = Msg & Chr(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
    '        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        TxtTransSerial.SetFocus
    '        Exit Function
    '    End If
    'End If
    'CheckDebitMaintaince = True
    'Exit Function
    'ErrTrap:
End Function

Public Function CheckDebitService()

End Function

Private Sub GetCheckInfo(LngNoteID As Long)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim IntTemp As Integer

    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, BanksData.BankName, Notes.ChqueNum, Notes.DueDate," & "Transactions.Transaction_Serial, Transactions.Transaction_Date," & "TblNotesTypes.NotesTypeName, TransactionTypes.TransactionTypeName," & "TblCustemers.CusName, TblMaintenece.MaintananceID, Notes.BankID,TblCustemers.CusID "
    StrSQL = StrSQL + " FROM TransactionTypes RIGHT JOIN (Transactions RIGHT JOIN " & "(TblNotesTypes INNER JOIN (TblMaintenece RIGHT JOIN (TblCustemers RIGHT JOIN " & "(BanksData RIGHT JOIN Notes ON BanksData.BankID = Notes.BankID) " & "ON TblCustemers.CusID = Notes.CusID) ON TblMaintenece.MaintananceID = " & "Notes.MaintananceID) ON TblNotesTypes.NotesType = Notes.NoteType) ON " & "Transactions.Transaction_ID = Notes.Transaction_ID) " & "ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type"

    StrSQL = StrSQL + " Where  Notes.NoteID=" & LngNoteID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.TxtChekNum.text = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)

        If Not IsNull(rs("NoteDate").value) Then
            Me.lbl(17).Caption = Format(rs("NoteDate").value, "yyyy/M/d")
        Else
            Me.lbl(17).Caption = ""
        End If
    
        If Not IsNull(rs("DueDate").value) Then
            Me.lbl(18).Caption = Format(rs("DueDate").value, "yyyy/M/d")
        Else
            Me.lbl(18).Caption = ""
        End If

        If rs("NoteType").value = 2 Then
            Me.lbl(22).Caption = "‘Ìﬂ ··‘—ﬂ…"
        ElseIf rs("NoteType").value = 13 Then
            Me.lbl(22).Caption = "‘Ìﬂ ⁄·Ï «·‘—ﬂ…"
        End If

        Me.lbl(22).Tag = rs("NoteType").value
        Me.lbl(19).Caption = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
    
        If IsNull(rs("TransactionTypeName").value) Then
            Me.lbl(20).Caption = "Õ—ﬂ… ’Ì«‰… —ﬁ„ " & IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
        Else
            Me.lbl(20).Caption = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value) & " —ﬁ„ " & IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
        End If

        XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)

        If Not IsNull(rs("CusID").value) Then
            IntTemp = GetDealerType(rs("CusID").value)

            If IntTemp = -1 Then
                Me.CboDealerType.ListIndex = -1
            Else
                Me.CboDealerType.ListIndex = IntTemp - 1
            End If

            If Me.DBCboClientName.BoundText <> rs("CusID").value Then
                Me.DBCboClientName.BoundText = rs("CusID").value
            End If

        Else
            Me.DBCboClientName.BoundText = ""
            Me.CboDealerType.ListIndex = -1
        End If
    
    End If

End Sub

Private Function CheckCheckState() As Boolean
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    '---------------------------------------------------------------------------------------------
    StrSQL = "Select * From NOTES Where ChqueNum='" & Trim$(Me.TxtChekNum.text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        If rs("CusID").value <> val(Me.DBCboClientName.BoundText) Then
            Msg = "«·‘Ìﬂ «·„Õœœ €Ì— „”Ã· „⁄ Â–« «·⁄„Ì· «Ê«·„Ê—œ «·–Ï ﬁ„  » ⁄œÌ·Â"
            Msg = Msg & Chr(13) & "·«Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰« ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Function
        End If

        If rs("NoteDate").value > Me.XPDtbTrans.value Then
            Msg = "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «·⁄„·Ì… »⁄œ  «—ÌŒ  Õ—Ì— «·‘Ìﬂ"
            Msg = Msg & Chr(13) & "·«Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰« ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Function
        End If

    Else
        Msg = "—ﬁ„ «·‘Ìﬂ «·„œŒ· €Ì— ’ÕÌÕ"
        Msg = Msg & Chr(13) & "·«Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰« ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    '---------------------------------------------------------------------------------------------
    StrSQL = "SELECT TblCheckRelease.OperaID, TblCheckRelease.OperaDate, TblCheckRelease.OperaType," & "TblCheckRelease.NoteValue, TblCheckRelease.Remarks, TblBoxesData.BoxName, TblUsers.UserName," & "Notes.NoteDate, Notes.DueDate, BanksData.BankName, TblCheckReleaseType.CheckTypeName," & "Notes.ChqueNum, TblCustemers.CusName, TblCustemers.Type "
    StrSQL = StrSQL + " FROM TblCustemers INNER JOIN (TblCheckReleaseType INNER JOIN " & "((BanksData INNER JOIN Notes ON BanksData.BankID = Notes.BankID) INNER JOIN " & "((TblCheckRelease INNER JOIN TblBoxesData ON TblCheckRelease.BoxID = TblBoxesData.BoxID)" & "INNER JOIN TblUsers ON TblCheckRelease.UserID = TblUsers.UserID) " & "ON Notes.NoteID = TblCheckRelease.NoteID) ON TblCheckReleaseType.CheckTypeID =" & "TblCheckRelease.OperaType) ON TblCustemers.CusID = Notes.CusID "

    StrSQL = StrSQL + " Where TblCheckRelease.NoteID=" & Me.TXTNoteID.text & ""

    If Me.TxtModFlg.text = "E" Then
        StrSQL = StrSQL + " And TblCheckRelease.OperaID <> " & Me.XPTxtID.text & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        CheckCheckState = True
    Else
        Msg = "⁄›Ê«..."

        If rs("OperaType").value = 0 Then
            Msg = Msg & Chr(13) & "·ﬁœ  „  Õ’Ì· Â–« «·‘Ìﬂ „”»ﬁ«"
        ElseIf rs("OperaType").value = 1 Then
            Msg = Msg & Chr(13) & "·ﬁœ  „ ”œ«œ Â–« «·‘Ìﬂ „”»ﬁ«"
        End If

        Msg = Msg & Chr(13) & "»ÌÌ«‰«  «·‘Ìﬂ:-"
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "—ﬁ„ «·‘Ìﬂ: " & IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "«”„ «·»‰ﬂ: " & IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "«”„ «·„ ⁄«„·: " & IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "ﬁÌ„… «·‘Ìﬂ: " & IIf(IsNull(rs("NoteValue").value), "", rs("NoteValue").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & " «—ÌŒ «· Õ—Ì—: " & IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & " «—ÌŒ «·√” Õﬁ«ﬁ: " & IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)

        If rs("OperaType").value = 0 Then
            Msg = Msg & Chr(13) & "»Ì«‰«  ⁄„·Ì… «· Õ’Ì·:-"
        ElseIf rs("OperaType").value = 1 Then
            Msg = Msg & Chr(13) & "»Ì«‰«  ⁄„·Ì… «·”œ«œ:-"
        End If

        Msg = Msg & Chr(13) & String(10, Chr(32)) & "—ﬁ„ «·⁄„·Ì…: " & IIf(IsNull(rs("OperaID").value), "", rs("OperaID").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & " «—ÌŒ «·⁄„·Ì…: " & IIf(IsNull(rs("OperaDate").value), "", rs("OperaDate").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "«”„ «·Œ“‰…: " & IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "«”„ «·„ ”Œœ„ «·„Õ——: " & IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
        Msg = Msg & Chr(13) & String(10, Chr(32)) & "«·„·«ÕŸ«  «·„”Ã·… ⁄·Ï «·⁄„·Ì…: " & Chr(13) & IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CheckCheckState = False
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub WriteDev()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.DCboCashType.ListIndex = 0 Then
            ' Õ’Ì· ‘Ìﬂ ··‘—ﬂ…
        
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBox.BoundText))
            'Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBox.BoundText), "Account_Code1")
                Else
                    Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBox.BoundText), "Account_Code")
                End If
            End If
        
            'Me.DcboCreditSide.BoundText = "a1a2a4"
        ElseIf Me.DCboCashType.ListIndex = 1 Then
            '”œ«œ ‘Ìﬂ ⁄·Ï «·‘—ﬂ…
        
            'Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    Me.DcboDebitSide.BoundText = get_bank_Account(val(Me.DcboBox.BoundText), "Account_Code2")
                Else
                    Me.DcboDebitSide.BoundText = get_bank_Account(val(Me.DcboBox.BoundText), "Account_Code")
                End If
            End If
        
            '        Me.DcboDebitSide.BoundText = "a2a3a2"
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBox.BoundText))
            'Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
        ElseIf Me.DCboCashType.ListIndex = 2 Then
            '‘Ìﬂ „— œ ⁄·Ï ⁄„Ì· «Ê „Ê—œ
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            Me.DcboCreditSide.BoundText = "a1a2a4"
        ElseIf Me.DCboCashType.ListIndex = 3 Then
            '‘Ìﬂ „— œ ⁄·Ï «·‘—ﬂ…
            Me.DcboDebitSide.BoundText = "a2a3a2"
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        Else
            Me.DcboDebitSide.BoundText = ""
            Me.DcboCreditSide.BoundText = ""
        End If
    End If

End Sub
