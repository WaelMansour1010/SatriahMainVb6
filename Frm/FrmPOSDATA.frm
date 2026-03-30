VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOSDATA 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
   ClientHeight    =   5148
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5712
   Icon            =   "FrmPOSDATA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5148
   ScaleWidth      =   5712
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1305
      Index           =   1
      Left            =   -2520
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2760
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtopening_balance_voucher_id 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ‰"
         Height          =   255
         Index           =   0
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«∆‰"
         Height          =   255
         Index           =   1
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Õœœ"
         Height          =   255
         Index           =   2
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   210
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.TextBox TxtOpenBalance 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   510
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker Dtp 
         Height          =   330
         Left            =   360
         TabIndex        =   39
         Top             =   870
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   572
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   206438403
         CurrentDate     =   38718
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   285
         Index           =   9
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—’Ìœ "
         Height          =   255
         Index           =   8
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.CheckBox chkChequeBox 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " › Õ Õ«›Ÿ… ‘Ìﬂ« "
      Enabled         =   0   'False
      Height          =   255
      Left            =   -2160
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox XPTxtBoxNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1920
      Width           =   2865
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄ÂœÂ"
      Height          =   195
      Left            =   -240
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ“Ì‰…"
      Height          =   195
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1680
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   555
      Left            =   1380
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3540
      Width           =   2865
   End
   Begin VB.TextBox XPTxtBoxName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1485
      Width           =   2865
   End
   Begin VB.TextBox XPTxtBoxID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   750
      Width           =   1185
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5715
      _cx             =   10075
      _cy             =   1037
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      CaptionStyle    =   1
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   120
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   614
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPOSDATA.frx":038A
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
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   3
         Top             =   120
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   614
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPOSDATA.frx":0724
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
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   614
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPOSDATA.frx":0ABE
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
         Height          =   345
         Index           =   3
         Left            =   615
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   614
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPOSDATA.frx":0E58
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   0
      Left            =   4596
      TabIndex        =   9
      Top             =   4656
      Width           =   708
      _ExtentX        =   1249
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   4656
      Width           =   708
      _ExtentX        =   1249
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Index           =   2
      Left            =   3108
      TabIndex        =   11
      Top             =   4656
      Width           =   708
      _ExtentX        =   1249
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Index           =   3
      Left            =   2352
      TabIndex        =   12
      Top             =   4656
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Index           =   4
      Left            =   1572
      TabIndex        =   13
      Top             =   4656
      Width           =   768
      _ExtentX        =   1355
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Index           =   6
      Left            =   840
      TabIndex        =   14
      Top             =   4656
      Width           =   708
      _ExtentX        =   1249
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Height          =   372
      Left            =   1596
      TabIndex        =   15
      Top             =   4656
      Width           =   792
      _ExtentX        =   1397
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1355
      _ExtentY        =   656
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   7.8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin MSDataListLib.DataCombo DcEmp 
      Height          =   315
      Left            =   1320
      TabIndex        =   28
      Top             =   2640
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   508
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   1380
      TabIndex        =   29
      Top             =   1080
      Width           =   2865
      _ExtentX        =   5059
      _ExtentY        =   508
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCBoxID 
      Height          =   315
      Left            =   1320
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5059
      _ExtentY        =   508
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcSalePriceNames 
      Height          =   312
      Left            =   2280
      TabIndex        =   45
      Top             =   3000
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   508
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· ”⁄Ì—"
      Height          =   312
      Index           =   10
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3120
      Width           =   972
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   7
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4845
      TabIndex        =   30
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„œÌ— «·‰ﬁÿÂ"
      Height          =   315
      Index           =   6
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Œ“Ì‰…"
      Height          =   315
      Index           =   5
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   312
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4200
      Width           =   828
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ "
      Height          =   285
      Index           =   0
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   312
      Index           =   1
      Left            =   4536
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3780
      Width           =   972
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   312
      Index           =   2
      Left            =   3876
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   1152
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   315
      Index           =   3
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   312
      Index           =   4
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4200
      Width           =   1152
   End
End
Attribute VB_Name = "FrmPOSDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip


Function CheckDelete() As Boolean
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String

sql = "SELECT pointid FROM cachierData where pointid =" & val(XPTxtBoxID.text)
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
CheckDelete = True
Else
CheckDelete = False
End If
End Function
Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim FirstPeriodDateInthisYear As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPTxtBoxID.text = CStr(new_id("Tblposdata", "BoxID", "", True))
            '        XPTxtBoxName.SetFocus
            Option1.Enabled = True
            Option2.Enabled = True
            Option1.value = True
 
            Me.dcBranch.BoundText = branch_id
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
        
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Option1.Enabled = False
            Option2.Enabled = False

        Case 2
            '   If SystemOptions.ChequeBox = True And Option1.value = True Then
            '           chkChequeBox.value = vbChecked
            '    Else
            '            chkChequeBox.value = vbUnchecked
            '   End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            
        If CheckDelete() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ Õ–› ‰ﬁÿ… «·»Ì⁄ ·«‰Â«  „ —»ÿÂ« »ﬂ«‘Ì—"
            Else
                MsgBox "Can not be edited. Linked to deliver the custody of the staff"
            End If
            Exit Sub
        End If
            Del_Company

        Case 5

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

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

Private Sub Form_Activate()
    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcEmp
 
    Dcombos.GetBoxes Me.DCBoxID

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  »Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄ "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData  order by branch_id  "
    'fill_combo DcBranch, My_SQL
     Dcombos.GetSalePriceNames dcSalePriceNames
    Dcombos.GetBranches dcBranch
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    If SystemOptions.ChequeBox = False Then
        Me.chkChequeBox.Visible = False
  
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    rs.Open "Tblposdata", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Lbl(5).Caption = "Type"
     Lbl(5).Caption = "Price Plane"
    Option1.Caption = "Box"
    Option2.Caption = "Era"
    chkChequeBox.Caption = "Have Cheque Box"

    Me.Fra(1).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    Lbl(8).Caption = "Balance Value"
    Lbl(9).Caption = "Rec Date"

    Me.Caption = "POS Data"
    EleHeader.Caption = Me.Caption
    Lbl(6).Caption = "Employee"
    Lbl(0).Caption = "Box Code"
    Label3.Caption = "Branch"
    Lbl(3).Caption = " Name Ar"
    Lbl(7).Caption = " Name En"

    Lbl(1).Caption = "Remarks"
    Lbl(2).Caption = "Current Record"
    Lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «·Œ“‰   Ê «·⁄Âœ  "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄"
            Else
                Me.Caption = "POS Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = True
            Me.XPMTxtRemark.locked = True

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
                Me.Caption = " »Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄( ÃœÌœ )"
            Else
                Me.Caption = "POS Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " »Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄( ÃœÌœ )"
            Else
                Me.Caption = "POS Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '       Me.XPBtnMove(0).Enabled = False
            '       Me.XPBtnMove(1).Enabled = False
            '       Me.XPBtnMove(2).Enabled = False
            '       Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄(  ⁄œÌ· )"
            Else
                Me.Caption = "POS Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    XPTxtBoxID.text = IIf(IsNull(rs("BoxID").value), "", val(rs("BoxID").value))
'    DCBoxID.BoundText = IIf(IsNull(rs("BoxID1").value), "", val(rs("BoxID1").value))

    XPTxtBoxName.text = IIf(IsNull(rs("BoxName").value), "", Trim(rs("BoxName").value))
    XPTxtBoxNamee.text = IIf(IsNull(rs("BoxNameE").value), "", Trim(rs("BoxNameE").value))
     dcSalePriceNames.BoundText = IIf(IsNull(rs("priceID").value), "", rs("priceID").value)
    XPMTxtRemark.text = IIf(IsNull(rs("Comments").value), "", Trim(rs("Comments").value))
    DcEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    'If rs("type").value = 0 Then
    'Option1.value = True
    'Else

    'Option2.value = True
    'End If
 
    'If rs("ChequeBox").value = False Then
    'Me.chkChequeBox.value = vbUnchecked
    'ElseIf rs("ChequeBox").value = True Then
    'Me.chkChequeBox.value = vbChecked
    'End If
    
    '  rs("OpenBalanceDate").value = Me.Dtp.value

    'txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    'If Not (IsNull(rs("OpenBalanceDate").value)) Then
    '       Me.Dtp.value = rs("OpenBalanceDate").value
    ' Me.Dtp.Enabled = True
    '   Else
        
    '       Me.Dtp.value = Date
    '  Me.Dtp.Enabled = False
    '   End If
    
    '   If Not IsNull(rs("OpenBalanceType").value) Then
    '       Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))
    '       If rs("OpenBalanceType").value = 0 Then
    '           OptType(0).value = True
    '           OptType_Click 0
    '       ElseIf rs("OpenBalanceType").value = 1 Then
    '           OptType(1).value = True
    '           OptType_Click 1
    '       End If
    '
    '   Else
    '       Me.TxtOpenBalance.text = 0
    '       Me.OptType(2).value = True
    '       OptType_Click 2
    '   End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·Œ“‰ Ê «·⁄Âœ " & CHR(13) & " ﬂÊœ «·Œ“Ì‰…/«·⁄ÂœÂ  " & XPTxtBoxID.text & CHR(13) & " «·›—⁄ " & dcBranch.text & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtBoxName & CHR(13) & " «·‰Ê⁄     "

    If Option1.value = True Then
        LogTextA = LogTextA & " Œ“Ì‰…  "
    ElseIf Option2.value = True Then
        LogTextA = LogTextA & "  ⁄Âœ…  "
  
    End If

    LogTextA = LogTextA & CHR(13) & "› Õ Õ«›Ÿ… «·‘Ìﬂ«  "

    If chkChequeBox.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„ "
    Else
        LogTextA = LogTextA & "·« "
    End If

    LogTextA = LogTextA & CHR(13) & "«”„ «·„ÊŸ›   " & DcEmp.text
                    
    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "„·«ÕŸ«    " & XPMTxtRemark
      
    'sssssssssssssssssss
    LogTexte = "  Save Screen " & " Boxes Data " & CHR(13) & " Code " & XPTxtBoxID.text & CHR(13) & " Branch " & dcBranch.text & CHR(13) & "«Name  " & XPTxtBoxName & CHR(13) & " Type     "

    If Option1.value = True Then
        LogTexte = LogTexte & " Box  "
    ElseIf Option2.value = True Then
        LogTexte = LogTexte & "  Era  "
  
    End If

    LogTexte = LogTexte & CHR(13) & "Open CHeque Box "

    If chkChequeBox.value = vbChecked Then
        LogTexte = LogTexte & "Yes "
    Else
        LogTexte = LogTexte & "No "
    End If

    LogTexte = LogTexte & CHR(13) & " Employee Name" & DcEmp.text
                    
    LogTexte = LogTexte & CHR(13) & "Opening Balance Type"

    If OptType(0).value = True Then
        LogTexte = LogTexte & "Debit"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "Credit"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "Na"
    End If

    LogTexte = LogTexte & CHR(13) & " Opening Balance  Value " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & " Remarks   " & XPMTxtRemark
      
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If Trim(dcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Departement"
            Else
                Msg = "Õœœ «·›—⁄ «Ê·« "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcBranch.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

'        If Trim(DCBoxID.BoundText) = "" Then
'            If SystemOptions.UserInterface = EnglishInterface Then
'                Msg = "Specify Box"
'            Else
'                Msg = "Õœœ «·Œ“Ì‰… «Ê·« "
'            End If
'
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DCBoxID.SetFocus
'            SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
    
        If XPTxtBoxName.text = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·‰ﬁÿ… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtBoxName.SetFocus
            Exit Sub
        End If
    
        Select Case Me.TxtModFlg.text

            Case "N"
                StrSQL = "select * From  Tblposdata where BoxName='" & Trim(XPTxtBoxName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ ‰ﬁÿ… „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  Tblposdata where BoxName='" & Trim(XPTxtBoxName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("BoxID").value <> val(XPTxtBoxID.text) Then
                        Msg = "Â‰«ﬂ ‰ﬁÿ…  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        Select Case Me.TxtModFlg.text

            Case "N"

                '       Dim Account_Code_dynamic As String
                '       Option1.value = True
         
                '       If Option1.value = True Then
                '              Account_Code_dynamic = get_account_code_branch(6, my_branch)
                
                '              If Account_Code_dynamic = "NO branch" Then
                '              MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                '              GoTo ErrTrap
                '              Else
                '              If Account_Code_dynamic = "NO account" Then
                '                 MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··’‰«œÌﬁ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                '              GoTo ErrTrap
                 
                '              End If
                '              End If
                '       Else
         
                '                Account_Code_dynamic = get_account_code_branch(35, my_branch)
                '
                '                If Account_Code_dynamic = "NO branch" Then
                '                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                '                GoTo ErrTrap
                '                Else
                '                If Account_Code_dynamic = "NO account" Then
                '                   MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··⁄Âœ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                '                GoTo ErrTrap
                '
                '                End If
                '                End If
                '
         
                '       End If
        
                rs.AddNew
        End Select

        Cn.BeginTrans
        BeginTrans = True
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("BoxID").value = val(XPTxtBoxID.text)
        rs("BoxName").value = Trim(XPTxtBoxName.text)
        rs("BoxNamee").value = Trim(XPTxtBoxNamee.text)
        rs("priceID").value = IIf(Me.dcSalePriceNames.BoundText = "", Null, Me.dcSalePriceNames.BoundText)
        rs("EmpID").value = IIf(DcEmp.BoundText = "", Null, val(DcEmp.BoundText))
    
        rs("Comments").value = IIf(XPMTxtRemark.text = "", Null, Trim(XPMTxtRemark.text))
     ' rs("BoxID1").value = IIf(DCBoxID.BoundText = "", Null, val(DCBoxID.BoundText))
     
        '    If Option1.value = True Then
        '    rs("type").value = 0
        '    Else
        '    rs("type").value = 1
        '    End If
        '   Dim ParentAccount As String
        '   If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        '       If Me.TxtModFlg.text = "N" Then
 
        '     rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtBoxName.text), True, False, XPTxtBoxNamee.text)
     '   rs("Account_Code").value = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DCBoxID.BoundText)) '«·„»Ì⁄« 
        '       Else
        '            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.text)
        '            Cn.Execute StrSQL, , adExecuteNoRecords
 
        '           If Not IsNull(rs("Account_Code").value) Then
        '               ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtBoxName.text, XPTxtBoxNamee.text, , , , , , , , , , , , , , , , , True
        '           End If
          
        '       End If
          
        '       End If
    
        '   If Val(TxtOpenBalance.text) = 0 Then
        '       txtopening_balance_voucher_id = 0
        '       End If
        '
        '             If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
        '
        '                    If Val(Me.txtopening_balance_voucher_id.text) = 0 Then
        '                            txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
        '
        '                    End If '
        '                End If '
        'rs("opening_balance_voucher_id").value = Val(txtopening_balance_voucher_id.text)

        '  If Me.OptType(2).value = True Then
        '         rs("OpenBalance").value = 0
        '         rs("OpenBalanceType").value = Null
        '     ElseIf Me.OptType(0).value = True Then
        '         rs("OpenBalance").value = Val(Me.TxtOpenBalance.text)
        '         rs("OpenBalanceType").value = 0
        '     ElseIf Me.OptType(1).value = True Then
        '         rs("OpenBalance").value = Val(Me.TxtOpenBalance.text)
        '         rs("OpenBalanceType").value = 1
        '     End If

        rs.update
    
        ' Dim StrDes As String
        '     If SystemOptions.UserInterface = ArabicInterface Then
        '     StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ "
        '     Else
        '     StrDes = " Opening Balance For: "
        '     End If
        '
        '     If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
        '         If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        '             Dim LngDevID As Long
        '             Dim LngOpenID As Long
        '             Dim Account_Code_dynamic1 As String
        '
        'LngOpenID = ModAccounts.AddNewOpenBalance(Val(Me.XPTxtCusID.text), Me.Dtp.value)
        ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        '             LngOpenID = 1
        '             LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
        '
        '             If Me.OptType(0).value = True Then
        '
        '                 Account_Code_dynamic1 = get_account_code_branch(57, my_branch)
        '
        '                 If Account_Code_dynamic1 = "NO branch" Then
        '                     MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        '                     GoTo ErrTrap
        '                   Else
        ''
        '                       If Account_Code_dynamic1 = "NO account" Then
        '                           MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        '                           GoTo ErrTrap
        '                       End If
        '                   End If
        '
        '                   If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBoxName.text) & "  " & Trim$(Me.XPTxtBoxNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
        '                       GoTo ErrTrap
        '                   End If
        '
        '                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBoxName.text) & "  " & Trim$(Me.XPTxtBoxNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
        '                        GoTo ErrTrap
        '                    End If
        '
                    
        '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
           Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
        '         GoTo ErrTrap
        ' End If
        '                ElseIf Me.OptType(1).value = True Then
        '                    Account_Code_dynamic1 = get_account_code_branch(57, my_branch)
        '
        '                    If Account_Code_dynamic1 = "NO branch" Then
        '                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        '                        GoTo ErrTrap
        '                    Else
        '
        '                        If Account_Code_dynamic1 = "NO account" Then
        '                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        '                            GoTo ErrTrap
        '                        End If
        '                    End If
        '
        '                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBoxName.text) & "  " & Trim$(Me.XPTxtBoxNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
        '                        GoTo ErrTrap
        '                    End If
        '
        ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
          Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
        '       GoTo ErrTrap
        'End If
        '                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBoxName.text) & "  " & Trim$(Me.XPTxtBoxNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
        '                        GoTo ErrTrap
        '                    End If
        '                End If
        '              '   update_account_opening_balance rs("Account_Code").value
        '                 'update_account_opening_balance Account_Code_dynamic1
        '
        '            End If
        '        End If

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·Œ“‰…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "BoxID='" & val(XPTxtBoxID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
    On Error GoTo ErrTrap
            
    If XPTxtBoxID.text <> "" Then
        StrAccountCode = rs("Account_Code").value & ""
   
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        '    If Not IsNull(rs("Account_Code1").value) Then
        '    StrSQL = StrSQL & " or   Account_Code1='" & rs("Account_Code1").value & "'"
        '    End If
    
        '    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        '    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        '        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Chr(13)
        '        Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·Œ“‰…"
        '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        Msg = "”Ì „ Õ–› »Ì«‰«  «·Œ“‰… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBoxID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            DeleteOpeningBalance
    
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                If chkChequeBox.value = vbChecked Then
                    StrAccountCode1 = rs("Account_Code1").value
                    ParentAccount = rs("ParentAccount").value

                    If ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                        rs.delete
                        Msg = " „  ⁄„·Ì… «·Õ–›."
                        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            
                    Else
                        GoTo ErrTrap
                    End If

                Else

                    If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                        CuurentLogdata ("D")
                        rs.delete
                    Else
                        Exit Sub
                    End If
                End If
            
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.Description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.text = 0
    Cmd_Click (2)

End Function

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
