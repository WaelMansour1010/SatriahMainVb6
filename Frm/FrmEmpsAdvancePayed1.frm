VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmEmpsAdvancePayed1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄œÌ· /«Ìﬁ«› / —œ  «·”·›"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   Icon            =   "FrmEmpsAdvancePayed1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   12675
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   1440
      Width           =   3615
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   88
         Top             =   0
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "⁄·Ï „” ÊÏ «·”·›…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   89
         Top             =   0
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ﬂ· «·”·›"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox TxtBaseValue 
      Height          =   285
      Left            =   6480
      TabIndex        =   79
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtOldValue 
      Height          =   285
      Left            =   8760
      TabIndex        =   78
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TotalValue 
      Height          =   285
      Left            =   7680
      TabIndex        =   77
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—œ ”·›…"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   4680
      Width           =   5775
      Begin VB.ComboBox CboYear 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰ ÿ—Ìﬁ Œ’„ „»«‘—… „‰ «·—« »"
         Height          =   195
         Index           =   7
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   480
         Width           =   3015
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰ ÿ—Ìﬁ «·”œ«œ ··Õ”«»« "
         Height          =   195
         Index           =   6
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰…"
         Height          =   315
         Index           =   26
         Left            =   1260
         TabIndex        =   84
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·‘Â—"
         Height          =   315
         Index           =   25
         Left            =   3060
         TabIndex        =   83
         Top             =   870
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·Â «·”œ«œ «·Ã“∆Ì"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   2640
      Width           =   5775
      Begin VB.TextBox TxtDiffValuee 
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox TxtPayeValuee 
         Height          =   285
         Left            =   1560
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox TxtValuee 
         Height          =   285
         Left            =   3600
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «÷«›… «·›—ﬁ ⁄·Ì ﬁ”ÿ „Õœœ"
         Height          =   315
         Index           =   5
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1080
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Ê“Ì⁄ «·›—ﬁ ⁄·Ì »«ﬁÌ «·œ›⁄« "
         Height          =   375
         Index           =   4
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1320
         Width           =   5175
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«‰‘«¡ ﬁ”ÿ «ŒÌ— »«·›—ﬁ"
         Height          =   315
         Index           =   3
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1680
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo DcbPay2 
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—ﬁ"
         Height          =   315
         Index           =   22
         Left            =   840
         TabIndex        =   72
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€ «·„—«œ ”œ«œÂ"
         Height          =   315
         Index           =   21
         Left            =   2280
         TabIndex        =   70
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
         Height          =   315
         Index           =   20
         Left            =   4200
         TabIndex        =   68
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ —ﬁ„ «·œ›⁄Â"
         Height          =   285
         Index           =   19
         Left            =   1200
         TabIndex        =   63
         Top             =   960
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·Â «· √ÃÌ·"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   1440
      Width           =   5775
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Ê“Ì⁄ ⁄·Ì »«ﬁÌ «·œ›⁄« "
         Height          =   195
         Index           =   2
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   840
         Width           =   5415
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " —ÕÌ· «·œ›⁄«   Ê≈‰‘«¡ ﬁ”ÿ ÃœÌœ"
         Height          =   195
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " —ÕÌ· Ê«÷«›… «·ﬁÌ„… ⁄·Ì ﬁ”ÿ „Õœœ"
         Height          =   195
         Index           =   0
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   240
         Value           =   -1  'True
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo DcbPay 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ —ﬁ„ «·œ›⁄Â"
         Height          =   285
         Index           =   18
         Left            =   1440
         TabIndex        =   55
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FrmEmpsAdvancePayed1.frx":038A
      Left            =   120
      List            =   "FrmEmpsAdvancePayed1.frx":0397
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1080
      Width           =   2835
   End
   Begin VB.TextBox TxtReson 
      Alignment       =   1  'Right Justify
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   47
      Top             =   6000
      Width           =   11175
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10020
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmEmpsAdvancePayed1.frx":03B6
      Left            =   13680
      List            =   "FrmEmpsAdvancePayed1.frx":03C3
      TabIndex        =   42
      Text            =   "ÌÊ„"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   13560
      TabIndex        =   41
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "«Ìﬁ«› ·„œ…"
      Height          =   255
      Left            =   13440
      TabIndex        =   40
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ Ì«— «·œ›⁄Â"
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
      Height          =   2175
      Index           =   0
      Left            =   13800
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1560
      Width           =   6255
      Begin VB.TextBox TxtPaymentCounts 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   36
         Top             =   2490
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·œ›⁄« "
         Height          =   255
         Index           =   12
         Left            =   5220
         TabIndex        =   37
         Top             =   2550
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   135
         Index           =   11
         Left            =   2460
         TabIndex        =   35
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”œ«œ"
         Height          =   255
         Index           =   10
         Left            =   3300
         TabIndex        =   34
         Top             =   2550
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… —’Ìœ «·„ÊŸ›:"
         Height          =   255
         Index           =   5
         Left            =   900
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.TextBox TxtAdvanceValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10020
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   735
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   12675
      _cx             =   22357
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
      Caption         =   " ⁄œÌ· /«Ìﬁ«› / —œ  «·”·›"
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
         TabIndex        =   4
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
         ButtonImage     =   "FrmEmpsAdvancePayed1.frx":03D6
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
         TabIndex        =   5
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
         ButtonImage     =   "FrmEmpsAdvancePayed1.frx":0770
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
         TabIndex        =   6
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
         ButtonImage     =   "FrmEmpsAdvancePayed1.frx":0B0A
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
         TabIndex        =   7
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
         ButtonImage     =   "FrmEmpsAdvancePayed1.frx":0EA4
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2160
         Picture         =   "FrmEmpsAdvancePayed1.frx":123E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7770
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   41640
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   660
      Left            =   3270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7740
      Width           =   7185
      _cx             =   12674
      _cy             =   1164
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   11
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   5295
         TabIndex        =   12
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   4455
         TabIndex        =   13
         Top             =   75
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
         Left            =   3600
         TabIndex        =   14
         Top             =   75
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
         Left            =   2745
         TabIndex        =   15
         Top             =   315
         Visible         =   0   'False
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
         Left            =   360
         TabIndex        =   16
         Top             =   75
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
      Begin ImpulseButton.ISButton Cmdprint 
         Height          =   405
         Left            =   1455
         TabIndex        =   17
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   714
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
         Left            =   2700
         TabIndex        =   31
         Top             =   0
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
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   7740
      TabIndex        =   18
      Top             =   7320
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   13800
      TabIndex        =   38
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   41640
   End
   Begin MSDataListLib.DataCombo DCbBranch 
      Height          =   315
      Left            =   4200
      TabIndex        =   44
      Top             =   720
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbByeEmp 
      Height          =   315
      Left            =   120
      TabIndex        =   48
      Top             =   720
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   1725
      Left            =   5940
      TabIndex        =   73
      Top             =   1920
      Width           =   5415
      _cx             =   9551
      _cy             =   3043
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmEmpsAdvancePayed1.frx":4EA6
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
      ExplorerBar     =   7
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
   Begin MSComCtl2.DTPicker IsuuDate 
      Height          =   315
      Left            =   10020
      TabIndex        =   75
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   13440
      TabIndex        =   80
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   41640
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg1 
      Height          =   1725
      Left            =   5940
      TabIndex        =   85
      Top             =   4200
      Width           =   5415
      _cx             =   9551
      _cy             =   3043
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
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmEmpsAdvancePayed1.frx":4FEC
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
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Êﬁ› «·›⁄·Ì"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   27
      Left            =   7680
      TabIndex        =   86
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ "
      Height          =   285
      Index           =   24
      Left            =   11550
      TabIndex        =   76
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õœœ «·œ›⁄Â"
      Height          =   285
      Index           =   23
      Left            =   11310
      TabIndex        =   74
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊŸ›"
      Height          =   285
      Index           =   17
      Left            =   11550
      TabIndex        =   50
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁ«∆„ »«·⁄„·Ì…"
      Height          =   285
      Index           =   16
      Left            =   3000
      TabIndex        =   49
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”»»"
      Height          =   285
      Index           =   15
      Left            =   11550
      TabIndex        =   46
      Top             =   6240
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   285
      Index           =   14
      Left            =   6960
      TabIndex        =   45
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " €ÌÌ— «· «—ÌŒ «·Ï"
      Height          =   285
      Index           =   13
      Left            =   13800
      TabIndex        =   39
      Top             =   4560
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   11550
      TabIndex        =   29
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   28
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   2
      Left            =   13950
      TabIndex        =   27
      Top             =   5385
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   9150
      TabIndex        =   26
      Top             =   750
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ…"
      Height          =   270
      Index           =   8
      Left            =   11175
      TabIndex        =   25
      Top             =   7155
      Width           =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2550
      TabIndex        =   24
      Top             =   7350
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   810
      TabIndex        =   23
      Top             =   7350
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   22
      Top             =   7260
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1860
      TabIndex        =   21
      Top             =   7260
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   405
      Index           =   0
      Left            =   30
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "FrmEmpsAdvancePayed1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim EmpReport As ClsEmployeeReport
Dim Employee_account As String
Public Msg1 As String

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            XPTxtID.Text = CStr(new_id("TblEmpAdvance", "AdvanceID", "", True))
            Me.DCboUserName.BoundText = user_id
            Me.DcbBranch.BoundText = Current_branch
            'TxtPaymentCounts.text = 1
            XPDtbTrans.SetFocus
            Rd(0).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
        If val(Combo2.ListIndex) = 2 And Opt(7).value = True Then
        If val(CmbMonth.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·‘Â—"
        Else
        MsgBox "Please Select Month"
        End If
        CmbMonth.SetFocus
        Exit Sub
        End If
        If val(CboYear.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·”‰…"
        Else
        MsgBox "Please Select Year"
        End If
        CboYear.SetFocus
        Exit Sub
        End If
            If ChekPayedSalary(val(CboYear.Text), val(CmbMonth.ListIndex) + 1, val(Me.DcbBranch.BoundText)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ–› ﬁÌœ «·—Ê« »  ··‘Â— «·„Õœœ «Ê·«"
            Else
            MsgBox "Delete Salary Allocation JL"
            End If
            Exit Sub
            End If
        End If
        Dim i As Integer
        Dim FLgCH As Boolean
        Dim FLgProce As Boolean
        FLgCH = False
        FLgProce = False
        With FG
        For i = .FixedRows To .Rows - 1
       If .Cell(flexcpChecked, i, .ColIndex("Checked")) = flexChecked And val(.TextMatrix(i, .ColIndex("AdvanceID"))) <> 0 Then
       FLgCH = True
       End If
        Next i
        End With
       If FLgCH = False Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ «Œ Ì«— œ›⁄… Ê«Õœ… «Ê «ﬂÀ—"
       Else
       MsgBox "Please Select One Payment or More than One"
       End If
       Exit Sub
       End If
              
        If val(Combo2.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·⁄„·Ì…"
        Else
        MsgBox "Please Select Type Process"
        End If
        Combo2.SetFocus
        Exit Sub
        End If
       For i = 0 To 7
       If Opt(i).value = True Then
       FLgProce = True
       End If
       Next i
    If FLgProce = False Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ  ÕœÌœ ⁄„·Ì…"
       Else
       MsgBox "Please Select One Process "
       End If
       Exit Sub
       End If
         If val(DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
        Else
        MsgBox "Please Select Employee"
        End If
        DcboEmpName.SetFocus
        Exit Sub
        End If
        
           If val(DcbBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
        Else
        MsgBox "Please Select Branch"
        End If
        DcbBranch.SetFocus
        Exit Sub
        End If
 
        If val(Combo2.ListIndex) = 0 And Opt(0).value = True Then
        If val(DcbPay.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·œ›⁄…"
        Else
        MsgBox "Please Select Payment"
        End If
             DcbPay.SetFocus
        Exit Sub
        End If
   
        End If
         If val(Combo2.ListIndex) = 2 And Opt(5).value = True Then
        If val(DcbPay2.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·œ›⁄…"
        Else
        MsgBox "Please Select Payment"
        End If
        DcbPay2.SetFocus
        Exit Sub
        End If
        End If
        If val(Combo2.ListIndex) = 2 And Opt(7).value = True Then
         If CheckDate = False Then
            Exit Sub
        End If
       End If
       
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

                  Load FrmNotesSearch
                   FrmNotesSearch.SearchType = 3
                 FrmNotesSearch.show vbModal
        Case 6
            Unload Me

        Case 8
         '   CalCulateParts
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub CmdPrint_Click()
print_report
End Sub

Private Sub Combo2_Change()
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
If Combo2.ListIndex = 0 Then
    Frame1.Enabled = True
    TxtValuee.Text = 0
TxtPayeValuee.Text = 0
TxtDiffValuee.Text = 0
DcbPay2.BoundText = 0
ElseIf Combo2.ListIndex = 1 Then
DcbPay.BoundText = 0
    Frame2.Enabled = True
ElseIf Combo2.ListIndex = 2 Then
TxtValuee.Text = 0
TxtPayeValuee.Text = 0
TxtDiffValuee.Text = 0
DcbPay2.BoundText = 0
DcbPay.BoundText = 0
    Frame3.Enabled = True
End If
RelinGrid
RelinChecked
End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
  
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
     If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    GetEmpAdv val(DcboEmpName.BoundText)
End Sub

Private Sub DcbPay_Change()
DcbPay_Click (0)
End Sub

Private Sub DcbPay_Click(Area As Integer)
RelinChecked
End Sub

Private Sub DcbPay2_Change()
DcbPay2_Click (0)
End Sub

Private Sub DcbPay2_Click(Area As Integer)
RelinChecked
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
      RelinGrid
   ' CalCulateParts
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col = FG.ColIndex("Checked") Then
        Cancel = False
    Else
        Cancel = True
    End If

End Sub
'''Payed1 IS NULL
Function GetCountPayment(Optional advanceID As Double = 0) As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     COUNT(PartNo) AS Cnt"
sql = sql & " From dbo.TblEmpAdvanceDetails"
sql = sql & " Where (Payed Is Null or Payed=0)and (Payed1 Is Null) AND (StutsID Is Null) and (AdvanceID = " & advanceID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetCountPayment = IIf(IsNull(Rs7("Cnt").value), 0, Rs7("Cnt").value)
Else
GetCountPayment = 0
End If
End Function
Function GetCountPayment2(Optional EmpID As Double = 0) As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     COUNT(dbo.TblEmpAdvanceDetails.PartNO) AS Cnt"
sql = sql & " FROM         dbo.TblEmpAdvance INNER JOIN"
sql = sql & "                       dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
sql = sql & "    Where (dbo.TblEmpAdvanceDetails.payed = 0 Or dbo.TblEmpAdvanceDetails.payed Is Null)"
sql = sql & " And (dbo.TblEmpAdvance.Emp_id = " & EmpID & ")"
sql = sql & " and (dbo.TblEmpAdvanceDetails.StutsID Is Null)"
sql = sql & " and (dbo.TblEmpAdvanceDetails.Payed1 Is Null)"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetCountPayment2 = IIf(IsNull(Rs7("Cnt").value), 0, Rs7("Cnt").value)
Else
GetCountPayment2 = 0
End If
End Function
Function GetMaxDatePayment2(Optional EmpID As Double = 0) As Date
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String

sql = " SELECT     max(dbo.TblEmpAdvanceDetails.PartDate) AS Cnt"
sql = sql & " FROM         dbo.TblEmpAdvance INNER JOIN"
sql = sql & "                      dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
'Sql = Sql & " Where (dbo.TblEmpAdvanceDetails.Payed=0 or dbo.TblEmpAdvanceDetails.Payed is null) and (dbo.TblEmpAdvanceDetails.Payed1 Is Null) "
sql = sql & " Where (dbo.TblEmpAdvance.Emp_id = " & EmpID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetMaxDatePayment2 = IIf(IsNull(Rs7("Cnt").value), Date, Rs7("Cnt").value)
Else
GetMaxDatePayment2 = Date
End If
End Function
Function GetMaxDatePayment(Optional advanceID As Double = 0) As Date
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     max(PartDate) AS Cnt"
sql = sql & " From dbo.TblEmpAdvanceDetails"
'Sql = Sql & " Where (Payed=0 or Payed is null) and(Payed1 Is Null)AND "
sql = sql & " where (AdvanceID = " & advanceID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetMaxDatePayment = IIf(IsNull(Rs7("Cnt").value), Date, Rs7("Cnt").value)
Else
GetMaxDatePayment = Date
End If
End Function
Function GetMaxPayment2(Optional EmpID As Double) As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     max(dbo.TblEmpAdvanceDetails.PartNo) AS Cnt"
sql = sql & " FROM         dbo.TblEmpAdvance INNER JOIN"
sql = sql & "                      dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
'Sql = Sql & " Where (dbo.TblEmpAdvanceDetails.payed Is Null or dbo.TblEmpAdvanceDetails.payed=0) and (dbo.TblEmpAdvanceDetails.Payed1 Is Null)and"
sql = sql & " Where (dbo.TblEmpAdvance.Emp_id = " & EmpID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetMaxPayment2 = IIf(IsNull(Rs7("Cnt").value), 0, Rs7("Cnt").value)
Else
GetMaxPayment2 = 0
End If
End Function
Function GetMaxPayment(Optional advanceID As Double) As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     max(PartNo) AS Cnt"
sql = sql & " From dbo.TblEmpAdvanceDetails"
'Sql = Sql & " Where (Payed Is Null or Payed=0) and (Payed1 Is Null)and"
sql = sql & " Where (AdvanceID = " & advanceID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetMaxPayment = IIf(IsNull(Rs7("Cnt").value), 0, Rs7("Cnt").value)
Else
GetMaxPayment = 0
End If
End Function
Sub TabAdvance(Optional TableID As Double = 0, Optional valuee As Double, Optional advanceID As Integer, Optional Remark As String)
Dim StrSQL As String
Dim StutsID As Integer
If val(Combo2.ListIndex) = 0 Then
      If Opt(0).value = True Then
      StutsID = 11
      ElseIf Opt(1).value = True Then
      StutsID = 12
       ElseIf Opt(2).value = True Then
      StutsID = 13
     End If
  ElseIf val(Combo2.ListIndex) = 1 Then
      If Opt(5).value = True Then
      StutsID = 21
      ElseIf Opt(4).value = True Then
      StutsID = 22
       ElseIf Opt(3).value = True Then
      StutsID = 23
      End If
      ElseIf val(Combo2.ListIndex) = 2 Then
      If Opt(6).value = True Then
      StutsID = 31
      ElseIf Opt(7).value = True Then
      StutsID = 32
    
     End If
     
 End If
If StutsID = 32 Then
StrSQL = "Update TblEmpAdvanceDetails Set YearID2=" & val(CboYear.Text) & ",MothID2=" & val(CmbMonth.ListIndex) + 1 & ",  StutsID=" & StutsID & "  ,Remark='" & Remark & "', EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
 ElseIf StutsID = 21 Then
StrSQL = "Update TblEmpAdvanceDetails Set  StutsID=21 ,Payed=0, PartValue=" & val(TxtPayeValuee.Text) & " ,Remark='" & Remark & "', EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
  ElseIf StutsID = 23 Then
StrSQL = "Update TblEmpAdvanceDetails Set  StutsID=23 ,Payed=0, PartValue=" & val(TxtPayeValuee.Text) & " ,Remark='" & Remark & "', EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
   ElseIf StutsID = 22 Then
StrSQL = "Update TblEmpAdvanceDetails Set  StutsID=22 ,Payed=0, PartValue=" & val(TxtPayeValuee.Text) & " ,Remark='" & Remark & "', EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                
Else
StrSQL = "Update TblEmpAdvanceDetails Set Payed=1, StutsID=" & StutsID & "  ,Remark='" & Remark & "', EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
End If
End Sub
Sub TabAdvanceNewPart(Optional Index As Integer, Optional TableID As Double = 0, Optional valuee As Double, Optional advanceID As Double, Optional Remark As String, Optional TableID2 As Integer = 0, Optional EmpID As Double)
Dim Valu1 As Double
Dim MxNo As Double
Dim Rs8 As ADODB.Recordset
Dim sql As String
Dim StrSQL As String
Dim dat1 As Date
Dim cunt As Double
If Index = 1 Then
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue+ " & valuee & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & " "
Cn.Execute StrSQL, , adExecuteNoRecords
ElseIf Index = 11 Then
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue+ " & valuee & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID & " "
Cn.Execute StrSQL, , adExecuteNoRecords
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue - " & valuee & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where TableID=" & TableID2 & ""
Cn.Execute StrSQL, , adExecuteNoRecords

ElseIf Index = 2 Then
If Rd(0).value = True Then
cunt = GetCountPayment2(EmpID)
Else
cunt = GetCountPayment(advanceID)
End If
If cunt <> 0 Then
Valu1 = Round(valuee / cunt, 2)
End If
If Rd(0).value = True Then
  StrSQL = "  Update dbo.TblEmpAdvanceDetails"
  StrSQL = StrSQL & "  Set dbo.TblEmpAdvanceDetails.PartValue = dbo.TblEmpAdvanceDetails.PartValue +" & Valu1 & ""
  StrSQL = StrSQL & " FROM         dbo.TblEmpAdvance INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
  StrSQL = StrSQL & " Where (dbo.TblEmpAdvanceDetails.payed = 0 or dbo.TblEmpAdvanceDetails.payed is null) And (dbo.TblEmpAdvance.Emp_id = " & EmpID & ")"
  StrSQL = StrSQL & " and dbo.TblEmpAdvanceDetails.Payed1 Is Null and (dbo.TblEmpAdvanceDetails.StutsID is null or dbo.TblEmpAdvanceDetails.StutsID=21 or dbo.TblEmpAdvanceDetails.StutsID=22 or dbo.TblEmpAdvanceDetails.StutsID=23)"
  Cn.Execute StrSQL, , adExecuteNoRecords
Else
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue+ " & Valu1 & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where  AdvanceID=" & advanceID & " and(Payed=0 or Payed is null) and Payed1 Is Null and (StutsID is null or StutsID=21 or StutsID=22 or StutsID=23) "

                Cn.Execute StrSQL, , adExecuteNoRecords
End If
 ElseIf Index = 22 Then
If Rd(0).value = True Then
cunt = GetCountPayment2(EmpID)
Else
cunt = GetCountPayment(advanceID)
End If
If cunt <> 0 Then
Valu1 = Round(valuee / cunt, 2)
End If
'StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue+ " & Valu1 & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where  AdvanceID=" & advanceID & " and Payed1 Is Null and (StutsID is null or StutsID=21 or StutsID=22 or StutsID=23) "
'                Cn.Execute StrSQL, , adExecuteNoRecords
If Rd(0).value = True Then
  StrSQL = "  Update dbo.TblEmpAdvanceDetails"
  StrSQL = StrSQL & "  Set dbo.TblEmpAdvanceDetails.PartValue = dbo.TblEmpAdvanceDetails.PartValue +" & Valu1 & ""
  StrSQL = StrSQL & " FROM         dbo.TblEmpAdvance INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
  StrSQL = StrSQL & " Where (dbo.TblEmpAdvanceDetails.payed = 0 or dbo.TblEmpAdvanceDetails.payed is null) And (dbo.TblEmpAdvance.Emp_id = " & EmpID & ")"
  StrSQL = StrSQL & " and dbo.TblEmpAdvanceDetails.Payed1 Is Null and (dbo.TblEmpAdvanceDetails.StutsID is null or dbo.TblEmpAdvanceDetails.StutsID=21  or dbo.TblEmpAdvanceDetails.StutsID=23)"
  Cn.Execute StrSQL, , adExecuteNoRecords
Else
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue+ " & Valu1 & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where  AdvanceID=" & advanceID & " and(Payed=0 or Payed is null) and Payed1 Is Null and (StutsID is null or StutsID=21 or StutsID=22 or StutsID=23) "

                Cn.Execute StrSQL, , adExecuteNoRecords
End If
StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue - " & valuee & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where   TableID=" & TableID2 & " "
                Cn.Execute StrSQL, , adExecuteNoRecords
    ElseIf Index = 33 Then
  If Rd(0).value = True Then
  DTPicker2.value = GetMaxDatePayment2(EmpID)
  Else
  DTPicker2.value = GetMaxDatePayment(advanceID)
  End If
        dat1 = DateAdd("M", 1, DTPicker2.value)
  Set Rs8 = New ADODB.Recordset
        Rs8.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
        Rs8.AddNew
        Rs8("EmpAdPaID").value = val(XPTxtID.Text)
        Rs8("PartValue").value = valuee
        Rs8("AdvanceID").value = advanceID
        Rs8("PartNo").value = MxNo + 1
        Rs8("remark").value = Remark
        Rs8("PartDate").value = dat1
        
        Rs8.update
        StrSQL = "Update TblEmpAdvanceDetails Set PartValue=PartValue - " & valuee & "   , EmpAdPaID =" & val(XPTxtID.Text) & " Where   TableID=" & TableID2 & " "
                Cn.Execute StrSQL, , adExecuteNoRecords
  ElseIf Index = 3 Then
    If Rd(0).value = True Then
  DTPicker2.value = GetMaxDatePayment2(EmpID)
  Else
  DTPicker2.value = GetMaxDatePayment(advanceID)
  End If
        dat1 = DateAdd("M", 1, DTPicker2.value)
  Set Rs8 = New ADODB.Recordset
     If Rd(0).value = True Then
     MxNo = GetMaxPayment2(EmpID)
     Else
     MxNo = GetMaxPayment(advanceID)
     End If
        Rs8.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
        Rs8.AddNew
        Rs8("EmpAdPaID").value = val(XPTxtID.Text)
        Rs8("PartValue").value = valuee
        Rs8("AdvanceID").value = advanceID
        Rs8("PartNo").value = MxNo + 1
        Rs8("remark").value = Remark
        
        Rs8("PartDate").value = dat1
        
        Rs8.update
       
End If

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
YearMonth
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set cmdPrint.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcbByeEmp
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.DcbBranch
    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAdvancePayed  Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    Retrive
    Me.TxtModFlg.Text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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
lbl(25).Caption = "Month"
lbl(26).Caption = "Year"
lbl(27).Caption = "Actual"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    cmdPrint.Caption = "Print"
    Cmd(6).Caption = "Exit"

    lbl(1).Caption = "Date"
  Combo2.Clear
  Combo2.AddItem "postponed"
  Combo2.AddItem "Partial pay"
  Combo2.AddItem "Total Pay"


    With Me.FG1
        .TextMatrix(0, .ColIndex("Checked")) = "select"
        .TextMatrix(0, .ColIndex("AdvanceID")) = "AdvanceID"
        .TextMatrix(0, .ColIndex("PartNO")) = "PartNO"
        .TextMatrix(0, .ColIndex("PartValue")) = "PartValue"
        .TextMatrix(0, .ColIndex("PartDate")) = "PartDate"

    End With
    
    With Me.FG
        .TextMatrix(0, .ColIndex("Checked")) = "select"
        .TextMatrix(0, .ColIndex("AdvanceID")) = "AdvanceID"
        .TextMatrix(0, .ColIndex("PartNO")) = "PartNO"
        .TextMatrix(0, .ColIndex("PartValue")) = "PartValue"
        .TextMatrix(0, .ColIndex("PartDate")) = "PartDate"

    End With
    Opt(7).RightToLeft = False
    Opt(7).Caption = "Payroll Deduction"
    Opt(6).RightToLeft = False
    Opt(6).Caption = "Payment Through Accounts"
    Opt(5).RightToLeft = False
    Opt(3).RightToLeft = False
    Opt(4).RightToLeft = False
    Opt(4).Caption = "Distribution of the rest of the Payments"
    Opt(3).Caption = "Create a new Payment"
    Frame3.Caption = "Pay"
    lbl(19).Caption = "Select"
    Opt(5).Caption = "Add Different to Payment"
    lbl(22).Caption = "Different"
    lbl(21).Caption = "Paid Amount"
    Frame2.Caption = "State of Partial Payment"
    Opt(0).RightToLeft = False
    lbl(18).Caption = "Select"
    Opt(0).Caption = "Deported and to add value to Payment"
    Frame1.Caption = "State of Delay"
    Opt(2).RightToLeft = False
    Opt(2).Caption = "Distribution of the rest of the Payments"
    Opt(1).RightToLeft = False
    Opt(1).Caption = "Deported payments and create a new Payment"
lbl(23).Caption = "Select Payment"
Rd(1).Caption = "By Trans"
Rd(0).Caption = "To All"

    Me.Caption = "Advance Modifications"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "No"
    lbl(20).Caption = "Payment Amount"
    lbl(3).Caption = "Name"
    lbl(2).Caption = "Value"
    lbl(0).Caption = "Box"
    lbl(3).Caption = "Name"
    Fra(0).Caption = "Employee advances"
    lbl(12).Caption = "Count"
    lbl(10).Caption = "Payed"
    lbl(5).Caption = "Balance"
    lbl(8).Caption = "By"
    lbl(16).Caption = "By"
    lbl(17).Caption = "Employee"
    lbl(14).Caption = "Branch"
    lbl(24).Caption = "Due Date"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "Rec. Count"
    lbl(3).Caption = "Operation"
    lbl(15).Caption = "Reason"

End Sub


Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
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
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub


Private Sub IsuuDate_Change()
DcboEmpName_Click (0)
End Sub

Private Sub Opt_Click(Index As Integer)
Dim i As Integer
DcbPay.Enabled = False
DcbPay2.Enabled = False
If Opt(0).value = True Then
DcbPay.Enabled = True
ElseIf Opt(5).value = True Then
DcbPay2.Enabled = True
End If
For i = 0 To 7
If i = Index Then
Opt(i).value = True
Else
Opt(i).value = False
End If
Next i
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
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
            TxtAdvanceValue.locked = True
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
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
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
            TxtAdvanceValue.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰(  ⁄œÌ· )"
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
            TxtAdvanceValue.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPayeValuee_Change()
TxtDiffValuee.Text = val(TxtValuee.Text) - val(TxtPayeValuee.Text)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub TxtValuee_Change()
TxtDiffValuee.Text = val(TxtValuee.Text) - val(TxtPayeValuee.Text)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

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
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecoedDate").value), Date, rs("RecoedDate").value)
    IsuuDate.value = IIf(IsNull(rs("IsuuDate").value), Date, rs("IsuuDate").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.DcbByeEmp.BoundText = IIf(IsNull(rs("ByEmpID").value), "", rs("ByEmpID").value)
    Me.DcbPay.BoundText = IIf(IsNull(rs("PaymentID").value), "", rs("PaymentID").value)
    Me.DcbPay2.BoundText = IIf(IsNull(rs("PaymentID1").value), "", rs("PaymentID1").value)
    Me.Combo2.ListIndex = IIf(IsNull(rs("TypeSele").value), -1, rs("TypeSele").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    TxtReson.Text = IIf(IsNull(rs("Reaseon").value), "", (rs("Reaseon").value))
    TxtValuee.Text = IIf(IsNull(rs("Valuee").value), "", val(rs("Valuee").value))
    TxtPayeValuee.Text = IIf(IsNull(rs("PayeValuee").value), "", val(rs("PayeValuee").value))
    TxtDiffValuee.Text = IIf(IsNull(rs("DiffValuee").value), "", val(rs("DiffValuee").value))
    If Not IsNull(rs("TypRd").value) Then
    If (rs("TypRd").value) = 0 Then
    Rd(1).value = True
    Else
    Rd(0).value = True
    End If
    Else
    Rd(0).value = True
    End If
    If Not (IsNull(rs("TypeOper").value)) Then
    If (rs("TypeOper").value) = 0 Then
    Opt(0).value = True
    ElseIf (rs("TypeOper").value) = 1 Then
    Opt(1).value = True
    ElseIf (rs("TypeOper").value) = 2 Then
    Opt(2).value = True
     ElseIf (rs("TypeOper").value) = 3 Then
    Opt(3).value = True
      ElseIf (rs("TypeOper").value) = 4 Then
    Opt(4).value = True
      ElseIf (rs("TypeOper").value) = 5 Then
    Opt(5).value = True
      ElseIf (rs("TypeOper").value) = 6 Then
    Opt(6).value = True
      ElseIf (rs("TypeOper").value) = 7 Then
    Opt(7).value = True
    
    End If
    End If
    CboYear.ListIndex = IIf(IsNull(rs("YearID").value), -1, val(rs("YearID").value))
    CmbMonth.ListIndex = IIf(IsNull(rs("MonthID").value), -1, val(rs("MonthID").value))
    Set RsDetails = New ADODB.Recordset
    
    StrSQL = "Select * From  TblEmpAdvancePayedDet Where EmpAdPaID=" & val(XPTxtID.Text) & " and TypeID=0 "
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = FG.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        FG.Rows = FG.FixedRows + RsDetails.RecordCount

        For i = Me.FG.FixedRows To FG.Rows - 1
           FG.TextMatrix(i, FG.ColIndex("TableID")) = IIf(IsNull(RsDetails("TableID2").value), 0, RsDetails("TableID2").value)
            FG.TextMatrix(i, FG.ColIndex("AdvanceID")) = IIf(IsNull(RsDetails("AdvanceID").value), 0, RsDetails("AdvanceID").value)
            FG.TextMatrix(i, FG.ColIndex("PartNO")) = IIf(IsNull(RsDetails("PartNO").value), 0, RsDetails("PartNO").value)
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = IIf(IsNull(RsDetails("PartValue").value), 0, RsDetails("PartValue").value)
            FG.TextMatrix(i, FG.ColIndex("PartDate")) = IIf(IsNull(DisplayDate(RsDetails("PartDate").value)), "", DisplayDate(CDate(RsDetails("PartDate").value)))
            FG.Cell(flexcpChecked, i, FG.ColIndex("Checked")) = flexChecked
            RsDetails.MoveNext
        Next i
    End If
    RsDetails.Close
    Set RsDetails = Nothing
   ''/////////////
       Set RsDetails = New ADODB.Recordset
    
    StrSQL = "Select * From  TblEmpAdvancePayedDet Where EmpAdPaID=" & val(XPTxtID.Text) & " and TypeID=1 "
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = FG1.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        FG1.Rows = FG1.FixedRows + RsDetails.RecordCount

        For i = Me.FG1.FixedRows To FG1.Rows - 1
            FG1.TextMatrix(i, FG1.ColIndex("AdvanceID")) = RsDetails("AdvanceID").value
            FG1.TextMatrix(i, FG1.ColIndex("PartNO")) = RsDetails("PartNO").value
            FG1.TextMatrix(i, FG1.ColIndex("PartValue")) = RsDetails("PartValue").value
            FG1.TextMatrix(i, FG1.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
            RsDetails.MoveNext
        Next i
    End If
    RsDetails.Close
    Set RsDetails = Nothing
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Sub RelinChecked()
Dim i As Integer

With FG
For i = .FixedRows To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PartNO"))) <> 0 Then
If .Cell(flexcpChecked, i, .ColIndex("Checked")) = flexChecked Then
If val(Combo2.ListIndex) = 0 Then
If val(.TextMatrix(i, .ColIndex("TableID"))) = val(DcbPay.BoundText) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· —ÕÌ· «·Ï ‰›” «·œ›⁄Â «·„Œ «—…"
Else
MsgBox "You can not convert the value to the same payment"
End If
DcbPay.BoundText = 0
End If
ElseIf val(Combo2.ListIndex) = 1 Then
If val(.TextMatrix(i, .ColIndex("TableID"))) = val(DcbPay2.BoundText) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· —ÕÌ· «·Ï ‰›” «·œ›⁄Â «·„Œ «—…"
Else
MsgBox "You can not convert the value to the same payment"
End If
DcbPay2.BoundText = 0
End If
End If
End If
End If
Next i
End With
End Sub

Sub RelinGrid()
Dim i As Integer
Dim Sm As Double
txtoldvalue.Text = 0
TxtValuee.Text = 0
Sm = 0
With FG
For i = .FixedRows To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PartNO"))) <> 0 Then
If .Cell(flexcpChecked, i, .ColIndex("Checked")) = flexChecked Then
Sm = Sm + val(.TextMatrix(i, .ColIndex("PartValue")))
End If
End If
Next i
If val(Combo2.ListIndex) = 0 Then
txtoldvalue.Text = Sm
ElseIf val(Combo2.ListIndex) = 1 Then
TxtValuee.Text = Sm
End If
End With
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim CountGrid As Double
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim StrAccountCode As String
    Dim cunt As Double
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            Else
            Msg = "Please Select Employee "
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
         SendKeys "{F4}"
            Exit Sub
        End If
With FG
If Opt(2).value = True Or Opt(4).value = True Then
If .Rows = 1 Then
If FG.Cell(flexcpChecked, 1, FG.ColIndex("Checked")) = flexChecked Then
If val(FG.TextMatrix(1, FG.ColIndex("AdvanceID"))) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «·Õ›Ÿ ·«  ÊÃœ œ›«⁄«  „ »ﬁÌ… "
Else
MsgBox "There are no residual payments"
End If
Exit Sub
End If
End If
End If
End If
End With
With FG
CountGrid = 0
For i = 1 To .Rows - 1
If Opt(2).value = True Or Opt(4).value = True Then
If FG.Cell(flexcpChecked, i, .ColIndex("Checked")) = flexChecked Then
CountGrid = CountGrid + 1
If val(FG.TextMatrix(1, FG.ColIndex("AdvanceID"))) <> 0 Then
If Rd(0).value = True Then
cunt = GetCountPayment2(val(Me.DcboEmpName.BoundText))
Else
cunt = GetCountPayment(val(.TextMatrix(i, FG.ColIndex("AdvanceID"))))
End If
If cunt = 0 Or cunt = 1 Or cunt = CountGrid Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «·Õ›Ÿ ·«  ÊÃœ œ›«⁄«  „ »ﬁÌ… "
Else
MsgBox "There are no residual payments"
End If
Exit Sub
End If
End If
End If
End If
Next i
End With

      '  CalCulateParts
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("TblEmpAdvancePayed", "ID", "", True))
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblEmpAdvancePayedDet Where EmpAdPaID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
           
        End If
    
        rs("ID").value = val(XPTxtID.Text)
        rs("RecoedDate").value = XPDtbTrans.value
        rs("IsuuDate").value = IsuuDate.value
        rs("BranchID").value = val(Me.DcbBranch.BoundText)
        rs("EmpID").value = val(Me.DcboEmpName.BoundText)
        rs("ByEmpID").value = val(Me.DcbByeEmp.BoundText)
        rs("TypeSele").value = val(Me.Combo2.ListIndex)
        rs("UserID").value = val(Me.DCboUserName.BoundText)
        rs("Reaseon").value = Me.TxtReson.Text
        If Opt(0).value = True Then
        rs("TypeOper").value = 0
        ElseIf Opt(1).value = True Then
        rs("TypeOper").value = 1
        ElseIf Opt(2).value = True Then
        rs("TypeOper").value = 2
        ElseIf Opt(3).value = True Then
        rs("TypeOper").value = 3
        ElseIf Opt(4).value = True Then
        rs("TypeOper").value = 4
        ElseIf Opt(5).value = True Then
        rs("TypeOper").value = 5
        ElseIf Opt(6).value = True Then
        rs("TypeOper").value = 6
        ElseIf Opt(7).value = True Then
        rs("TypeOper").value = 7
       End If
       If Rd(1).value = True Then
       rs("TypRd").value = 1
       Else
       rs("TypRd").value = 0
       End If
       rs("PaymentID").value = val(Me.DcbPay.BoundText)
       rs("PaymentID1").value = val(Me.DcbPay2.BoundText)
       rs("Valuee").value = val(TxtValuee.Text)
       rs("PayeValuee").value = val(TxtPayeValuee.Text)
       rs("DiffValuee").value = val(TxtDiffValuee.Text)
       rs("YearID").value = val(Me.CboYear.ListIndex)
       rs("MonthID").value = val(Me.CmbMonth.ListIndex)
       
        msgstr
        
        rs.update
        Set RsDetails = New ADODB.Recordset
                        
        RsDetails.Open "TblEmpAdvancePayedDet", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
Dim Remark As String
        For i = Me.FG.FixedRows To FG.Rows - 1

            If FG.Cell(flexcpChecked, i, FG.ColIndex("Checked")) = flexChecked Then
                 RsDetails.AddNew
                    RsDetails("EmpAdPaID").value = val(XPTxtID.Text)
                     RsDetails("TableID2").value = FG.TextMatrix(i, FG.ColIndex("TableID"))
                     RsDetails("PartNO").value = FG.TextMatrix(i, FG.ColIndex("PartNO"))
                     RsDetails("PartValue").value = FG.TextMatrix(i, FG.ColIndex("PartValue"))
                     RsDetails("PartDate").value = FG.TextMatrix(i, FG.ColIndex("PartDate"))
                     RsDetails("AdvanceID").value = FG.TextMatrix(i, FG.ColIndex("AdvanceID"))
                     RsDetails("TypeID").value = 0
                      TabAdvance val(FG.TextMatrix(i, FG.ColIndex("TableID"))), val(FG.TextMatrix(i, FG.ColIndex("PartValue"))), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Msg1
             
                    
                 RsDetails.update
           
      
            End If

        Next i
     '''''''''
             For i = Me.FG.FixedRows To FG.Rows - 1

            If FG.Cell(flexcpChecked, i, FG.ColIndex("Checked")) = flexChecked Then
                      If val(Combo2.ListIndex) = 0 Then
                      If Opt(0).value = True Then
                      TabAdvanceNewPart 1, val(DcbPay.BoundText), val(txtoldvalue.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, , val(DcboEmpName.BoundText)
                      ElseIf Opt(2).value = True Then
                      TabAdvanceNewPart 2, val(DcbPay.BoundText), val(txtoldvalue.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, , val(DcboEmpName.BoundText)
                      ElseIf Opt(1).value = True Then
                      TabAdvanceNewPart 3, 0, val(txtoldvalue.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, , val(DcboEmpName.BoundText)
                      End If
                      
                      ElseIf val(Combo2.ListIndex) = 1 Then
                      If Opt(5).value = True Then
                      TabAdvanceNewPart 11, val(DcbPay2.BoundText), val(TxtDiffValuee.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, val(FG.TextMatrix(i, FG.ColIndex("PartNO"))), val(DcboEmpName.BoundText)
                      ElseIf Opt(4).value = True Then
                      TabAdvanceNewPart 22, , val(TxtDiffValuee.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, val(FG.TextMatrix(i, FG.ColIndex("PartNO"))), val(DcboEmpName.BoundText)
                      ElseIf Opt(3).value = True Then
                      TabAdvanceNewPart 33, 0, val(TxtDiffValuee.Text), val(FG.TextMatrix(i, FG.ColIndex("AdvanceID"))), Remark, val(FG.TextMatrix(i, FG.ColIndex("PartNO"))), val(DcboEmpName.BoundText)
                     End If

                End If
                GoTo l
            End If

        Next i
l:
  ''///////////////////////
            Set RsDetails = New ADODB.Recordset
                        
        RsDetails.Open "TblEmpAdvancePayedDet", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
        For i = Me.FG1.FixedRows To FG1.Rows - 1
        If val(FG1.TextMatrix(i, FG1.ColIndex("PartNO"))) <> 0 Then
                 RsDetails.AddNew
                    RsDetails("EmpAdPaID").value = val(XPTxtID.Text)
                     RsDetails("PartNO").value = FG1.TextMatrix(i, FG1.ColIndex("PartNO"))
                     RsDetails("PartValue").value = FG1.TextMatrix(i, FG1.ColIndex("PartValue"))
                     RsDetails("PartDate").value = FG1.TextMatrix(i, FG1.ColIndex("PartDate"))
                     RsDetails("AdvanceID").value = FG1.TextMatrix(i, FG1.ColIndex("AdvanceID"))
                     RsDetails("TypeID").value = 1
                 RsDetails.update
           End If
        Next i

     '   If detect_employee_work_type = 1 Then
     '       Msg = "—œ ”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
     '       LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
     '       StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
'
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.lbl(11).Caption), 0, Msg, , , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text)) = False Then
'                GoTo ErrTrap
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
'            'StrAccountCode = "a1a3a4" –„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.lbl(11).Caption), 1, Msg, , , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text)) = False Then
'                GoTo ErrTrap
'            End If
'        End If


        Cn.CommitTrans
        BeginTrans = False
        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              Else
              Msg = "This Record Already Saved " & CHR(13)
              Msg = Msg & "Do You Want Enter Another Reacord"
              End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             Else
             MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             End If
        End Select

        TxtModFlg.Text = "R"
    End If

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
        Msg = "Can not Save " & CHR(13)
        Msg = Msg + "have been insert incorrect values " & CHR(13)
        Msg = Msg + "Make sure of the validity of the data"
     End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry ....error douring save data"
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblEmpAdvancePayed.ID, dbo.TblEmpAdvancePayed.RecoedDate, dbo.TblEmpAdvancePayed.BranchID, dbo.TblBranchesData.branch_name, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblEmpAdvancePayed.EmpID, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee4, TblEmployee_1.Emp_Namee3,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee, dbo.TblEmpAdvancePayed.ByEmpID,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name AS ByEmp_Name, TblEmployee_2.Emp_Name1 AS ByEmp_Name1, TblEmployee_2.Emp_Name2 AS ByEmp_Name2,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name3 AS ByEmp_Name3, TblEmployee_2.Emp_Name4 AS ByEmp_Name4, TblEmployee_2.Fullcode AS ByFullcode,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee4 AS ByEmp_NameE4, TblEmployee_2.Emp_Namee3 AS ByEmp_NameE3, TblEmployee_2.Emp_Namee2 AS ByEmp_NameE2,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee1 AS ByEmp_NameE1, TblEmployee_2.Emp_Namee AS ByEmp_NameE, dbo.TblEmpAdvancePayed.TypeOper,"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayed.Reaseon, dbo.TblEmpAdvancePayed.Valuee, dbo.TblEmpAdvancePayed.PayeValuee, dbo.TblEmpAdvancePayed.DiffValuee,"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayed.TypeSele, dbo.TblEmpAdvancePayed.IsuuDate, dbo.TblEmpAdvancePayed.YearID, dbo.TblEmpAdvancePayed.MonthID,"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayedDet.PartDate, dbo.TblEmpAdvancePayedDet.AdvanceID, dbo.TblEmpAdvancePayedDet.Checked1, dbo.TblEmpAdvancePayedDet.PartValue,"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayedDet.PartNO, dbo.TblEmpAdvancePayed.PaymentID, dbo.TblEmpAdvancePayed.PaymentID1,"
MySQL = MySQL & "                      TblEmpAdvanceRequestDetails_2.PartNo AS PartNo1, TblEmpAdvanceRequestDetails_1.PartNo AS PartNo2 ,TblEmpAdvancePayedDet.TypeID"
MySQL = MySQL & " FROM         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayed LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAdvanceRequestDetails TblEmpAdvanceRequestDetails_2 ON"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayed.PaymentID = TblEmpAdvanceRequestDetails_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAdvanceRequestDetails TblEmpAdvanceRequestDetails_1 ON dbo.TblEmpAdvancePayed.PaymentID1 = TblEmpAdvanceRequestDetails_1.id ON"
MySQL = MySQL & "                      TblEmployee_2.Emp_ID = dbo.TblEmpAdvancePayed.ByEmpID ON TblEmployee_1.Emp_ID = dbo.TblEmpAdvancePayed.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpAdvancePayedDet ON dbo.TblEmpAdvancePayed.ID = dbo.TblEmpAdvancePayedDet.EmpAdPaID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblEmpAdvancePayed.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblEmpAdvancePayed.ID = " & val(XPTxtID.Text) & ")"
 
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\AdvanceRequestPayed.rpt"
             
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\AdvanceRequestPayed.rpt"
              
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
    Dim i As Integer
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    Else
    Msg = "Confirm Delete"
   End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
    
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblEmpAdvancePayedDet Where EmpAdPaID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords

              
               ' For i = Me.Fg.FixedRows To Fg.Rows - 1
    '
    '                StrSQL = "Update TblEmpAdvanceDetails Set Payed=Null , OrgTableID = Null Where TableID=" & val(Fg.TextMatrix(i, Fg.ColIndex("TableID"))) & ""
    '                Cn.Execute StrSQL, , adExecuteNoRecords
    '
    '            Next i
    
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
       Else
       Msg = "This process is not available There are no records"
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
 Msg = "Sorry... error douring delete"
 End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Sub msgstr()
If val(Combo2.ListIndex) = 0 Then
If Opt(0).value = True Then
Msg1 = Opt(0).Caption & " " & DcbPay.Text
ElseIf Opt(1).value = True Then
Msg1 = Opt(1).Caption
ElseIf Opt(2).value = True Then
Msg1 = Opt(2).Caption
End If
ElseIf val(Combo2.ListIndex) = 1 Then
If Opt(5).value = True Then
Msg1 = Opt(5).Caption & "  " & DcbPay.Text
ElseIf Opt(3).value = True Then
Msg1 = Opt(3).Caption
ElseIf Opt(4).value = True Then
Msg1 = Opt(4).Caption
End If
ElseIf val(Combo2.ListIndex) = 2 Then
If Opt(6).value = True Then
Msg1 = Opt(6).Caption
ElseIf Opt(7).value = True Then
Msg1 = Opt(7).Caption
End If
End If
End Sub
Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2010 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex
End Sub
Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String
 



    If year(Date) > val(Me.CboYear.Text) Then ' ⁄«„ „÷Ï
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
       Else
       Msg = "Can not be a date before today's date"
       End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckDate = False
        Exit Function
    ElseIf year(Date) = val(Me.CboYear.Text) Then '‰›” «·⁄«„

        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
           If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
       Else
       Msg = "Can not be a date before today's date"
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckDate = False
            Exit Function
        End If
    End If

    CheckDate = True
End Function
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
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, " —œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " —œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " —œ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "—œ ”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmdPrint, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtAdvanceValue.Text, 0)
End Sub


Private Sub GetEmpAdv(Optional EmpID As Double = 0)
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Me.DcboEmpName.BoundText = "" Then
        Me.FG.Rows = FG.FixedRows
        Me.FG1.Rows = FG1.FixedRows
        Me.lbl(9).Caption = 0
    Else
   StrSQL = "  SELECT     dbo.TblEmpAdvanceDetails.PartNO, dbo.TblEmpAdvanceDetails.PartValue, dbo.TblEmpAdvanceDetails.PartDate, dbo.TblEmpAdvanceDetails.TableID,"
   StrSQL = StrSQL & "                   dbo.TblEmpAdvanceDetails.OrgTableID, dbo.TblEmpAdvanceDetails.Payed, dbo.TblEmpAdvanceDetails.StutsID, dbo.TblEmpAdvanceDetails.Payed1,"
   StrSQL = StrSQL & "                     dbo.TblEmpAdvanceDetails.advanceID , dbo.TblEmpAdvance.Emp_id"
   StrSQL = StrSQL & " FROM         dbo.TblEmpAdvance INNER JOIN"
   StrSQL = StrSQL & "                    dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
   StrSQL = StrSQL & "     WHERE     (dbo.TblEmpAdvance.Emp_id= " & EmpID & ") AND (dbo.TblEmpAdvanceDetails.PartDate >=" & SQLDate(IsuuDate.value, True) & ") AND"
   StrSQL = StrSQL & "           (dbo.TblEmpAdvanceDetails.Payed IS NULL or  dbo.TblEmpAdvanceDetails.Payed=0) and (dbo.TblEmpAdvanceDetails.Payed1 IS NULL) and( dbo.TblEmpAdvanceDetails.StutsID IS NULL or dbo.TblEmpAdvanceDetails.StutsID=666 or dbo.TblEmpAdvanceDetails.StutsID=21 or dbo.TblEmpAdvanceDetails.StutsID=22 or dbo.TblEmpAdvanceDetails.StutsID=23) "
        
'        StrSQL = "SELECT     dbo.TblEmpAdvanceRequestDetails.AdvanceID, dbo.TblEmpAdvanceRequestDetails.PartNo, dbo.TblEmpAdvanceRequestDetails.PartValue, "
'        StrSQL = StrSQL & "              dbo.TblEmpAdvanceRequestDetails.PartDate , dbo.TblEmpAdvanceRequest.Emp_id, dbo.TblEmpAdvanceRequestDetails.Payed1"
'        StrSQL = StrSQL & "  FROM         dbo.TblEmpAdvanceRequestDetails RIGHT OUTER JOIN"
'        StrSQL = StrSQL & "              dbo.TblEmpAdvanceRequest ON dbo.TblEmpAdvanceRequestDetails.AdvanceID = dbo.TblEmpAdvanceRequest.AdvanceID"
'         StrSQL = StrSQL & "     WHERE     (dbo.TblEmpAdvanceRequest.Emp_id = " & EmpID & ") AND (dbo.TblEmpAdvanceRequestDetails.PartDate >=" & SQLDate(IsuuDate.value, True) & ") AND"
'        StrSQL = StrSQL & "             (dbo.TblEmpAdvanceRequestDetails.Payed1 IS NULL) and( dbo.TblEmpAdvanceRequestDetails.StutsID IS NULL or dbo.TblEmpAdvanceRequestDetails.StutsID=21 or dbo.TblEmpAdvanceRequestDetails.StutsID=22 or dbo.TblEmpAdvanceRequestDetails.StutsID=23) "
  '  ( dbo.TblEmpAdvanceRequest.AccAproved=1)and
        StrSQL = StrSQL + " Order By PartDate,PartNo "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            '        AdvanceValue
            '      TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
     
            rs.MoveFirst
            Me.FG.Rows = FG.FixedRows + rs.RecordCount
Dcombos.GetAdvancedPartNo DcbPay, EmpID, IsuuDate.value
Dcombos.GetAdvancedPartNo DcbPay2, EmpID, IsuuDate.value
            For i = 1 To rs.RecordCount

                With Me.FG
                    .TextMatrix(i, .ColIndex("TableID")) = IIf(IsNull(rs("TableID").value), "", rs("TableID").value)
                    .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                    .TextMatrix(i, .ColIndex("PartDate")) = IIf(IsNull(rs("PartDate").value), "", rs("PartDate").value)
                    .TextMatrix(i, .ColIndex("PartValue")) = IIf(IsNull(rs("PartValue").value), "", rs("PartValue").value)
                    .TextMatrix(i, .ColIndex("PartNO")) = IIf(IsNull(rs("PartNO").value), "", rs("PartNO").value)
                    '.TextMatrix(I, .ColIndex("TableID")) = IIf(IsNull(Rs("TableID").Value), "", Rs("TableID").Value)
                End With

                rs.MoveNext
            Next i

        Else
            Me.FG.Rows = FG.FixedRows
            Me.lbl(9).Caption = 0
            TxtAdvanceValue.Text = ""
        End If
     ''//////////////////
          Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Me.FG1.Rows = FG1.FixedRows + rs.RecordCount
            For i = 1 To rs.RecordCount

                With Me.FG1

                    .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                    .TextMatrix(i, .ColIndex("PartDate")) = IIf(IsNull(rs("PartDate").value), "", rs("PartDate").value)
                    .TextMatrix(i, .ColIndex("PartValue")) = IIf(IsNull(rs("PartValue").value), "", rs("PartValue").value)
                    .TextMatrix(i, .ColIndex("PartNO")) = IIf(IsNull(rs("PartNO").value), "", rs("PartNO").value)
                End With

                rs.MoveNext
            Next i

        Else
            Me.FG1.Rows = FG1.FixedRows
        End If
    End If

End Sub
