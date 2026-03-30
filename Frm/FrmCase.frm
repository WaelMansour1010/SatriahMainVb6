VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCase 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·«’Ê· «·À«» …"
   ClientHeight    =   7890
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10155
   Icon            =   "FrmCase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   10155
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   " „ «·«Â·«ﬂ"
      Height          =   195
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "„ Êﬁ›"
      Height          =   195
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtinstallDo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   3000
      Width           =   1965
   End
   Begin VB.TextBox txtinstallmentresult 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   3000
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Ã„Ê⁄Â «·«’·"
      Height          =   2895
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   3480
      Width           =   6375
      Begin VB.TextBox TXT24 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1080
         Width           =   3885
      End
      Begin VB.TextBox TXT26 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1440
         Width           =   3885
      End
      Begin VB.TextBox TXT25 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Width           =   3885
      End
      Begin VB.TextBox TXT31 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2160
         Width           =   3885
      End
      Begin VB.TextBox TXT40 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   2520
         Width           =   3885
      End
      Begin VB.TextBox txtPercentage2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   720
         Width           =   3885
      End
      Begin VB.TextBox TXtPercentage1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   360
         Width           =   3885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» «·«’·  »«·„Ì“«‰Ì…"
         Height          =   255
         Index           =   19
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» „Ã„⁄ «·«Â·«ﬂ"
         Height          =   255
         Index           =   20
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»    „’—Ê›«  «·«Â·«ﬂ"
         Height          =   255
         Index           =   21
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   «—»«Õ »Ì⁄"
         Height          =   255
         Index           =   22
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   2160
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   Œ”«∆— »Ì⁄"
         Height          =   255
         Index           =   23
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ"
         Height          =   255
         Index           =   24
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›"
         Height          =   255
         Index           =   25
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   600
         Width           =   1995
      End
   End
   Begin VB.ComboBox cStatus 
      Height          =   315
      ItemData        =   "FrmCase.frx":000C
      Left            =   240
      List            =   "FrmCase.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox dcType 
      Height          =   315
      ItemData        =   "FrmCase.frx":002E
      Left            =   6600
      List            =   "FrmCase.frx":0038
      TabIndex        =   50
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox TxtnoOfInst 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3000
      Width           =   1965
   End
   Begin VB.TextBox txtinstallValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   2640
      Width           =   2325
   End
   Begin VB.TextBox TxtAccumulated 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   2280
      Width           =   1605
   End
   Begin VB.TextBox TxtAge 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2640
      Width           =   525
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·Â «Â·«ﬂ"
      Height          =   225
      Index           =   0
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·Ì” ·Â «Â·«ﬂ"
      Height          =   255
      Index           =   1
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox XPTxtID 
      Height          =   285
      Left            =   6960
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtRealValue 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2280
      Width           =   1965
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   2235
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4080
      Width           =   3405
   End
   Begin VB.TextBox TxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3990
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1215
      Width           =   4605
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6750
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1845
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   10155
      _cx             =   17912
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      Caption         =   "«·«’Ê· «·À«» …"
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
         Left            =   1155
         TabIndex        =   5
         Top             =   120
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
         ButtonImage     =   "FrmCase.frx":005B
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
         Left            =   90
         TabIndex        =   6
         Top             =   120
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
         ButtonImage     =   "FrmCase.frx":03F5
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
         Left            =   1680
         TabIndex        =   7
         Top             =   120
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
         ButtonImage     =   "FrmCase.frx":078F
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
         Left            =   615
         TabIndex        =   8
         Top             =   120
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
         ButtonImage     =   "FrmCase.frx":0B29
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
      Height          =   375
      Index           =   0
      Left            =   7350
      TabIndex        =   9
      Top             =   7035
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   6600
      TabIndex        =   10
      Top             =   7035
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   5715
      TabIndex        =   11
      Top             =   7035
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   4845
      TabIndex        =   12
      Top             =   7035
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   3000
      TabIndex        =   13
      Top             =   7020
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   2130
      TabIndex        =   14
      Top             =   7035
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   345
      Left            =   9960
      TabIndex        =   28
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92471297
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   7560
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtSttart 
      Height          =   345
      Left            =   240
      TabIndex        =   32
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92471297
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcEmployee 
      Height          =   315
      Left            =   3960
      TabIndex        =   46
      Top             =   1920
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   240
      TabIndex        =   48
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92471297
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   52
      Top             =   7035
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "’Ê—… «·«’·"
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
      Index           =   7
      Left            =   6600
      TabIndex        =   53
      Top             =   7440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«Ìﬁ«› «·«Â·«ﬂ"
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
      Index           =   8
      Left            =   5040
      TabIndex        =   54
      Top             =   7440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "≈⁄«œ…  ‘€Ì· «·«Â·«ﬂ"
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
      Index           =   9
      Left            =   3480
      TabIndex        =   55
      Top             =   7440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«· Œ·’ „‰ «·«’·"
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
   Begin MSDataListLib.DataCombo DCGroup 
      Height          =   315
      Left            =   3960
      TabIndex        =   56
      Top             =   1560
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   5880
      TabIndex        =   72
      Top             =   840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰›–"
      Height          =   255
      Index           =   26
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ »ﬁÏ"
      Height          =   255
      Index           =   4
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê’› «·«’·"
      Height          =   195
      Index           =   0
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   3840
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·«Â·«ﬂ"
      Height          =   255
      Index           =   18
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2760
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«” ·«„"
      Height          =   375
      Index           =   17
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»⁄ÂœÂ"
      Height          =   315
      Index           =   16
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   315
      Index           =   15
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ã„Ê⁄Â"
      Height          =   315
      Index           =   14
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «ﬁ”«ÿ   «·«Â·«ﬂ"
      Height          =   255
      Index           =   13
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ"
      Height          =   255
      Index           =   12
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
      Height          =   255
      Index           =   11
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·«Â·«ﬂ"
      Height          =   375
      Index           =   10
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—"
      Height          =   255
      Index           =   9
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2760
      Width           =   2115
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·«’·"
      Height          =   255
      Index           =   8
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   5
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7560
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   375
      Left            =   8280
      TabIndex        =   24
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LngDevID 
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
      Height          =   315
      Index           =   3
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   5010
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6510
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4710
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6480
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·«’·"
      Height          =   315
      Index           =   1
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·«’·"
      Height          =   315
      Index           =   2
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "FrmCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RSAss As New ADODB.Recordset

Dim TTP As clstooltip



Private Sub Cmd_Click(Index As Integer)
'On Error GoTo ErrTrap


Select Case Index
    Case 0
        TxtModFlg.text = "N"
        clear_all Me
 
        Me.DCboUserName.BoundText = user_id
   
    Case 1
        TxtModFlg.text = "E"
    Case 2
    
If DCPreFix.text = "" Then
MsgBox "Õœœ «·Ã“¡ «·À«Ì "
DCPreFix.SetFocus
SendKeys "{F4}"

Exit Sub
End If
Dim currentcode As String
If txtid.text = "" Then
currentcode = get_coding(branch_id, "FixedAssets", 1, Me.DCPreFix.text)
If currentcode = "miniError" Then
MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
Exit Sub
            
ElseIf currentcode = "Manual" Then
MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·”‰œ« "
Else
txtid = currentcode
End If
End If
Exit Sub


        SaveData
    Case 3
       Call Undo
    Case 4
       Del_AssetType
    Case 5
    VIEW_ATTACH
    
    Case 6
        Unload Me
End Select
Exit Sub
ErrTrap:
End Sub
Function VIEW_ATTACH()
'On Error Resume Next

 
'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

 

imaged.Show
imaged.Label9.Caption = "„—›ﬁ«  «·«’· —ﬁ„"
imaged.Caption = "„—›ﬁ«  «·«’·  "
imaged.txtopeation_type = "„—›ﬁ«  «·«’·"
imaged.SUBJECT_NO = 1 'TxtEmp_Code.text
imaged.Label6.Caption = "ﬂÊœ «·«’·"
imaged.Adodc1.CommandType = adCmdText
imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·«’·' and subject_no='" & "1" & "'"
imaged.Adodc1.Refresh
If imaged.Adodc1.Recordset.RecordCount > 0 Then

imaged.DBPix201.Visible = True
Else
imaged.DBPix201.Visible = False
End If
End Function
Private Sub DcboBox_Change()
 

End Sub

Private Sub cStatus_Click()
If cStatus.ListIndex > -1 Then
    If cStatus.ListIndex = 0 Then 'ÃœÌœ
    txtRealValue.Enabled = True
    If Opt(0).value = True Then
    TxtAccumulated.Enabled = True
     dcType.Enabled = True
     TxtAge.Enabled = True
     TxtnoOfInst.Enabled = True
     txtinstallValue.Enabled = True
     txtinstallDo.Enabled = True
    dtSttart.Enabled = True
     Else
     TxtAccumulated.Enabled = False
     dcType.Enabled = False
     TxtAge.Enabled = False
     TxtnoOfInst.Enabled = False
     txtinstallValue.Enabled = False
     txtinstallDo.Enabled = False
    dtSttart.Enabled = False
     
      TxtAccumulated.text = ""
     dcType.text = ""
     TxtAge.text = ""
     TxtnoOfInst.text = ""
     txtinstallValue.text = ""
     txtinstallDo.text = ""
    
    
    End If
    
    ElseIf cStatus.ListIndex = 1 Then    'Ã«—Ì «·«Â·«ﬂ
   If Opt(0).value = True Then
    
     TxtAccumulated.Enabled = True
     dcType.Enabled = True
     TxtAge.Enabled = True
     TxtnoOfInst.Enabled = True
     txtinstallValue.Enabled = True
     txtinstallDo.Enabled = True
    dtSttart.Enabled = True
 
   Else
    
       TxtAccumulated.Enabled = False
     dcType.Enabled = False
     TxtAge.Enabled = False
     TxtnoOfInst.Enabled = False
     txtinstallValue.Enabled = False
     txtinstallDo.Enabled = False
    dtSttart.Enabled = False
     
      TxtAccumulated.text = ""
     dcType.text = ""
     TxtAge.text = ""
     TxtnoOfInst.text = ""
     txtinstallValue.text = ""
     txtinstallDo.text = ""
     
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·Õ«·Â Ã«—Ì «·«Â·«ﬂ ›Ì ÕÌ‰ «‰ «·«’· ·Ì” ·Â «Â·«ﬂ"
    Else
         MsgBox "cahnge status"
    End If
     cStatus.SetFocus
     SendKeys "{F4}"
    End If
    End If

End If
End Sub

Private Sub DCGroup_Click(Area As Integer)
 If Val(Me.dcBranch.BoundText) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
 Else
 MsgBox "Select Branch Firstly    ", vbCritical
 End If
 dcBranch.SetFocus
 SendKeys "{F4}"
 End If
 Dim AccountName As String
Dim Percentage1 As Integer
Dim Percentage2 As Integer
 GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 24, Val(Me.dcBranch.BoundText), , AccountName 'Õ”«» «·«’·
  TXT24.text = AccountName
 GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 26, Val(Me.dcBranch.BoundText), , AccountName  '„’—Ê›«  «·«Â·«ﬂ
   TXT26.text = AccountName
 GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 25, Val(Me.dcBranch.BoundText), , AccountName '„Ã„⁄ «·«Â·«ﬂ
   TXT25.text = AccountName
  GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 31, Val(Me.dcBranch.BoundText), , AccountName '«—»«Õ »Ì⁄
    TXT31.text = AccountName
  GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , AccountName  'Œ”«∆— »Ì⁄
    TXT40.text = AccountName
 GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , , Percentage1   '  ‰”»… «·«Â·«ﬂ
 TXtPercentage1.text = Percentage1
 GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , , , Percentage2 '  ‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›
 txtPercentage2.text = Percentage2
 
End Sub

Private Sub Form_Activate()
'XPTxtID.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Cmd_Click (6)
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Load()
On Error GoTo ErrTrap
Dim Dcombos As New ClsDataCombos

'Dcombos.GetBoxes Me.DcboBox
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetFixedAssetsGroup DCGroup

Dcombos.GetPrefix Me.DCPreFix, 1, Val(branch_id)
 
Dim My_SQL As String


My_SQL = "  select branch_id,branch_name from branches   "
fill_combo dcBranch, My_SQL

My_SQL = "  select Emp_ID,Emp_name  from TblEmployee order by Emp_name   "
fill_combo DcEmployee, My_SQL


Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
'Dcombos.GetAccountingCodes Me.DcboCreditSide

Resize_Form Me
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If

AddTip
Set rs = New ADODB.Recordset
rs.Open "FixedAssets", Cn, adOpenStatic, adLockOptimistic, adCmdTable

Set RSAss = New ADODB.Recordset
Dim StrSQL As String

StrSQL = "select * From  Notes where NoteType='300' order by NoteID"
RSAss.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

Me.TxtModFlg.text = "R"
Retrive
If OPEN_NEW_SCREEN = True Then
Cmd_Click (0)
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
ErrTrap:
 End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
If rs.state = adStateOpen Then
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
     '   Me.Caption = "«·«’Ê· «·À«» …"
        Me.Cmd(2).Enabled = False
        Me.Cmd(3).Enabled = False
        
        Me.Cmd(0).Enabled = True
        Me.Cmd(1).Enabled = True
        Me.Cmd(4).Enabled = True
        
        Me.XPBtnMove(0).Enabled = True
        Me.XPBtnMove(1).Enabled = True
        Me.XPBtnMove(2).Enabled = True
        Me.XPBtnMove(3).Enabled = True
        
       
      
        Me.XPMTxtRemark.locked = True
        If rs.RecordCount < 1 Then
      '      Me.XPBtnMove(0).Enabled = False
      '      Me.XPBtnMove(1).Enabled = False
      '      Me.XPBtnMove(2).Enabled = False
      '      Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        End If
    Case "N"
     '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        
      
        Me.XPMTxtRemark.locked = False
    Case "E"
     '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        
       
        Me.XPMTxtRemark.locked = False
End Select
Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
On Error GoTo ErrTrap


If RSAss.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
 
'XPTxtID.text = RSAss("noteID")
XPDtbTrans.value = RSAss("noteDate")

Dim RsAssTmp As New ADODB.Recordset
Set RsAssTmp = New ADODB.Recordset
Dim StrSQL As String
StrSQL = "SELECT TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Code FROM dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code Where  (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = '0') and dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID =" & Me.XPTxtID.text
RsAssTmp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
RsAssTmp.Close

StrSQL = "SELECT TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Code FROM dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code Where  (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1) and dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID =" & Me.XPTxtID.text
RsAssTmp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'DcboBox.text = IIf(IsNull(RsAssTmp("Account_name").value), "", Val(RsAssTmp("Account_name").value))
RsAssTmp.Close

'XPTxtBankName.text = IIf(IsNull(RsAssTmp("Name").Value), "", Trim(RsAssTmp("Name").Value))
'XPMTxtRemark.text = IIf(IsNull(RSAss("Remarks").Value), "", Trim(RSAss("Remarks").Value))
XPTxtCurrent.Caption = RSAss.AbsolutePosition
XPTxtCount.Caption = RSAss.RecordCount
Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
'Set RSRet = New ADODB.Recordset
'Dim strsql As String
'
'strsql = "select * From  Notes where NoteType='300' order by NoteID"
'RSAss.Open strsql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Me.TxtModFlg.text = "N" Then
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
End If

Select Case Index
    Case 0
        If Not (RSAss.EOF Or RSAss.BOF) Then
            RSAss.MovePrevious
            If RSAss.BOF Then RSAss.MoveFirst
        End If
    Case 1
        If Not (RSAss.EOF Or RSAss.BOF) Then
            RSAss.MoveFirst
        End If
    Case 2
        If Not (RSAss.EOF Or RSAss.BOF) Then
            RSAss.MoveLast
        End If
    Case 3
        If Not (RSAss.EOF Or RSAss.BOF) Then
            RSAss.MoveNext
            If RSAss.EOF Then RSAss.MoveLast
        End If
End Select
Retrive
Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim Msg As String
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset
Dim RsTempM As New ADODB.Recordset
Dim RsDev As New ADODB.Recordset
Dim RsNot As New ADODB.Recordset

Dim BeginTrans As Boolean
'On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then

    
    If TxtName.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ «”„ «·«’·  «Ê·«", vbCritical
    Else
    MsgBox "Select Name Firstly    ", vbCritical
    End If
    
        TxtName.SetFocus
        Exit Sub
    End If
    
 If Val(Me.dcBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
    Else
    MsgBox "Select Branch Firstly    ", vbCritical
    End If
 dcBranch.SetFocus
 SendKeys "{F4}"
 End If
 
   If Val(Me.DCGroup.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ „Ã„Ê⁄Â «·«’·  «Ê·«", vbCritical
    Else
    MsgBox "Select Group Firstly    ", vbCritical
    End If
     DCGroup.SetFocus
     SendKeys "{F4}"
    End If
 
  If Me.cStatus.ListIndex = -1 Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ  Õ«·Â  «·«’·     ", vbCritical
    Else
    MsgBox "Select Status     ", vbCritical
    End If
    cStatus.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If

    
    
        If Me.DcEmployee.BoundText = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ   «·«’· »⁄ÂœÂ  «Ê·«", vbCritical
    Else
    MsgBox "Select Holder Name   ", vbCritical
    End If
    
        DcEmployee.SetFocus
         SendKeys "{F4}"
        Exit Sub
    End If
    
    
  If cStatus.ListIndex = 0 Then 'ÃœÌœ
   If Not IsNumeric(txtRealValue.text) Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ ﬁÌ„Â «·«’· «·œ› —Ì…", vbCritical
    Else
    MsgBox "Write Real value ", vbCritical
    End If
    
        txtRealValue.SetFocus
        Exit Sub
    End If
    If Opt(0).value = True Then
    If Me.dcType.ListIndex = -1 Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ     'ÿ—Ìﬁ… «·«Â·«ﬂ     ", vbCritical
    Else
    MsgBox "Specify DDD type       ", vbCritical
    End If
    dcType.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If


   If Not IsNumeric(TxtAge.text) Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ ⁄„— «·«’· «·«› —«÷Ì", vbCritical
    Else
    MsgBox "Write Default Age value ", vbCritical
    End If
    
        TxtAge.SetFocus
        Exit Sub
    End If
    
    
       If Not IsNumeric(TxtnoOfInst.text) Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ ⁄œœ «ﬁ”«ÿ «·«Â·«ﬂ", vbCritical
    Else
    MsgBox "Write No Of DDD Installments", vbCritical
    End If
    
        TxtnoOfInst.SetFocus
        Exit Sub
    End If
    
  If Not IsNumeric(txtinstallValue.text) Then
     If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õœœ ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ", vbCritical
    Else
    MsgBox "Write   DDD Installment value", vbCritical
    End If
    
        txtinstallValue.SetFocus
        Exit Sub
    End If
    
    
    
 
    
    
    End If
  
  ElseIf cStatus.ListIndex = 1 Or cStatus.ListIndex = 2 Or cStatus.ListIndex = 3 Then    '
   
  ElseIf cStatus.ListIndex = 4 Then     '
    
  End If
    
    Select Case Me.TxtModFlg.text
        Case "N"
                

            StrSQL = "select * From  AssetType where Name='" & Trim(XPTxtBankName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
                Msg = "Â‰«ﬂ ‰Ê⁄ «’Ê· „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBankName.SetFocus
                Exit Sub
            End If
        Case "E"
        
            StrSQL = "select * From  AssetType where Name='" & Trim(XPTxtBankName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
            If RsTemp("ID").value <> Val(XPTxtID.text) Then
                Msg = "Â‰«ﬂ ‰Ê⁄ „’—Ê›«  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBankName.SetFocus
                Exit Sub
            End If
            End If
    End Select
     If Me.TxtModFlg.text = "N" Then
   
        End If
    Cn.BeginTrans
    BeginTrans = True
    Select Case Me.TxtModFlg.text
        Case "N"
            rs.AddNew
            rs("ID").value = Val(XPTxtID.text)
            
            Set RsNot = New ADODB.Recordset
            RsNot.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            RsNot("NoteID") = XPTxtID.text
        Case "E"
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End Select
 End If
 
 
'**************************************************************************
    
    
    
    
    
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Select Case Me.TxtModFlg.text
        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‰Ê⁄" & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            Cmd_Click (0)
            Exit Sub
            End If
            
        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End Select
    TxtModFlg.text = "R"
 
Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
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
         rs.find "ID=" & Val(XPTxtID.text) & "", , adSearchForward, adBookmarkFirst
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
Private Sub Del_AssetType()
Dim Msg As String
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset

On Error GoTo ErrTrap
If XPTxtID.text <> "" Then
    StrSQL = "select * From Notes where ExpensesID=" & Trim(XPTxtID.text)
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Chr(13)
        Msg = Msg + "· ﬂ«„· «·»Ì«‰« "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Msg = "”Ì „ Õ–› »Ì«‰«  «·‰Ê⁄ —ﬁ„ " & Chr(13)
    Msg = Msg + (XPTxtID.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
        If Not rs.RecordCount < 1 Then
            Dim StrAccountCode As String
            StrAccountCode = rs("Account_Code").value
            If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                rs.Delete
            Else
                Exit Sub
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
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtModFlg_Change
    Exit Sub
End If
TxtModFlg_Change
Exit Sub
ErrTrap:
If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·‰Ê⁄ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
    rs.CancelUpdate
End If
End Sub
Private Sub AddTip()
Dim Wrap As String
On Error GoTo ErrTrap
Set TTP = New clstooltip
Wrap = Chr(13) + Chr(10)
If SystemOptions.UserInterface = ArabicInterface Then
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(0), _
            "ÃœÌœ ..." & Wrap & _
            "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(1), _
            " ⁄œÌ· ..." & Wrap & _
            "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(2), _
            "Õ›Ÿ ..." & Wrap & _
            "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & _
             "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(3), _
            " —«Ã⁄ ..." & Wrap & _
            "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
             "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
         With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(4), _
            "Õ–› ..." & Wrap & _
            "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(6), _
            "Œ—ÊÃ ..." & Wrap & _
            "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(1), _
            "«·√Ê· ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(0), _
            "«·”«»ﬁ ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(3), _
            "«· «·Ì ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(2), _
            "«·√ŒÌ— ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
ElseIf SystemOptions.UserInterface = EnglishInterface Then

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

Me.Caption = "Assets"
Me.Ele.Caption = Me.Caption
Me.lbl(2).Caption = "Code"
Me.lbl(14).Caption = "Group"
Opt(0).Caption = "opening Balance"
Opt(1).Caption = "Purchase"
Me.lbl(1).Caption = "Name"
Me.lbl(15).Caption = "Branch"
Me.lbl(16).Caption = "Employee"
Me.lbl(17).Caption = "Received"
Me.lbl(3).Caption = "Value"
Me.lbl(11).Caption = "Damage"
Me.lbl(11).Caption = "Depreciation"
Me.lbl(8).Caption = "Status"
Me.lbl(18).Caption = "Depreciation Type"
Me.lbl(9).Caption = "Lifespan"
Me.lbl(10).Caption = "Start Deprec"
Me.lbl(13).Caption = "no of installm."
Me.lbl(12).Caption = "installent Value."
Me.lbl(19).Caption = "Asset Account"
Me.lbl(20).Caption = "Accumulated depreciation Acc."
Me.lbl(21).Caption = "depreciation Acc. Expenses"
Me.lbl(22).Caption = "Payments Acc."
Me.lbl(5).Caption = "By"

 
 
Me.lbl(0).Caption = "Remark"
Cmd(7).Caption = "Stop Depreciation"
Cmd(8).Caption = "Depreciation Restart"
Cmd(9).Caption = "Asset Disposal"
Cmd(5).Caption = "Asset Image"


Me.lbl(7).Caption = "Current Record:"
Me.lbl(6).Caption = "Records NO:"
Me.Cmd(0).Caption = "New"
Me.Cmd(1).Caption = "Edit"
Me.Cmd(2).Caption = "Save"
Me.Cmd(3).Caption = "Undo"
Me.Cmd(4).Caption = "Delete"
Me.Cmd(6).Caption = "Exit"

End Sub







