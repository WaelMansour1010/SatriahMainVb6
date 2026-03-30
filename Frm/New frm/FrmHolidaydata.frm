VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmHolidayData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·«Ã«“… "
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   Icon            =   "FrmHolidaydata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   12465
   Begin VB.CommandButton menue 
      Height          =   270
      Index           =   2
      Left            =   6360
      Picture         =   "FrmHolidaydata.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   1185
      Width           =   375
   End
   Begin XtremeSuiteControls.RadioButton RdDaye 
      Height          =   255
      Left            =   1440
      TabIndex        =   132
      Top             =   1200
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "„œ… «·«Ã«“…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox TxtOrderQupemp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   131
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox noOfMonth 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox visano 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   115
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   1185
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtInterval 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   -150
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   735
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13140
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   12525
      _cx             =   22093
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
      Caption         =   "»Ì«‰«  «·«Ã«“… "
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
         ButtonImage     =   "FrmHolidaydata.frx":07EE
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
         ButtonImage     =   "FrmHolidaydata.frx":0B88
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
         ButtonImage     =   "FrmHolidaydata.frx":0F22
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
         ButtonImage     =   "FrmHolidaydata.frx":12BC
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
         Left            =   6000
         Picture         =   "FrmHolidaydata.frx":1656
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   37
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7740
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   93585409
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      Top             =   1185
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   2790
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6060
      Width           =   8745
      _cx             =   15425
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
         Left            =   7200
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
         Left            =   6375
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
         Left            =   5535
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
         Left            =   4680
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
         Left            =   3825
         TabIndex        =   15
         Top             =   75
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
         Left            =   0
         TabIndex        =   16
         Top             =   60
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
         Left            =   855
         TabIndex        =   17
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   2760
         TabIndex        =   30
         Top             =   60
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   42
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8580
      TabIndex        =   18
      Top             =   5640
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13200
      TabIndex        =   19
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   13560
      TabIndex        =   32
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmHolidaydata.frx":52BE
      Height          =   315
      Left            =   3840
      TabIndex        =   34
      Top             =   720
      Width           =   3135
      _ExtentX        =   5530
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   3735
      Left            =   0
      TabIndex        =   43
      Top             =   1560
      Width           =   12480
      _cx             =   22013
      _cy             =   6588
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmHolidaydata.frx":52D3
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3270
         Left            =   13125
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   5768
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   1590
            Left            =   -120
            TabIndex        =   45
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   2805
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
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmHolidaydata.frx":566D
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label111000 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3270
         Index           =   15
         Left            =   45
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   5768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   12
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
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
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmHolidaydata.frx":57B0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3240
            Index           =   16
            Left            =   15
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   5715
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
            Appearance      =   5
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
            Begin VB.CheckBox RegCK 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—”„Ì…"
               Height          =   255
               Index           =   0
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   2880
               Width           =   855
            End
            Begin VB.CheckBox RegCK 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷—«—Ì…"
               Height          =   255
               Index           =   1
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   2880
               Width           =   855
            End
            Begin MSComCtl2.DTPicker ExpectedReturndate 
               Height          =   360
               Left            =   8880
               TabIndex        =   128
               Top             =   2760
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   635
               _Version        =   393216
               Format          =   93585409
               CurrentDate     =   38784
            End
            Begin VB.TextBox txtRemark 
               Alignment       =   1  'Right Justify
               Height          =   2160
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   88
               Top             =   360
               Width           =   4755
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   13410
               MaxLength       =   10
               TabIndex        =   86
               Top             =   2100
               Width           =   1425
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   3765
               Index           =   0
               Left            =   14145
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   360
               Width           =   6135
               Begin VB.TextBox TxtPaymentCounts 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4110
                  MaxLength       =   2
                  TabIndex        =   79
                  Top             =   240
                  Width           =   825
               End
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   78
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.CheckBox ChkSaleryDis 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   77
                  Top             =   2160
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   76
                  Top             =   1320
                  Width           =   1095
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   75
                  Top             =   1680
                  Width           =   1965
                  _ExtentX        =   3466
                  _ExtentY        =   767
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈Õ”»  Ê«—ÌŒ «·”œ«œ"
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
                  ButtonImage     =   "FrmHolidaydata.frx":57E4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2325
                  Left            =   90
                  TabIndex        =   80
                  Top             =   210
                  Width           =   3855
                  _cx             =   6800
                  _cy             =   4101
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
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmHolidaydata.frx":5B7E
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
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·œ›⁄« "
                  Height          =   285
                  Index           =   9
                  Left            =   4830
                  TabIndex        =   85
                  Top             =   300
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Ê· œ›⁄…"
                  Height          =   285
                  Index           =   10
                  Left            =   4380
                  TabIndex        =   84
                  Top             =   690
                  Width           =   1665
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
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
                  Height          =   255
                  Left            =   60
                  TabIndex        =   83
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   11
                  Left            =   5250
                  TabIndex        =   82
                  Top             =   990
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   12
                  Left            =   5250
                  TabIndex        =   81
                  Top             =   1320
                  Width           =   405
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1005
               Left            =   14760
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   0
               Width           =   6015
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   65
                  Top             =   360
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
                  Height          =   285
                  Index           =   17
                  Left            =   3960
                  TabIndex        =   73
                  Top             =   720
                  Width           =   1965
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
                  Height          =   285
                  Index           =   18
                  Left            =   1560
                  TabIndex        =   72
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·› ·„  ”œœ"
                  Height          =   285
                  Index           =   19
                  Left            =   1800
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   285
                  Index           =   16
                  Left            =   -240
                  TabIndex        =   70
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   20
                  Left            =   960
                  TabIndex        =   69
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   21
                  Left            =   960
                  TabIndex        =   68
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   22
                  Left            =   3240
                  TabIndex        =   67
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   14
                  Left            =   4800
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1125
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   1905
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   0
               Width           =   6105
               Begin VB.TextBox TxtNumPasp 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   600
                  Width           =   1995
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.TextBox TxtSearchCode1 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.TextBox TxtNumEkama 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   600
                  Width           =   1935
               End
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBIssueDate 
                  Height          =   315
                  Left            =   6480
                  TabIndex        =   58
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcemplocation 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   108
                  Top             =   960
                  Width           =   3435
                  _ExtentX        =   6059
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcemplocation1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   109
                  Top             =   2400
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboMangerName 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   119
                  Top             =   1320
                  Width           =   3435
                  _ExtentX        =   6059
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ÃÊ«“"
                  Height          =   285
                  Index           =   44
                  Left            =   2280
                  TabIndex        =   123
                  Top             =   600
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„œÌ— «·„»«‘—"
                  Height          =   315
                  Index           =   40
                  Left            =   4860
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   1320
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«ﬁ«„…"
                  Height          =   285
                  Index           =   41
                  Left            =   5280
                  TabIndex        =   113
                  Top             =   600
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„· «·Õ«·Ì"
                  Height          =   405
                  Index           =   38
                  Left            =   1920
                  TabIndex        =   110
                  Top             =   2400
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„· "
                  Height          =   405
                  Index           =   37
                  Left            =   4920
                  TabIndex        =   107
                  Top             =   960
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   6600
                  TabIndex        =   63
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   6360
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   285
                  Index           =   15
                  Left            =   2280
                  TabIndex        =   61
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Index           =   24
                  Left            =   5400
                  TabIndex        =   60
                  Top             =   240
                  Width           =   645
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   510
               Left            =   240
               TabIndex        =   90
               Top             =   2595
               Visible         =   0   'False
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
            Begin MSComCtl2.DTPicker TxtExpectedouttime 
               Height          =   315
               Left            =   8880
               TabIndex        =   95
               Top             =   4005
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   93585411
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtExpectedIntime 
               Height          =   375
               Left            =   8880
               TabIndex        =   96
               Top             =   4305
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   93585411
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualouttime 
               Height          =   315
               Left            =   5760
               TabIndex        =   99
               Top             =   4005
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   93585411
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualIntime 
               Height          =   375
               Left            =   5760
               TabIndex        =   100
               Top             =   4305
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   93585411
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker Returnbeforedate 
               Height          =   360
               Left            =   8880
               TabIndex        =   111
               Top             =   2040
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   635
               _Version        =   393216
               Format          =   93585409
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal ReturnbeforedateH 
               Height          =   315
               Left            =   7680
               TabIndex        =   112
               Top             =   2040
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker DeparDate 
               Height          =   360
               Left            =   8880
               TabIndex        =   125
               Top             =   2400
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   635
               _Version        =   393216
               Format          =   93585409
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal DeparDateH 
               Height          =   315
               Left            =   7680
               TabIndex        =   126
               Top             =   2400
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal ExpectedReturndateH 
               Height          =   315
               Left            =   7680
               TabIndex        =   129
               Top             =   2760
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
            End
            Begin XtremeSuiteControls.RadioButton RdRetBefore 
               Height          =   255
               Left            =   10320
               TabIndex        =   133
               Top             =   2040
               Width           =   1905
               _Version        =   786432
               _ExtentX        =   3360
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„Õœœ «·⁄Êœ… ﬁ»·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·⁄Êœ… «·„ Êﬁ⁄"
               Height          =   315
               Index           =   46
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   2760
               Width           =   1905
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·”›— «·„ Êﬁ⁄"
               Height          =   315
               Index           =   45
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   2400
               Width           =   1905
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·›⁄·Ì"
               Height          =   255
               Index           =   35
               Left            =   7320
               TabIndex        =   102
               Top             =   4320
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·›⁄·Ì"
               Height          =   210
               Index           =   34
               Left            =   7320
               TabIndex        =   101
               Top             =   4080
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·„ Êﬁ⁄"
               Height          =   255
               Index           =   32
               Left            =   10680
               TabIndex        =   94
               Top             =   4305
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·„ Êﬁ⁄"
               Height          =   210
               Index           =   31
               Left            =   10680
               TabIndex        =   93
               Top             =   4005
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   330
               Index           =   28
               Left            =   5160
               TabIndex        =   89
               Top             =   345
               Width           =   840
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
               Height          =   330
               Index           =   26
               Left            =   12045
               TabIndex        =   87
               Top             =   1425
               Width           =   2280
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2520
               Index           =   62
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1155
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3240
            Index           =   9
            Left            =   15
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   5715
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   2430
               Left            =   3225
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   675
               Width           =   660
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   1815
               Left            =   4065
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   885
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1815
               Index           =   67
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   885
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1620
               Index           =   68
               Left            =   3885
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   1080
               Width           =   30
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1905
               Index           =   69
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   885
               Width           =   315
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DcOutType 
      Height          =   315
      Left            =   3720
      TabIndex        =   98
      Top             =   240
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker Indate 
      Height          =   360
      Left            =   1560
      TabIndex        =   104
      Top             =   240
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      _Version        =   393216
      Format          =   93585409
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal indateH 
      Height          =   315
      Left            =   120
      TabIndex        =   105
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«” »Ì«‰ —ﬁ„"
      Height          =   435
      Index           =   47
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   130
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ÌÊ„"
      Height          =   435
      Index           =   43
      Left            =   -480
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «· √‘Ì—…"
      Height          =   435
      Index           =   36
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·«–‰"
      Height          =   285
      Index           =   33
      Left            =   6120
      TabIndex        =   97
      Top             =   360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”«⁄Â"
      Height          =   285
      Index           =   29
      Left            =   960
      TabIndex        =   92
      Top             =   -120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·» ”›… ‰ﬁœÌ… ÊÌ „ «Õ ”«» ﬁÌ„… «·œ›⁄ «·Ì«"
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
      Height          =   660
      Index           =   25
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3330
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Index           =   39
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   780
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì… "
      Height          =   285
      Index           =   4
      Left            =   11430
      TabIndex        =   29
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   11430
      TabIndex        =   28
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„œ…"
      Height          =   285
      Index           =   2
      Left            =   3510
      TabIndex        =   27
      Top             =   -15
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   8670
      TabIndex        =   26
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11325
      TabIndex        =   25
      Top             =   5595
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2550
      TabIndex        =   24
      Top             =   5670
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
      Top             =   5670
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   22
      Top             =   5700
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1860
      TabIndex        =   21
      Top             =   5700
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   20
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmHolidayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
 Dim ExpectedReturndatea As Date

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(Me.xptxtid.Text))
End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
               ExpectedReturndatea = ExpectedReturndate.value
            Me.DCboUserName.BoundText = user_id
            TxtPaymentCounts.Text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               
             
    Grid2.Clear flexClearScrollable, flexClearEverything
    Grid2.Rows = 1
    
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
ExpectedReturndatea = ExpectedReturndate.value
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Load FrmSearchDataVocation
            FrmSearchDataVocation.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
            CalCulateParts
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.xptxtid.Text) <> 0 Then
                print_report val(Me.xptxtid.Text)
        
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
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
 
   
   MySQL = "SELECT     dbo.tblHolidayData.branch_no, dbo.tblHolidayData.Emp_ID, dbo.tblHolidayData.DeparmentID, dbo.tblHolidayData.recorddate, dbo.tblHolidayData.JobTypeID, "
   MySQL = MySQL & "  dbo.tblHolidayData.visano, dbo.tblHolidayData.noofmonth, dbo.tblHolidayData.Returnbeforedate, dbo.tblHolidayData.ReturnbeforedateH,"
   MySQL = MySQL & " dbo.tblHolidayData.DeparDate, dbo.tblHolidayData.DeparDateH, dbo.tblHolidayData.ExpectedReturndate, dbo.tblHolidayData.ExpectedReturndateH,"
   MySQL = MySQL & " dbo.tblHolidayData.mangerid, dbo.tblHolidayData.UserID, dbo.tblHolidayData.Remark, dbo.tblHolidayData.id, dbo.TblEmployee.Emp_Code,"
   MySQL = MySQL & " dbo.TblEmployee.Emp_Name, dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPasp, dbo.EmpGroupDep.GroupName, dbo.TblEmployee.Fullcode,"
    MySQL = MySQL & " dbo.tblHolidayData.chekdate , dbo.tblHolidayData.PostedDate, dbo.tblHolidayData.Posted"
  MySQL = MySQL & "  FROM   dbo.EmpGroupDep RIGHT OUTER JOIN"
  MySQL = MySQL & " dbo.TblEmployee ON dbo.EmpGroupDep.GroupID = dbo.TblEmployee.GroupID RIGHT OUTER JOIN"
   MySQL = MySQL & "  dbo.tblHolidayData ON dbo.TblEmployee.Emp_ID = dbo.tblHolidayData.Emp_ID"
 MySQL = MySQL & "  WHERE     (dbo.tblHolidayData.id =" & val(xptxtid.Text) & ")"
  

 
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "HolidayData.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "HolidayData.rpt"
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
     
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
 
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboMangerName_Change()
DcboMangerName_Click (0)
End Sub

Private Sub DcboMangerName_Click(Area As Integer)
       If val(DcboMangerName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboMangerName.BoundText, EmpCode
    TxtSearchCode1.Text = EmpCode
End Sub

Private Sub DeparDate_Change()
       If Me.TxtModFlg.Text <> "R" Then
             
                  DeparDateH.value = ToHijriDate(DeparDate.value)
               
        End If
End Sub

Private Sub ExpectedReturndate_Change()
     If Me.TxtModFlg.Text <> "R" Then
             
               ExpectedReturndateH.value = ToHijriDate(ExpectedReturndate.value)
               
        End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub Indate_Change()
        If Me.TxtModFlg.Text <> "R" Then
             
                  indateH.value = ToHijriDate(Indate.value)
               
        End If
End Sub

Private Sub menue_Click(Index As Integer)
If Index = 2 Then
  Load FrmEmployee
            FrmEmployee.show
End If
End Sub

Private Sub noOfMonth_Change()
If Me.noOfMonth.Text <> "" Then
ExpectedReturndate.value = DateAdd("d", val(Me.noOfMonth.Text), ExpectedReturndatea)
End If
End Sub

Private Sub noOfMonth_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, noOfMonth.Text, 0)
End Sub



Private Sub RdDaye_Click()
If Me.RdDaye.value = True Then
Me.noOfMonth.Enabled = True
ExpectedReturndate.Enabled = False
ExpectedReturndateH.Enabled = False
Me.RdRetBefore.value = False
End If
End Sub

Private Sub RdRetBefore_Click()
If Me.RdRetBefore.value = True Then
Me.noOfMonth.Enabled = False
noOfMonth.Text = ""
ExpectedReturndate.Enabled = True
ExpectedReturndateH.Enabled = True
Me.RdDaye.value = False
End If
End Sub

Private Sub Returnbeforedate_Change()
        If Me.TxtModFlg.Text <> "R" Then
             
                 ReturnbeforedateH.value = ToHijriDate(Returnbeforedate.value)
               
        End If
End Sub

Sub returnQuepemp()
Dim SpecificHolidyaType1 As Boolean
Dim SpecificHolidyaType2 As Boolean
Dim EmpID As Integer
GetQuepEmpVocation , EmpID, val(Me.TxtOrderQupemp.Text), , , , , , , , , , , , SpecificHolidyaType1, SpecificHolidyaType2
Me.DcboEmpName.BoundText = EmpID
If SpecificHolidyaType1 = True Then
  RegCK(0).value = vbChecked
       RegCK(1).value = vbUnchecked
    End If
If SpecificHolidyaType2 = True Then
  RegCK(1).value = vbChecked
       RegCK(0).value = vbUnchecked
       
       End If
End Sub





Private Sub TxtOrderQupemp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
returnQuepemp

End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
        
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 7
     '   Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    
   'If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

 
        
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_Code  As String
        Dim Balance As String
        Dim endContractPerMonth As Double
        Dim mangerid As Integer
        Dim GroupID As Integer
    Dim NumEkama As String
Dim NumPasp  As String
Dim visano As String

 Dim swapedempid As Integer
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, , mangerid, swapedempid, GroupID, NumPasp, NumEkama, , , , , , , , , , visano
' Me.visano.text = visano
        DBIssueDate.value = IssueDate
        DcboEmpDepartments.BoundText = DepID
        DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeID
       ' lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        DcboMangerName.BoundText = mangerid
       dcemplocation.BoundText = GroupID
       TxtNumEkama.Text = NumEkama
       TxtNumPasp.Text = NumPasp

    'End If

End Sub

 

Private Sub XPDtbTrans_Change()

    If Trim(txtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = txtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    txtNoteSerial1.Text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.Text = ""
    txtNoteSerial1.Text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
   
  
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
     '  Dcombos.GetOutType Me.DcOutType
       
    Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetEmployees Me.DcboMangerName
     
    Dcombos.GetBranches Me.Dcbranch

    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetEmpJobsTypes Me.DcboJobsType

    Dcombos.GetEmpLocations Me.dcemplocation
    Dcombos.GetEmpLocations Me.dcemplocation1
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblHolidayData     Order By id"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
    Retrive


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
    Label1.Visible = False
Me.RdDaye.Caption = "Long Vocation"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
RegCK(0).Caption = "Official"
RegCK(1).Caption = "COMPELLING"
    Me.Caption = "Data Of Vocation"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lbl(39).Caption = "Branch"
    lbl(47).Caption = "Ques No"
    lbl(36).Caption = "NoVisa"
    RdDaye.RightToLeft = False
    RdDaye.Caption = "Long Vocation"
     lbl(43).Caption = "Day"
     Frame1.Caption = "Data of Emp"
    lbl(24).Caption = "Job"
    lbl(15).Caption = "Manag"
    lbl(41).Caption = "IqamaNo"
    lbl(44).Caption = "PasNo"
    lbl(37).Caption = "Location"
    lbl(40).Caption = "ManaDirect"
    lbl(28).Caption = "Remarks"
    RdRetBefore.RightToLeft = False
    RdRetBefore.Caption = "Set bak Before"
    lbl(45).Caption = "Expected Date Travel"
    lbl(46).Caption = "Expected Date Return"
    XPTab301.Caption = "Data"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

'    With Me.Fg
'        .TextMatrix(0, .ColIndex("PartNO")) = "NO"
'        .TextMatrix(0, .ColIndex("PartValue")) = "Value"
'        .TextMatrix(0, .ColIndex("PartDate")) = "Date"

'    End With

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

Private Sub TxtInterval_LostFocus()
    Dim StrSQL As String
    Dim Mytot As String
    Dim MySal As String
    Exit Sub
    Dim Myrs As New ADODB.Recordset
    'StrSQL =
    Myrs.Open "SELECT * From TblEmployee  where Emp_ID=" & val(DcboEmpName.BoundText), Cn, adOpenStatic, adLockReadOnly

    If Not Myrs.EOF And Not IsNull(Myrs!Emp_Salary) Then
        MySal = Myrs!Emp_Salary
        Mytot = val(MySal) * 5

        If val(TxtInterval.Text) >= Mytot Then
            MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & CHR(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.title
            Exit Sub
   
        End If
  
    End If
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '        Me.Caption = "»Ì«‰«  «·«Ã«“… "
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
            TxtInterval.locked = True
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
            '        Me.Caption = "»Ì«‰«  «·«Ã«“… ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
            TxtInterval.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "»Ì«‰«  «·«Ã«“… (  ⁄œÌ· )"
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
            TxtInterval.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPaymentCounts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtPaymentCounts.Text, 1)
End Sub

Private Sub TxtPaymentCounts_LostFocus()

    If val(TxtPaymentCounts.Text) > 84 Then
        MsgBox "«·œ›«⁄  «ﬂ»— „‰ «·Õœ ", vbOKOnly, App.title
        Exit Sub
    End If
 
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

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
'''
    xptxtid.Text = IIf(IsNull(rs("id").value), "", val(rs("id").value))
    XPDtbTrans.value = IIf(IsNull(rs("recorddate").value), Date, rs("recorddate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.TxtNumEkama.Text = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
  
  
  Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
  
     visano.Text = IIf(IsNull(rs("visano").value), "", rs("visano").value)
      noOfMonth.Text = IIf(IsNull(rs("noOfMonth").value), 0, rs("noOfMonth").value)
      
 

        DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

   
  DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
 dcemplocation.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
 
     
Returnbeforedate.value = IIf(IsNull(rs("Returnbeforedate").value), Date, rs("Returnbeforedate").value)
ReturnbeforedateH.value = IIf(IsNull(rs("ReturnbeforedateH").value), ToHijriDate(Date), rs("ReturnbeforedateH").value)
 
 
 DeparDate.value = IIf(IsNull(rs("DeparDate").value), Date, rs("DeparDate").value)
DeparDateH.value = IIf(IsNull(rs("DeparDateH").value), ToHijriDate(Date), rs("DeparDateH").value)
 
 
ExpectedReturndate.value = IIf(IsNull(rs("ExpectedReturndate").value), Date, rs("ExpectedReturndate").value)
ExpectedReturndateH.value = IIf(IsNull(rs("ExpectedReturndateH").value), ToHijriDate(Date), rs("ExpectedReturndateH").value)
 
  
  
  DcboMangerName.BoundText = IIf(IsNull(rs("mangerid").value), "", rs("mangerid").value)


txtRemark.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)

  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
   If rs("chekdate").value = True Then
   Me.RdRetBefore.value = True
    Else
Me.RdDaye.value = True
    End If
      If rs("SpecificHolidyaType1").value = True Then
     RegCK(1).value = vbChecked
       RegCK(0).value = vbUnchecked
   Else
   RegCK(1).value = vbUnchecked
   RegCK(0).value = vbChecked
   End If
   
  
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
       If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
 
    
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String
Dim Sql As String
    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If


       ' If Me.visano.text = "" Then
       '     Msg = "ÌÃ»  ÕœÌœ —ﬁ„ «· √”Ì—…  ..!! "
       '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
       '     visano.SetFocus
             
       '     Exit Sub
       ' End If
        
 If Me.RdDaye.value = True Then
        If Me.noOfMonth.Text = "" Then
            Msg = "ÌÃ»  ÕœÌœ „œ… «·«Ã«“…  ..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            noOfMonth.SetFocus
             
            Exit Sub
        End If
    End If
         
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            xptxtid.Text = CStr(new_id("tblHolidayData", "id", "", True))
   
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
   
            StrSQL = "Delete From TblEmpHolidaysDetails Where IDHoli=" & val(Me.xptxtid.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If

               
      
          
    
Sql = "update TblEmployee set   jopstatusid =2  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute Sql
 Sql = "update TblEmployee set   lastHolidaydate =  " & SQLDate(DeparDate.value, True) & "   where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute Sql
   Sql = "update TblEmployee set   lastHolidaydateH =' " & DeparDateH.value & " '  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute Sql
        rs("branch_no").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
 
        rs("id").value = val(xptxtid.Text)
        rs("recorddate").value = XPDtbTrans.value
        rs("Emp_ID").value = Me.DcboEmpName.BoundText
        
     rs("DeparmentID").value = val(Me.DcboEmpDepartments.BoundText)
   rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
     rs("visano").value = IIf(visano.Text = "", Null, (visano.Text))
   rs("noOfMonth").value = IIf(noOfMonth.Text = "", Null, val(noOfMonth.Text))
   
   rs("NumEkama").value = IIf(TxtNumEkama.Text = "", Null, (TxtNumEkama.Text))
   
      rs("Returnbeforedate").value = Returnbeforedate.value
   
   rs("ReturnbeforedateH").value = ReturnbeforedateH.value
   
   

   rs("DeparDate").value = DeparDate.value
   
   rs("DeparDateH").value = DeparDateH.value
   
   rs("ExpectedReturndate").value = ExpectedReturndate.value
   
   rs("ExpectedReturndateH").value = ExpectedReturndateH.value
   rs("ProjectID").value = val(Me.dcemplocation.BoundText)
   
        rs("mangerid").value = val(Me.DcboMangerName.BoundText)
    rs("Remark").value = IIf(txtRemark.Text = "", Null, (txtRemark.Text))
    If Me.RdRetBefore.value = True Then
    rs("chekdate").value = 1
    Else
    rs("chekdate").value = 0
    End If
      If RegCK(0).value = vbChecked Then
      rs("SpecificHolidyaType1").value = 0
   Else
   If RegCK(1).value = vbChecked Then
   rs("SpecificHolidyaType1").value = 1
   End If
   End If
     
   
   
   rs("UserID").value = Me.DCboUserName.BoundText
   
 
   
         rs.update
 
        Set RsDetails = New ADODB.Recordset
        'RsDetails.Open "TblEmpHolidaysDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     * from dbo.TblEmpHolidaysDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    


        RsDetails.AddNew
                  If RegCK(1).value = vbChecked Then
      
                RsDetails("SpecificHolidyaType1").value = 1
                Else
                RsDetails("SpecificHolidyaType1").value = 0
                End If
                        RsDetails("IDHoli").value = val(xptxtid.Text)
                             RsDetails("Emp_ID").value = val(Me.DcboEmpName.BoundText)
                     RsDetails("fromdate").value = DeparDate.value
                     
                       RsDetails("DateExpectedM").value = ExpectedReturndate.value
                       
                       RsDetails("fromdateh").value = DeparDateH.value
                     
                       RsDetails("DateExpectedH").value = ExpectedReturndateH.value
                       
                        RsDetails("des").value = txtRemark.Text
                       'RsDetails("cheke").value = 0
                    '   RsDetails("day").value = .TextMatrix(i, .ColIndex("day"))
                    '   RsDetails("month").value = .TextMatrix(i, .ColIndex("month"))
                    '   RsDetails("year").value = .TextMatrix(i, .ColIndex("year"))
                    RsDetails.update
                        
RsDetails.Close
Set RsDetails = Nothing
    
        Cn.CommitTrans
        BeginTrans = False
    
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "iD='" & val(xptxtid.Text) & "'", , adSearchForward, adBookmarkFirst

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

    On Error GoTo ErrTrap

    If xptxtid.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.xptxtid.Text)
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim Sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    Sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  Sql = Sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  Sql = Sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  Sql = Sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  Sql = Sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
Sql = Sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
Sql = Sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.xptxtid.Text)
                   RSApproval("NoteSerial").value = val(Me.xptxtid.Text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.xptxtid.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
                 If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
                Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
                Else
                 Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
                 End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.Rows = 1
    End If
RsDetails.Close

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
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·«Ã«“… ", 1, 15204351, -2147483630
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

Private Sub TxtInterval_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtInterval.Text, 0)
End Sub

Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String

    If year(Date) > val(Me.CboYear.Text) Then ' ⁄«„ „÷Ï
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckDate = False
        Exit Function
    ElseIf year(Date) = val(Me.CboYear.Text) Then '‰›” «·⁄«„

        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
            'Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ...!!!"
            'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            'CheckDate = False
            'Exit Function
        End If
    End If

    CheckDate = True
End Function

Private Function CheckPartCal() As Boolean
    Dim Msg As String

    CheckPartCal = False

    If val(TxtInterval.Text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtInterval.SetFocus
        Exit Function
    End If

    If val(TxtPaymentCounts.Text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtInterval.SetFocus
        Exit Function
    End If

    If CmbMonth.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ‘Â— · ”œÌœ «·œ›⁄…..!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CmbMonth.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If

    If CboYear.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ”‰… · ”œÌœ «·œ›⁄… ..!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboYear.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If

    CheckPartCal = True
End Function

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntPartCounts As Integer
    Dim SngPartValue As Single
    Dim m_FirstDate As Date

    If CheckPartCal = False Then
        Exit Sub
    End If

    If CheckDate = False Then
        Exit Sub
    End If

    SngPartValue = val(Me.TxtInterval.Text) / val(Me.TxtPaymentCounts.Text)
    IntPartCounts = val(Me.TxtPaymentCounts.Text)
    m_FirstDate = CDate("1-" & Me.CmbMonth.ListIndex + 1 & "-" & val(Me.CboYear.Text))

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + IntPartCounts
        .RowHeightMin = 300

        For i = 1 To IntPartCounts
            .TextMatrix(i, .ColIndex("PartNO")) = i
            .TextMatrix(i, .ColIndex("PartValue")) = SngPartValue
            .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
        Next i
    
    End With

End Sub

