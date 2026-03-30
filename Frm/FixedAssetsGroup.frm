VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{E1BFA30F-D929-4F80-AEDD-76FC2BDF5E23}#1.0#0"; "ciaXPPopUp30.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FixedAssetsGroup 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  „Ã„Ê⁄«  «·«’Ê·"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   HelpContextID   =   200
   Icon            =   "FixedAssetsGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   12060
   Begin VB.CheckBox chkIsContainer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„Ê⁄… Õ«ÊÌ« "
      Height          =   195
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   720
      Width           =   2775
      Begin XtremeSuiteControls.RadioButton RdMove 
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   58
         Top             =   120
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„‰ﬁÊ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdMove 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   59
         Top             =   120
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "€Ì— „‰ﬁÊ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3830
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   1125
   End
   Begin VB.TextBox XPTxtNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   60
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   4875
   End
   Begin VB.CheckBox DepType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·Â« «Â·«ﬂ"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox TxtPercentage2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2820
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   2640
      Width           =   2115
   End
   Begin VB.TextBox TxtPercentage1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2820
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«‰‘«¡ «·Õ”«»« "
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Chklast 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   2295
      Left            =   1980
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4320
      Width           =   4065
      _cx             =   7170
      _cy             =   4048
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
      ForeColor       =   4210752
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FixedAssetsGroup.frx":038A
      Caption         =   "„⁄·Ê„«  ≈÷«›Ì… ⁄‰ «·„Ã„Ê⁄…"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   15
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   14
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1470
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·„Ã„Ê⁄«  «·›—⁄Ì… «· Ï  Õ ÊÌÂ« «·„Ã„Ê⁄…"
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
         Height          =   225
         Index           =   4
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   330
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«’Ê· «· Ï  Õ ÊÌÂ« «·„Ã„Ê⁄…"
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
         Height          =   225
         Index           =   7
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   810
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ê· «’· „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
         Height          =   225
         Index           =   8
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1230
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ— «’· „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
         Height          =   225
         Index           =   9
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1710
         Width           =   2175
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   13
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1950
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   12
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   11
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   10
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   570
         Width           =   615
      End
   End
   Begin VB.TextBox TxtGroupCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox XPTxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   60
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1245
      Width           =   4875
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6060
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox TxtCutKey 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   90
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   150
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtMenuState 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Text            =   "N"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1185
      TabIndex        =   14
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FixedAssetsGroup.frx":0724
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
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FixedAssetsGroup.frx":0ABE
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
      Left            =   1710
      TabIndex        =   16
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FixedAssetsGroup.frx":0E58
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
      Left            =   645
      TabIndex        =   17
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FixedAssetsGroup.frx":11F2
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ciaXPPopMenu30.XPPopUp30 XPPopUp 
      Left            =   150
      Top             =   5310
      _ExtentX        =   900
      _ExtentY        =   873
      VisualStyle     =   0
      BeginProperty DefaultMenuItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuItemSpacing =   0
   End
   Begin MSDataListLib.DataCombo XPCboGroup 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1965
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.TreeView TreeGroups 
      Height          =   5685
      Left            =   6480
      TabIndex        =   21
      Top             =   750
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   10028
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   8475
      TabIndex        =   29
      Top             =   7560
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   7725
      TabIndex        =   30
      Top             =   7560
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   7005
      TabIndex        =   10
      Top             =   7560
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   6285
      TabIndex        =   31
      Top             =   7560
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   5355
      TabIndex        =   32
      Top             =   7560
      Width           =   915
      _ExtentX        =   1614
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
      Index           =   5
      Left            =   4500
      TabIndex        =   33
      Top             =   7560
      Width           =   825
      _ExtentX        =   1455
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
      Index           =   6
      Left            =   1950
      TabIndex        =   34
      Top             =   7560
      Width           =   705
      _ExtentX        =   1244
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   3660
      TabIndex        =   35
      Top             =   7560
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   2670
      TabIndex        =   36
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   2880
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DboParentAccount 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DboParentAccount1 
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ”«» «·«’·"
      Height          =   315
      Index           =   21
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Tag             =   "Ì „  ›⁄Ì·Â« ·«‰‘«¡ Õ”«»«  «·„Ã„Ê⁄Â «·Ì« ›Ì œ·Ì· «·Õ”«»« "
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ”«» „’—Ê› «·«Â·«ﬂ"
      Height          =   315
      Index           =   20
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Tag             =   "Ì „  ›⁄Ì·Â« ·«‰‘«¡ Õ”«»«  «·„Ã„Ê⁄Â «·Ì« ›Ì œ·Ì· «·Õ”«»« "
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   19
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›"
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   18
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·«Â·«ﬂ"
      Height          =   315
      Index           =   17
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
      Height          =   315
      Index           =   16
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Tag             =   "Ì „  ›⁄Ì·Â« ·«‰‘«¡ Õ”«»«  «·„Ã„Ê⁄Â «·Ì« ›Ì œ·Ì· «·Õ”«»« "
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
      Height          =   315
      Index           =   3
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   1
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7080
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   2
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
      Height          =   315
      Index           =   0
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1965
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   315
      Index           =   5
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1245
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„"
      Height          =   315
      Index           =   6
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7110
      Width           =   765
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "  »Ì«‰«  „Ã„Ê⁄«  «·«’Ê· "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   585
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   11955
   End
End
Attribute VB_Name = "FixedAssetsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim GroupReport As ClsGroupReport
Dim cSearchDcbo As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
   On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            DepType.value = vbChecked
            '        XPTxtName.SetFocus
            XPTxtID.Text = CStr(new_id("FixedAssetsGroup", "GroupID", "", True))
            txtPercentage2.Text = 0


            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(25, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ·„’—Ê›«  «·«Â·«ﬂ ··«’Ê·   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            
            
            
                Dim Account_Code_dynamic1 As String
            Account_Code_dynamic1 = get_account_code_branch(24, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ·„’—Ê›«  «·«Â·«ﬂ ··«’Ê·   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount1.BoundText = Account_Code_dynamic1
               
               
               TxtId.SetFocus
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            CuurentLogdata

        Case 2

            If TxtId = "" Then
                MsgBox "«œŒ· ﬂÊœ «·„Ã„Ê⁄Â     "
                Exit Sub
                'Else
                'txtid = CurrentCode
            End If

            If DepType.value = vbUnchecked Then
                TXtPercentage1.Text = 0
                txtPercentage2.Text = 0
            End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Group

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmGroupSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub Command1_Click()
    'Â–… «·œ«·…  ﬁÊ„ »Õ–› ﬂ· Õ”«»«  «·„Ã„Ê⁄«  »‘—ÿ ⁄œ„ «‘ —«ﬂÂ« ›Ì ﬁÌÊœ Ê«‰‘«¡ Õ”«»«  ÃœÌœ… ·ﬂ· «·„Ã„Ê⁄«  ÿ»ﬁ« ·Õ”«»«  «·›—⁄ «·ÃœÌœ…
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTemp2 As New ADODB.Recordset
    Dim i As Integer
 
    StrSQL = "select * From Groups  "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount
            delete_group_account RsTemp("GroupID").value
            RsTemp.MoveNext
        Next i

    End If

    RsTemp.Close
    StrSQL = "select * From Groups  where LastGroup=1 "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount

            If create_accounts(RsTemp("GroupID").value, RsTemp("GroupName").value) Then
                
            End If
         
            RsTemp.MoveNext
        Next i

    End If

    MsgBox "Done"

End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·„Ã„Ê⁄Â " & TxtId.Text & CHR(13) & " «”„  «·„Ã„Ê⁄Â   " & XPTxtName & CHR(13) & "   «·„Ã„Ê⁄Â «·—∆Ì”Ì… " & XPCboGroup & CHR(13) & "   ‰”»… «·«Â·«ﬂ  " & TXtPercentage1 & CHR(13) & "   ‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«› " & txtPercentage2

    If DepType.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "     ·Â «Â·«ﬂ   "
    End If
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Group Code " & TxtId.Text & CHR(13) & "  Group ¬Name " & XPTxtName & CHR(13) & "    Parent Group   " & XPCboGroup & CHR(13) & "    Dep %    " & TXtPercentage1 & CHR(13) & "   Dep % On Stop  " & txtPercentage2

    If DepType.value = Checked Then
        LogTexte = LogTexte & CHR(13) & " have Dep"
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 16915
    End If

End Sub

Private Sub DboParentAccount1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 251115
    End If
    
End Sub

Private Sub Form_Activate()
    'XPTxtID.SetFocus
End Sub

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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
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
    Dim Num As Integer
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

If mdifrmmain.Container.Visible = True Then
chkIsContainer.Visible = True
End If

    ScreenNameArabic = "»Ì«‰«  „Ã„Ê⁄«  «·«’Ê·"
    ScreenNameEnglish = "Fixed Assets Groups"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    Dim Dcombos As New ClsDataCombos
    
   Dcombos.GetAccountingCodes Me.DboParentAccount, False, True, 3
   
   Dcombos.GetAccountingCodes Me.DboParentAccount1, False, True, 0
   
 
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    TreeGroups.ImageList = mdifrmmain.ImgLstTree
    LoadTreeGroups Me.TreeGroups
    Set rs = New ADODB.Recordset
    StrSQL = "select * From FixedAssetsGroup where GroupID<>1"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FillGroupCombo
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = XPCboGroup
    XPBtnMove_Click 2
    LoadMenus
    Me.TxtModFlg.Text = "R"

    AddTip

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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
    Set GroupReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TreeGroups_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT
    Dim XNodeSeelcted As MSComctlLib.Node
 If Me.TxtModFlg = "E" Then Exit Sub
    If Me.TreeGroups.SelectedItem Is Nothing Then
        Exit Sub
    End If

    'TxtMenuState_Change
    If Button = vbRightButton Then
        GetCursorPos tp
        lX = (tp.X) * Screen.TwipsPerPixelX
        lY = tp.Y * Screen.TwipsPerPixelY
        XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TreeGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    'Dim NodeKey As String
    'NodeKey = left(Node.key, Len(Node.key) - 1)
    'If Node <> "" And NodeKey <> "1" Then
    '    Retrive (NodeKey)
    'End If
    Dim NodeKey As String
    If Me.TxtModFlg = "E" Then Exit Sub
    NodeKey = left(Node.Key, Len(Node.Key) - 1)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If right(Node.Key, 1) = "G" Then
        
            XPCboGroup.BoundText = val(Node.Key)
        
        End If

        Exit Sub
    End If

    If Node <> "" And NodeKey <> "1" Then
        Retrive (NodeKey)
    End If

End Sub

Private Sub TxtMenuState_Change()
    'Select Case TxtMenuState.Text
    '    Case "C"
    '        Me.XPPopUp.Menus(1).MenuItems(7).Enabled = True
    '        Me.XPBtnMove(0).Enabled = False
    '        Me.XPBtnMove(1).Enabled = False
    '        Me.XPBtnMove(2).Enabled = False
    '        Me.XPBtnMove(3).Enabled = False
    '    Case "N"
    '        Me.XPPopUp.Menus(1).MenuItems(7).Enabled = False
    '        Me.XPBtnMove(0).Enabled = True
    '        Me.XPBtnMove(1).Enabled = True
    '        Me.XPBtnMove(2).Enabled = True
    '        Me.XPBtnMove(3).Enabled = True
    'End Select
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   „Ã„Ê⁄«  «·«’Ê· «·À«» …"
            Else
                Me.Caption = "Fixed Assets Groups"
            End If
Chklast.Enabled = False
DepType.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = True
            Me.XPCboGroup.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            TreeGroups.Enabled = True
        
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  „Ã„Ê⁄«  «·«’Ê· «·À«» …( ÃœÌœ )"
            Else
                Me.Caption = "Fixed Assets   Groups(New Record)"
            End If
        Chklast.Enabled = True
DepType.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '     Me.XPBtnMove(0).Enabled = False
            '     Me.XPBtnMove(1).Enabled = False
            '     Me.XPBtnMove(2).Enabled = False
            '     Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '      TreeGroups.Enabled = False
            Me.lbl(10).Caption = ""
            Me.lbl(11).Caption = ""
            Me.lbl(12).Caption = ""
            Me.lbl(13).Caption = ""
            Me.lbl(14).Caption = ""
            Me.lbl(15).Caption = ""

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  „Ã„Ê⁄«  «·«’Ê· «·À«» …(  ⁄œÌ·)"
            Else
                Me.Caption = "Fixed Assets Groups(Edit Record)"
            End If
            If Chklast.value = vbChecked Then
        Chklast.Enabled = False
DepType.Enabled = False
Else
        Chklast.Enabled = True
DepType.Enabled = True
End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '  TreeGroups.Enabled = False
            '        TxtMenuState.Text = "C"
    End Select

    Exit Sub
ErrTrap:
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
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "GroupID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("GroupID").value), "", val(rs("GroupID").value))
    XPTxtName.Text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
    XPTxtNameE.Text = IIf(IsNull(rs("GroupNamee").value), "", Trim(rs("GroupNamee").value))
    DCPreFix.Text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.TxtId.Text = IIf(IsNull(rs("code").value), "", rs("code").value)
    If Not IsNull(rs("AsstMove").value) Then
    If (rs("AsstMove").value) = 1 Then
    RdMove(1).value = True
    Else
    RdMove(0).value = True
    End If
    Else
    RdMove(0).value = True
    End If
    

    If rs("IsContainer").value = vbTrue Then
        Me.chkIsContainer.value = vbChecked
    Else
        Me.chkIsContainer.value = vbUnchecked
    End If
 
        
    
'        rs("ParentExpensesAccount").value = (Me.DboParentAccount.BoundText)
    Me.DboParentAccount.BoundText = IIf(IsNull(rs("ParentExpensesAccount").value), get_account_code_branch(25, my_branch), (rs("ParentExpensesAccount")))
    Me.DboParentAccount1.BoundText = IIf(IsNull(rs("ParentEAssetAccount").value), get_account_code_branch(24, my_branch), (rs("ParentEAssetAccount")))
    
    
    Me.TXtPercentage1.Text = IIf(Not IsNumeric(rs("Percentage1").value), 0, val(rs("Percentage1")))
    Me.txtPercentage2.Text = IIf(Not IsNumeric(rs("Percentage2").value), 0, val(rs("Percentage2")))

    If IsNull(rs("DepType").value) Then
        DepType.value = vbUnchecked
   
    Else

        If rs("DepType").value = 1 Then
            DepType.value = vbChecked
        Else
            DepType.value = vbUnchecked
        End If
    End If
            
    If Not IsNull(rs("ParentID")) Then
        XPCboGroup.BoundText = rs("ParentID")
    Else
        XPCboGroup.Text = ""
    End If

    Me.TxtGroupCode.Text = IIf(IsNull(rs("GroupCode").value), "", Trim(rs("GroupCode").value))
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    If rs("LastGroup").value = True Then
        Chklast.value = vbChecked
    Else
        Chklast.value = Unchecked
    End If
        
    GetGroupInfo val(XPTxtID.Text)
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "GroupID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    FillGroupCombo

    If val(XPTxtID.Text) <> 0 Then
        Me.Retrive (XPTxtID.Text)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub delete_group_account(group_id As Integer)
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTemp2 As New ADODB.Recordset
    Dim i As Integer
    'On Error GoTo ErrTrap
    'groups_account_in_inventory

    '  StrSQL = "select * From Notes where BoxID=" & Trim(XPTxtBoxID.text)
    StrSQL = "select * From FixedAssetsGroupsAccount where group_id=" & group_id
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount
            StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & RsTemp("Account_Code").value & "'"
            RsTemp2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp2.RecordCount = 0 Then
                If ModAccounts.DeleteAccount(RsTemp("Account_Code").value) = True Then
       
                End If
                
            End If

            RsTemp2.Close
            RsTemp.MoveNext
        Next i
        
    End If

    StrSQL = "Delete From FixedAssetsGroupsAccount where group_id=" & group_id
    Cn.Execute StrSQL, , adExecuteNoRecords

End Sub

Private Sub Del_Group()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim StrAccountCode2 As String
    Dim StrAccountCode3 As String
    Dim ParetnAccount As String

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    
    If CheckChildforgroup("FixedAssetsGroup", "GroupID", "ParentID", val(XPTxtID.Text)) = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "·«Ì„ﬂ‰ Õ–› «·„Ã„Ê⁄Â ·ÊÃÊœ ·Â« «»‰«¡ ", vbCritical
     Else
     MsgBox "Can't Remove  Group it have Childs", vbCritical
     End If
         Exit Sub
    End If
    
        StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
        StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
        StrAccountCode3 = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
        ParetnAccount = IIf(IsNull(rs("ParetnAccount").value), "", rs("ParetnAccount").value)
    
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode1 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode2 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode3 & "'"
    
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–… «·„Ã„Ê⁄Â " & CHR(13)
            Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  ›Ì «·ﬁÌÊœ"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
    RsTemp.Close
 '   Set RsTemp = Nothing
    
    
    
            StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS1 where Account_Code='" & StrAccountCode & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode1 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode2 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode3 & "'"
    
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–… «·„Ã„Ê⁄Â " & CHR(13)
            Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«   ›Ì «·«—’œ… «·«›  «ÕÌ…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·„Ã„Ê⁄… —ﬁ„ " & CHR(13)
            Msg = Msg + (XPTxtID.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Confirm Delete Group " & CHR(13)
    
        End If
    
        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
            Exit Sub
        End If
    
        If Not rs.RecordCount < 1 Then

            ' delete_group_account Val(XPTxtID.text)
            '        TreeGroups.Nodes.Remove (Trim(Rs("GroupID").Value) & "G")
            If ModAccounts.DeleteAccount(StrAccountCode) = True And ModAccounts.DeleteAccount(StrAccountCode1) = True And ModAccounts.DeleteAccount(StrAccountCode2) = True And ModAccounts.DeleteAccount(StrAccountCode3) = True And ModAccounts.DeleteAccount(ParetnAccount) = True Then
                CuurentLogdata ("D")

                rs.delete
                Else
                Exit Sub
            End If
            
            TreeGroups.Nodes.Remove (Trim(XPTxtID.Text) & "G")
            rs.MoveFirst

            If rs.RecordCount < 1 Then
                FillGroupCombo
                Me.Retrive (XPTxtID.Text)
                clear_all Me
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
            Else
                FillGroupCombo
                Retrive
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

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–Â «·„Ã„Ê⁄… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Ã„Ê⁄… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·„Ã„Ê⁄…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„Ã„Ê⁄… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·„Ã„Ê⁄…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „Ã„Ê⁄…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record" & Wrap & "Enter New Group..." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print" & Wrap & "Print the current record data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this record data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or save the edit" & Wrap & "in the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete ...." & Wrap & "Delete the current record." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an Items Group" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Function create_accounts(group_id As Integer, group_name As String, Optional Checkonly As Boolean = False) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
  
 
 
    Dim Account_Code_dynamic As String
        If DboParentAccount1.BoundText = "" Then
    Account_Code_dynamic = get_account_code_branch(24, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬁÌ„… «·«’Ê· «·À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
       Else
       Account_Code_dynamic = DboParentAccount1.BoundText
       
       End If
       
       
    If DboParentAccount.BoundText = "" Then
    Dim Account_Code_dynamic1 As String
    Account_Code_dynamic1 = get_account_code_branch(25, my_branch)
        
    If Account_Code_dynamic1 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic1 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ      Õ”«» „’—Ê› «·«Â·«ﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
    
    Else
    
    Account_Code_dynamic1 = DboParentAccount.BoundText
    
    End If
    
    
    
    Dim Account_Code_dynamic2 As String
    Account_Code_dynamic2 = get_account_code_branch(26, my_branch)
        
    If Account_Code_dynamic2 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic2 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ   Õ”«» „Ã„⁄ «·«Â·«ﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Dim Account_Code_dynamic3 As String
    Dim Account_Code_dynamic4 As String
           
    If SystemOptions.AssetAccount1 = True Then
        Account_Code_dynamic3 = get_account_code_branch(31, my_branch)
        
        If Account_Code_dynamic3 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic3 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ     Õ”«» «—»«Õ »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
         
            End If
        End If
           
        Account_Code_dynamic4 = get_account_code_branch(40, my_branch)
        
        If Account_Code_dynamic4 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic4 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ  Õ”«» Œ”«—… »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
         
            End If
        End If
    End If
        
    If Checkonly = True Then
        GoTo ll
    End If
       
    Dim X As String

    If SystemOptions.AssetAccount = True Then
        X = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtName.Text, False, False, XPTxtNameE.Text)
        rs("ParetnAccount").value = X
        rs("Account_Code").value = ModAccounts.AddNewAccount(X, " ﬁÌ„Â " & XPTxtName.Text, True, False, XPTxtNameE.Text)
If DepType.value = vbChecked Then
        rs("Account_Code2").value = ModAccounts.AddNewAccount(X, "  „Ã„⁄ «Â·«ﬂ   " & XPTxtName.Text, True, False, XPTxtNameE.Text & " Accumulated depreciation")
 End If
    Else
        rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, " ﬁÌ„Â " & XPTxtName.Text, True, False, XPTxtNameE.Text & "Value")
       If DepType.value = vbChecked Then
        rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, "   „Ã„⁄ «Â·«ﬂ   " & XPTxtName.Text, True, False, XPTxtNameE.Text & " Accumulated depreciation")
        End If
    End If
     If DepType.value = vbChecked Then
    rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, "  „’—Ê›«    " & XPTxtName.Text, True, False, XPTxtNameE.Text & " Expenses ")
       End If
       
    If SystemOptions.AssetAccount1 = True Then
        rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, "  «—»«Õ »Ì⁄   " & XPTxtName.Text, True, False, XPTxtNameE.Text & " Sale Profit ")
        rs("Account_Code4").value = ModAccounts.AddNewAccount(Account_Code_dynamic4, " Œ”«—… »Ì⁄   " & XPTxtName.Text, True, False, XPTxtNameE.Text & " Sale Loss ")
    End If
    
    
        
ll:
  
    create_accounts = True
    Exit Function
ErrTrap:

    create_accounts = False

End Function

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim XNode As MSComctlLib.Node
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If XPTxtName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ã„Ê⁄…"
            Else
                Msg = "plz enter group name firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtName.SetFocus
            Exit Sub
        End If

        If XPCboGroup.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ÌÃ»  ÕœÌœ «·„Ã„Ê⁄… «·—∆Ì”Ì…" & CHR(13)
            Else
                Msg = "Select primary group First" & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboGroup.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
   
        If Me.Chklast = vbChecked Then
            If Not IsNumeric(TXtPercentage1.Text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "„‰ ›÷·ﬂ √œŒ·    ‰”»… «Â·«ﬂ ’ÕÌÕ…"
                Else
                    Msg = "plz enter Damage Percentage"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TXtPercentage1.SetFocus
                Exit Sub
        
            End If
   
            If Not IsNumeric(txtPercentage2.Text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "„‰ ›÷·ﬂ √œŒ·    ‰”»… «Â·«ﬂ ›Ì Õ«·… «· Êﬁ›  ’ÕÌÕ…"
                Else
                    Msg = "plz enter Damage Percentage On Stop"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                txtPercentage2.SetFocus
                Exit Sub
        
            End If
  
        End If
  'salimhere And DepType.value = vbChecked
        If Chklast.value = vbChecked Then '«· √ﬂÌœ  ⁄·Ï «·Õ”«»«  ›Ì Õ«·Â „Ã„Ê⁄Â ‰Â«∆Ì… ·Â« «Â·«ﬂ
            If create_accounts(XPTxtID.Text, XPTxtName.Text, True) = False Then
                Exit Sub
            End If
        End If
           
        Me.TxtGroupCode.Text = GetNewGroupCode(Me.XPCboGroup.BoundText)

        Select Case TxtModFlg.Text

            Case "N"
                StrSQL = "select * From FixedAssetsGroup where GroupName='" & Trim(XPTxtName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                
                    Else
                        Msg = "This group Name Already Exisi" & CHR(13)
            
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            
                XPTxtID.Text = CStr(new_id("FixedAssetsGroup", "GroupID", "", True))

            Case "E"
                StrSQL = "select * From FixedAssetsGroup where GroupName='" & Trim(XPTxtName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("GroupID").value <> val(XPTxtID.Text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                        Else
                            Msg = "This group Name Already Exisi" & CHR(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If

        End Select
     
        Select Case TxtModFlg.Text

            Case "N"
                Cn.BeginTrans
                BeginTrans = True
            
                rs.AddNew
                rs("GroupID").value = IIf(XPTxtID.Text = "", "", val(XPTxtID.Text))
      'salimhere And DepType.value = vbChecked
              If Chklast.value = vbChecked Then '«‰‘«¡ «·Õ”«»«  ›Ì Õ«·Â „Ã„Ê⁄Â ‰Â«∆Ì… ·Â« «Â·«ﬂ
                    If create_accounts(XPTxtID.Text, XPTxtName.Text) Then
                
                    End If
                End If
            
            Case "E"

                If XPTxtName.Text = XPCboGroup.Text Then
                    Msg = "·«Ì„ﬂ‰ √‰  ﬂÊ‰ «·„Ã„Ê⁄… «·—∆Ì”Ì… " & CHR(13)
                    Msg = Msg + "ÂÌ ‰›” «·„Ã„Ê⁄… «·›—⁄Ì…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                Cn.BeginTrans
                BeginTrans = True
        End Select
        
If RdMove(1).value = True Then
rs("AsstMove").value = 1
Else
rs("AsstMove").value = 0
End If
        rs("GroupName").value = IIf(XPTxtName.Text = "", "", Trim(XPTxtName.Text))
        rs("GroupNamee").value = IIf(XPTxtNameE.Text = "", "", Trim(XPTxtNameE.Text))
        
        rs("ParentID").value = XPCboGroup.BoundText
        rs("GroupCode").value = Me.TxtGroupCode.Text
        rs("Percentage1").value = val(Me.TXtPercentage1.Text)
        rs("Percentage2").value = val(Me.txtPercentage2.Text)
        rs("ParentExpensesAccount").value = (Me.DboParentAccount.BoundText)
        rs("ParentEAssetAccount").value = (Me.DboParentAccount1.BoundText)
        
        rs("code").value = TxtId.Text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(TxtId.Text) = "", Null, TxtId.Text)
        rs("prifix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
         
        If Me.chkIsContainer.value = vbChecked Then
            rs.Fields("IsContainer").value = 1
        Else
            rs.Fields("IsContainer").value = 0
        End If
          
         
        If Chklast.value = vbChecked Then
            rs("LastGroup").value = 1
        Else
            rs("LastGroup").value = 0
        End If
        
        If DepType.value = vbChecked Then
            rs("DepType").value = 1
        Else
            rs("DepType").value = 0
        End If
            
If IsNull(rs("Account_Code").value) Then
       
     ' Salim here And DepType.value = vbChecked
              If Chklast.value = vbChecked Then '«‰‘«¡ «·Õ”«»«  ›Ì Õ«·Â „Ã„Ê⁄Â ‰Â«∆Ì… ·Â« «Â·«ﬂ
                    If create_accounts(XPTxtID.Text, XPTxtName.Text) Then
                GoTo ll
                    End If
                End If
       
       
End If
            
        If Not IsNull(rs("ParetnAccount").value) Then
            ModAccounts.EditAccount rs("ParetnAccount").value, Me.XPTxtName.Text, Trim(XPTxtNameE.Text), , , , , , , , , , , , , , , , , False
        End If
            
        If Not IsNull(rs("Account_Code").value) Then
            ModAccounts.EditAccount rs("Account_Code").value, " ﬁÌ„Â " & Me.XPTxtName.Text, Trim(XPTxtNameE.Text) & " Value  ", , , , , , , , , , , , , , , , , True
        End If
            
       If DepType.value = vbChecked Then
        If Not IsNull(rs("Account_Code1").value) Then
            ModAccounts.EditAccount rs("Account_Code1").value, " „’—Ê›«  «Â·«ﬂ  " & Me.XPTxtName.Text, Trim(XPTxtNameE.Text) & " Expenses ", , , , , , , , , , , , , , , , , True
        End If
        End If
        
        
           If DepType.value = vbChecked Then
        If Not IsNull(rs("Account_Code2").value) Then
            ModAccounts.EditAccount rs("Account_Code2").value, " „Ã„⁄ «Â·«ﬂ  " & Me.XPTxtName.Text, Trim(XPTxtNameE.Text) & "   Accumulated Depreciation ", , , , , , , , , , , , , , , , , True
        End If
            
            End If
            
        If SystemOptions.AssetAccount1 = True Then
            If Not IsNull(rs("Account_Code3").value) Then
                ModAccounts.EditAccount rs("Account_Code3").value, " «—»«Õ »Ì⁄ " & Me.XPTxtName.Text, Trim(XPTxtNameE.Text) & " Sale Profit ", , , , , , , , , , , , , , , , , True
            End If
                            
            If Not IsNull(rs("Account_Code4").value) Then
                ModAccounts.EditAccount rs("Account_Code4").value, " Œ”«—… »Ì⁄ " & Me.XPTxtName.Text, Trim(XPTxtNameE.Text) & " Sale Loss ", , , , , , , , , , , , , , , , , True
            End If
        End If
ll:
        rs.update
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        FillGroupCombo

        If TxtModFlg.Text = "E" Then
            'TreeGroups.Nodes.Remove (Trim(Rs("GroupID").Value) & "G")
            If SystemOptions.UserInterface = ArabicInterface Then
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").Text = Trim(rs("Fullcode")) & "" & rs("GroupName").value
            Else
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").Text = Trim(rs("Fullcode")) & "" & rs("GroupNamee").value
            End If
        ElseIf TxtModFlg.Text = "N" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Set XNode = TreeGroups.Nodes.Add(Trim(rs("ParentID").value) & "G", tvwChild, Trim(rs("GroupID").value) & "G", Trim(rs("Fullcode")) & "" & Trim(rs("GroupName").value), "Closed_Node", "Open_Node")
         Else
             Set XNode = TreeGroups.Nodes.Add(Trim(rs("ParentID").value) & "G", tvwChild, Trim(rs("GroupID").value) & "G", Trim(rs("Fullcode")) & "" & Trim(rs("GroupNamee").value), "Closed_Node", "Open_Node")
         End If
         
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").Selected = True
            
        End If

        Me.Retrive (XPTxtID.Text)
        CuurentLogdata
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„Ã„Ê⁄…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Data was Saved , do you want to enter another data y/n" & CHR(13)
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If

        End Select

        TxtModFlg.Text = "R"
        TreeGroups.SetFocus
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
            Msg = "Can't Save error in entered data " & CHR(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« " & CHR(13)
        Else
            Msg = "Can't save Data , Reasons: Data integrity " & CHR(13)
 
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        rs.CancelUpdate
        Exit Sub
    End If

    Select Case Me.TxtModFlg.Text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Else
                Msg = "Sorry...... Error During Saving Data " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «· ⁄œÌ·«  " & CHR(13)
            Else
                Msg = "Sorry...... Error During Saving cahanges" & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End Select


End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Set GroupReport = New ClsGroupReport
        GroupReport.GroupData XPTxtID.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FillGroupCombo()
    On Error GoTo ErrTrap
    Dim Num As Integer
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
 
   If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT GroupID,   isnull(fullcode,' ') + ' ' +  isnull(GroupName,'')    as grupname FROM FixedAssetsGroup Order By GroupName"
   Else
   StrSQL = "SELECT GroupID,  isnull(fullcode,' ') + ' ' +  isnull(GroupNamee,'')  as grupnamee FROM FixedAssetsGroup Order By GroupNamee"
   End If
   
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    fill_combo XPCboGroup, StrSQL
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = XPCboGroup
    cSearchDcbo.Refresh
    'XPCboGroup.Clear
    'If Not (RsTemp.EOF Or RsTemp.BOF) Then
    '    RsTemp.MoveFirst
    '    'XPCboGroup
    '    For Num = 0 To RsTemp.RecordCount - 1
    '        XPCboGroup.AddItem RsTemp("GroupName").Value, Num
    '        XPCboGroup.ItemData(Num) = RsTemp("GroupID").Value
    '        RsTemp.MoveNext
    '    Next Num
    '    RsTemp.MoveFirst
    'End If
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

Private Sub LoadTreeGroups(ItemsTree As MSComctlLib.TreeView)
    Dim Rs_items As ADODB.Recordset
    Dim My_SQL As String
    Dim nodX As Node
    Dim nodz As Node
    Dim RsOptions As ADODB.Recordset
    Dim my_ch_rs As ADODB.Recordset
    Dim BolDisplayArabic As Boolean
    Dim LngLoop As Long
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "„Ã„Ê⁄… «·«’Ê· «·À«» …", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "1G", " Fixed Assets Groups ", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    End If

    Me.TreeGroups.Sorted = False
    '''''''''''''''''''''''''''' add group
    My_SQL = " SELECT FixedAssetsGroup.* "
    My_SQL = My_SQL + "  From FixedAssetsGroup  "
    My_SQL = My_SQL + " where (ParentID =1); "
    Set my_ch_rs = New ADODB.Recordset
    my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ' BolDisplayArabic = True

    If BolDisplayArabic = True Then
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "FixedAssetsGroup", "ParentID")
    Else
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "FixedAssetsGroup", "ParentID", , 8)
    End If

    ItemsTree.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadMenus()
    On Error GoTo ErrTrap

    With Me.XPPopUp
        '~~Clear the Menu and ToolBars
        .ClearAll

        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True
            .SetImageList mdifrmmain.img16

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, "...≈÷«›… „Ã„Ê⁄…", False, True, 2, , 2, , , "AddGroup", , , , "≈÷«›… „Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, " ⁄œÌ·", False, True, 3, , , , , "EditGroup", , , , " ⁄œÌ·"
                .MenuItems.Add tsMenuCaption, "Õ–›", False, True, 4, , , , , "DelGroup", , , , "Õ–› «·„Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, "„”Õ «·«Œ Ì«—", False, False, 5, , , , , "ClearGroup", , , , "„”Õ «·«Œ Ì«—"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ﬁ’", False, False, 6, , , True, , "CutGroup", , , , "ﬁ’"
                .MenuItems.Add tsMenuCaption, "·’ﬁ", False, False, 7, , , , , "PasteGroup", , , , "·’ﬁ"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "‰ﬁ· √’‰«› «·„Ã„Ê⁄… ≈·Ï ", False, False, 8, , , True, , "RemoveGroup", , , , "‰ﬁ· √’‰«› «·„Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, "Œ’«∆’", False, False, 9, , , True, , "GroupProperties", , , , "Œ’«∆’"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ÿ»«⁄… ‘Ã—… «·«’Ê·", False, False, 10, , , True, , "PrintGroup", , , , "ÿ»«⁄…"
            End With

        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .RightToLeft = False
            .SetImageList mdifrmmain.img16

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, "Add New Group...", False, True, 2, , 2, , , "AddGroup", , , , "Add New Group"
                .MenuItems.Add tsMenuCaption, "Edit", False, True, 3, , , , , "EditGroup", , , , "Edit this Group"
                .MenuItems.Add tsMenuCaption, "Delete", False, True, 4, , , , , "DelGroup", , , , "Delete This Group"
                .MenuItems.Add tsMenuCaption, "Clear Checked", False, False, 5, , , , , "ClearGroup", , , , "Clear This Group"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Cut", False, False, 6, , , True, , "CutGroup", , , , "Cut"
                .MenuItems.Add tsMenuCaption, "Paste", False, False, 7, , , , , "PasteGroup", , , , "Paste"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Move this items group to...", False, False, 8, , , True, , "RemoveGroup", , , , "Move the items of this group to another group"
                .MenuItems.Add tsMenuCaption, "Properties", False, False, 9, , , True, , "GroupProperties", , , , "Properties"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Print Fixed Assets Tree", False, False, 10, , , True, , "PrintGroup", , , , "Print all of the tree of items"
            End With

        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPPopUp_MenuItemClick(ByVal MenuIndex As Integer, _
                                  ByVal MenuID As String, _
                                  ByVal MenuItemIndex As Integer, _
                                  ByVal MenuItemID As String)
    On Error GoTo ErrTrap
    Dim XNode As MSComctlLib.Node
    Dim StrSQL As String

    Select Case MenuItemID

        Case "AddGroup"
            Cmd_Click (0)
            Set XNode = TreeGroups.SelectedItem
            XPCboGroup.BoundText = left(XNode.Key, Len(XNode.Key) - 1)

        Case "EditGroup"
            Cmd_Click (1)

        Case "DelGroup"
            Cmd_Click (4)

        Case "ClearGroup"

        Case "CutGroup"
            TreeGroups.SelectedItem.backcolor = vbGreen
            TxtCutKey.Text = (TreeGroups.SelectedItem.Key)

            '        Me.TxtMenuState.Text = "C"
        Case "PasteGroup"
            TreeGroups.Nodes.Remove (TxtCutKey.Text)
            Set XNode = TreeGroups.Nodes.Add(Trim(TreeGroups.SelectedItem.Key), tvwChild, rs("GroupID") & "G", rs("GroupName"), "Closed_Node", "Open_Node")
            StrSQL = "update Groups set ParentID=" & val(left(TreeGroups.SelectedItem.Key, Len(TreeGroups.SelectedItem.Key) - 1)) & " where GroupID=" & val(rs("GroupID").value)
            Cn.Execute StrSQL
            Retrive (val(rs("GroupID").value))

            '        Me.TxtMenuState.Text = "N"
        Case "GroupProperties"

        Case "PrintGroup"
    End Select

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
    RdMove(0).Caption = "Movable"
    RdMove(1).Caption = "Not Movable"
    lbl(17).Caption = "Depreciation %"
    lbl(18).Caption = "Stop %"
    DepType.Caption = "Have Depreciation"
    Me.Caption = "Fixed Assets Groups"
    Me.LblHeader.Caption = Me.Caption
    lbl(16).Caption = "Last Group"
    Me.lbl(6).Caption = "Group ID"
    Me.lbl(3).Caption = "Group Code"
    Me.lbl(5).Caption = " Name AR"
    Me.lbl(19).Caption = " Name Eng"
lbl(21).Caption = "Asset Accoount"
    Me.lbl(0).Caption = "Parent Group"
    Me.lbl(1).Caption = "Current Record"
    Me.lbl(2).Caption = "NO. Recordes"
lbl(20).Caption = "Expenses Acc."

    ELe.Caption = "More Information"
    ELe.Font.Bold = True
    Me.lbl(4).Caption = "Sub Groups Count"
    Me.lbl(7).Caption = "Items Group Count"
    Me.lbl(8).Caption = "Frist item added to the group"
    Me.lbl(9).Caption = "Last item added to the group"
    
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

End Sub

Private Function GetNewGroupCode(LngParentGroupID As Long) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String

    On Error GoTo ErrTrap
    StrSQL = "Select GroupCode From FixedAssetsGroup Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From FixedAssetsGroup Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function

Private Sub GetGroupInfo(Optional Lngid As Long = 0)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Lngid = 0 Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "SELECT Count(FixedAssetsGroup.GroupID) AS CountGroupID"
    StrSQL = StrSQL + " From FixedAssetsGroup WHERE (FixedAssetsGroup.ParentID=" & Lngid & ")"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Me.lbl(10).Caption = "0"
    Else
        Me.lbl(10).Caption = IIf(IsNull(rs("CountGroupID").value), 0, rs("CountGroupID").value)
    End If

    rs.Close
    Exit Sub
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT Count(Groups.GroupID) AS CountGroupID "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID "
    StrSQL = StrSQL + " Where (((TblItems.GroupID) =" & Lngid & "))"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Me.lbl(11).Caption = "0"
    Else
        Me.lbl(11).Caption = IIf(IsNull(rs("CountGroupID").value), 0, rs("CountGroupID").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT TblItems.* "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
    StrSQL = StrSQL + " WHERE (((TblItems.GroupID)=" & Lngid & ")) Order By TblItems.ItemID ASC;"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Me.lbl(12).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        Me.lbl(14).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
        rs.MoveLast
        Me.lbl(15).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        Me.lbl(13).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
    
    Else
        Me.lbl(12).Caption = ""
        Me.lbl(13).Caption = ""
        Me.lbl(14).Caption = ""
        Me.lbl(15).Caption = ""
    End If

End Sub

Private Sub XPTxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
