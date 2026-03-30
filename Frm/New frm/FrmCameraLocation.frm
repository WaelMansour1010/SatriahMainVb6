VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{E1BFA30F-D929-4F80-AEDD-76FC2BDF5E23}#1.0#0"; "ciaXPPopUp30.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form Frmcameralocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  „Ê«ﬁ⁄ «·ﬂ«„Ì—« "
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   HelpContextID   =   200
   Icon            =   "FrmCameraLocation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9120
   Begin VB.CommandButton CMDView 
      Caption         =   "⁄—÷"
      Height          =   615
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox TxtLink 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   2040
      Width           =   3195
   End
   Begin VB.ComboBox CboEXpirType 
      Height          =   315
      ItemData        =   "FrmCameraLocation.frx":038A
      Left            =   5640
      List            =   "FrmCameraLocation.frx":0397
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtEXpireValue 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6600
      TabIndex        =   47
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2210
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«‰‘«¡ «·Õ”«»« "
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Chklast 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
      Height          =   255
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   2295
      Left            =   10740
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4230
      Width           =   4545
      _cx             =   8017
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
      Picture         =   "FrmCameraLocation.frx":03AA
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
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   14
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   330
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√’‰«› «· Ï  Õ ÊÌÂ« «·„Ã„Ê⁄…"
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
         TabIndex        =   37
         Top             =   810
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ê· ’‰› „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
         TabIndex        =   36
         Top             =   1230
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ— ’‰› „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1950
         Width           =   3375
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   12
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   11
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   225
         Index           =   10
         Left            =   3390
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   570
         Width           =   855
      End
   End
   Begin VB.TextBox TxtGroupCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9540
      Width           =   1875
   End
   Begin VB.TextBox XPTxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   60
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1245
      Width           =   3315
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox TxtCutKey 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   90
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   150
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtMenuState 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "N"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1185
      TabIndex        =   5
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
      ButtonImage     =   "FrmCameraLocation.frx":0744
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
      TabIndex        =   6
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
      ButtonImage     =   "FrmCameraLocation.frx":0ADE
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
      TabIndex        =   7
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
      ButtonImage     =   "FrmCameraLocation.frx":0E78
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
      TabIndex        =   8
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
      ButtonImage     =   "FrmCameraLocation.frx":1212
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
      Left            =   9750
      Top             =   5070
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
      Top             =   1605
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.TreeView TreeGroups 
      Height          =   6045
      Left            =   4680
      TabIndex        =   12
      Top             =   630
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   10663
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   7155
      TabIndex        =   20
      Top             =   7200
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
      Left            =   6405
      TabIndex        =   21
      Top             =   7200
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
      Left            =   5685
      TabIndex        =   22
      Top             =   7200
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
      Left            =   4965
      TabIndex        =   23
      Top             =   7200
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
      Left            =   4035
      TabIndex        =   24
      Top             =   7200
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
      Left            =   3180
      TabIndex        =   25
      Top             =   7200
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
      Left            =   630
      TabIndex        =   26
      Top             =   7200
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
      Left            =   2340
      TabIndex        =   27
      Top             =   7200
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
      Left            =   1350
      TabIndex        =   28
      Top             =   7200
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
      Left            =   1320
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   1095
      Left            =   5400
      TabIndex        =   51
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
      _Version        =   131072
      _ExtentX        =   5530
      _ExtentY        =   1931
      _StockProps     =   1
      BackColor       =   12632256
      _Image          =   "FrmCameraLocation.frx":15AC
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   7080
      TabIndex        =   52
      Top             =   3240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«œ—«Ã ’Ê—… «·„Ã„Ê⁄Â"
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—»ÿ"
      Height          =   315
      Index           =   20
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’Ê—…"
      Height          =   315
      Index           =   19
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’·«ÕÌ…"
      Height          =   315
      Index           =   18
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„Êﬁ⁄"
      Height          =   315
      Index           =   17
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
      Height          =   315
      Index           =   16
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
      Height          =   315
      Index           =   3
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   1
      Left            =   3210
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6720
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   2
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
      Height          =   315
      Index           =   0
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Êﬁ⁄"
      Height          =   315
      Index           =   5
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·„Ã„Ê⁄…"
      Height          =   315
      Index           =   6
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   9540
      Width           =   1095
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6750
      Width           =   765
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "»Ì«‰«  „Ê«ﬁ⁄ «·ﬂ«„Ì—« "
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   8955
   End
End
Attribute VB_Name = "Frmcameralocation"
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
    Dim currentgroup As Integer
    'On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
        
            currentgroup = val(XPCboGroup.BoundText)
            TxtModFlg.text = "N"
        
            clear_all Me

            '        XPTxtID.text = CStr(new_id("Groups", "GroupID", "", True))
            If currentgroup = 0 Then
                XPCboGroup.BoundText = 1
            Else
                XPCboGroup.BoundText = currentgroup

            End If

            If Me.XPCboGroup.BoundText <> "" Then
            
            End If

            XPTxtName.SetFocus

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            CuurentLogdata

        Case 2

     '       Dim currentcode As String
 '
 '           If txtid = "" Then
 '               MsgBox "«œŒ· ﬂÊœ «·„Ã„Ê⁄Â     "
 '               Exit Sub
 '
 '           End If

 

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Group

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            FrmGroupSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
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

Private Sub CMDView_Click()
OpenWebSite Me.TxtLink
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

Private Sub Form_Activate()
    XPTxtID.SetFocus
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

    ScreenNameArabic = " »Ì«‰«     „Ê«ﬁ⁄ «·ﬂ«„Ì—«  "
    ScreenNameEnglish = "Cameras Locations "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

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
    StrSQL = "select * From TblCameraLocations where GroupID<>1"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FillGroupCombo
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = XPCboGroup
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetPrefix Me.DCPreFix, 2, val(branch_id)

    XPBtnMove_Click 2
    LoadMenus
    Me.TxtModFlg.text = "R"

    AddTip

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish

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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " ﬂÊœ «·„Ã„Ê⁄… " & txtid.text & Chr(13) & "   «”„ «·„Ã„Ê⁄Â" & XPTxtName.text & Chr(13) & " «·„Ã„Ê⁄Â «·—∆Ì”Ì…   " & XPCboGroup.text

    If Chklast.value = vbChecked Then
        LogTextA = LogTextA & " „Ã„Ê⁄Â ‰Â«∆Ì…"
    End If
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & "Code " & txtid.text & Chr(13) & "   Name  " & XPTxtName.text & Chr(13) & " Parent Group" & XPCboGroup.text

    If Chklast.value = vbChecked Then
        LogTextA = LogTextA & "  Final Group"
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
    
End Function

Private Sub ISButton1_Click()
Dim x As Integer
    If txtid.text = "" Then Exit Sub
    x = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If x = vbYes Then
        DBPix201.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If x = vbNo Then
            DBPix201.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix201.ImageSaveFile (App.path & "\images\pos\" & XPTxtID.text & ".JPG")
End Sub

Private Sub TreeGroups_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT
    Dim XNodeSeelcted As MSComctlLib.Node

    If Me.TreeGroups.SelectedItem Is Nothing Then
        Exit Sub
    End If

    'TxtMenuState_Change
    If Button = vbRightButton Then
        GetCursorPos tp
        lX = (tp.x) * Screen.TwipsPerPixelX
        lY = tp.Y * Screen.TwipsPerPixelY
        XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TreeGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim NodeKey As String
    NodeKey = left(Node.key, Len(Node.key) - 1)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If right(Node.key, 1) = "G" Then
        
            XPCboGroup.BoundText = val(Node.key)
        
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

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  „Ê«ﬁ⁄ «·⁄„·"
            Else
                Me.Caption = "Items Groups"
            End If

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
        
            XPCboGroup.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  „Ê«›⁄ «·⁄„·( ÃœÌœ )"
            Else
                Me.Caption = "Items Groups(New Record)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            XPCboGroup.Enabled = True
            '     Me.XPBtnMove(0).Enabled = False
            '     Me.XPBtnMove(1).Enabled = False
            '     Me.XPBtnMove(2).Enabled = False
            '     Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '   TreeGroups.Enabled = False
            Me.lbl(10).Caption = ""
            Me.lbl(11).Caption = ""
            Me.lbl(12).Caption = ""
            Me.lbl(13).Caption = ""
            Me.lbl(14).Caption = ""
            Me.lbl(15).Caption = ""

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  „Ê«ﬁ⁄ «·⁄„· (  ⁄œÌ·)"
            Else
                Me.Caption = "Items Groups(Edit Record)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            XPCboGroup.Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '   TreeGroups.Enabled = False
            '        TxtMenuState.Text = "C"
    End Select

    Exit Sub
ErrTrap:
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

    XPTxtID.text = IIf(IsNull(rs("GroupID").value), "", val(rs("GroupID").value))
    XPTxtName.text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))

    If Not IsNull(rs("ParentID")) Then
        XPCboGroup.BoundText = rs("ParentID")
    Else
        XPCboGroup.text = ""
    End If

    Me.TxtGroupCode.text = IIf(IsNull(rs("GroupCode").value), "", Trim(rs("GroupCode").value))
    Me.TxtLink.text = IIf(IsNull(rs("Link").value), "", Trim(rs("Link").value))
    
    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)

    If XPTxtID.text <> "" Then
        DBPix201.ImageClear
 

        If Dir(App.path & "\images\pos\" & XPTxtID.text & ".JPG") <> "" Then
            DBPix201.ImageLoadFile (App.path & "\images\pos\" & XPTxtID.text & ".JPG")
        End If

 
 
    End If
    
    CboEXpirType.ListIndex = IIf(IsNull(rs("EXpirType").value), -1, rs("EXpirType").value)

    If CboEXpirType.ListIndex = -1 Then
        TxtEXpireValue.text = ""
    Else
        Me.TxtEXpireValue.text = IIf(IsNull(rs("EXpireValue").value), "", rs("EXpireValue").value)
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    If rs("LastGroup").value = True Then
        Chklast.value = vbChecked
    Else
        Chklast.value = Unchecked
    End If
        
    GetGroupInfo val(XPTxtID.text)
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "GroupID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    FillGroupCombo

    If val(XPTxtID.text) <> 0 Then
        Me.Retrive (XPTxtID.text)
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
    StrSQL = "select * From groups_account_in_inventory where group_id=" & group_id
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

    StrSQL = "Delete From groups_account_in_inventory where group_id=" & group_id
    Cn.Execute StrSQL, , adExecuteNoRecords

End Sub

Private Sub Del_Group()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    
        StrSQL = "select * From  TblEmployee  where GroupID=" & Trim(XPTxtID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› Â–Â «·„Ã„Ê⁄… " & Chr(13)
                Msg = Msg + "Â‰«ﬂ „ÊŸ›Ì‰  ‰œ—Ã  Õ  Â–Â «·„Ã„Ê⁄…"
            Else
                Msg = "Can't Delete this Group because it have Items"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        RsTemp.Close
    
        StrSQL = "select * From TblCameraLocations where ParentID=" & Trim(XPTxtID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› Â–Â «·„Ã„Ê⁄… " & Chr(13)
                Msg = Msg + "Â‰«ﬂ „Ã„Ê⁄«    ‰œ—Ã  Õ  Â–Â «·„Ã„Ê⁄…"
            Else
                Msg = "Can't Delete this Group because it have Chilrd Goup "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·„Ã„Ê⁄… —ﬁ„ " & Chr(13)
            Msg = Msg + (XPTxtID.text) & Chr(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Confirm Delete Group " & Chr(13)
    
        End If
    
        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
            Exit Sub
        End If

        If Not rs.RecordCount < 1 Then
           
            CuurentLogdata ("D")
            rs.delete
            TreeGroups.Nodes.Remove (Trim(XPTxtID.text) & "G")
            rs.MoveFirst

            If rs.RecordCount < 1 Then
                FillGroupCombo
                Me.Retrive (XPTxtID.text)
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–Â «·„Ã„Ê⁄… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
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

Function create_accounts(group_id As Integer, group_name As String) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
        If rsOut!opt_group = False Then
            Current_case = -1
        ElseIf rsOut!opt_group = True And rsOut!Opt_Inventory_create_account = 1 Then
            Current_case = 0 '„Œ«“‰ ›ﬁÿ
        ElseIf rsOut!opt_group = True And rsOut!opt_inv_and_branch_create_account = 1 Then
            Current_case = 1 '„Œ«“‰ Ê›—⁄
        End If
    End If

    If Current_case = -1 Then Exit Function
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from TblStore "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    For i = 1 To Rs3.RecordCount

        If create_inventory_group(Rs3("StoreID").value, group_id, Rs3("StoreName").value & " " & group_name) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close

    Select Case Current_case

        Case 1:
            sql = "Select * from branches "
 
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
            If Rs3.RecordCount = 0 Then Exit Function

            For i = 1 To Rs3.RecordCount

                If create_Branch_group(Rs3("branch_id").value, group_id, group_name) = True Then
                End If

                Rs3.MoveNext
            Next i

            Rs3.Close

    End Select

End Function

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim XNode As MSComctlLib.Node
'    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ã„Ê⁄…"
            Else
                Msg = "plz enter group name firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtName.SetFocus
            Exit Sub
        End If

        If XPCboGroup.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ÌÃ»  ÕœÌœ «·„Ã„Ê⁄… «·—∆Ì”Ì…" & Chr(13)
            Else
                Msg = "Select primary group First" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboGroup.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
   
        Me.TxtGroupCode.text = GetNewGroupCode(Me.XPCboGroup.BoundText)

        Select Case TxtModFlg.text

            Case "N"
                StrSQL = "select * From TblCameraLocations where GroupName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"

                    Else
                        Msg = "This group Name Already Exisi" & Chr(13)

                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            
                XPTxtID.text = CStr(new_id("TblCameraLocations", "GroupID", "", True))

            Case "E"
                StrSQL = "select * From TblCameraLocations where GroupName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("GroupID").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                        Else
                            Msg = "This group Name Already Exisi" & Chr(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If

        End Select
     
        Select Case TxtModFlg.text

            Case "N"
                Cn.BeginTrans
                BeginTrans = True
            
                rs.AddNew
                rs("GroupID").value = IIf(XPTxtID.text = "", "", val(XPTxtID.text))

 
            
            Case "E"

                If XPTxtName.text = XPCboGroup.text Then
                    Msg = "·«Ì„ﬂ‰ √‰  ﬂÊ‰ «·„Ã„Ê⁄… «·—∆Ì”Ì… " & Chr(13)
                    Msg = Msg + "ÂÌ ‰›” «·„Ã„Ê⁄… «·›—⁄Ì…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                Cn.BeginTrans
                BeginTrans = True
        End Select

        rs("GroupName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
        rs("Link").value = IIf(TxtLink.text = "", "", Trim(TxtLink.text))
        
        rs("ParentID").value = XPCboGroup.BoundText
        rs("GroupCode").value = Me.TxtGroupCode.text
        rs("Branch_NO").value = val(branch_id)
        rs("code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
  
        If CboEXpirType.ListIndex = -1 Then
            rs("EXpirType").value = Null
            rs("EXpireValue").value = Null
        Else
            rs("EXpirType").value = (CboEXpirType.ListIndex)
            rs("EXpireValue").value = val(TxtEXpireValue.text)
        End If
  
        If Chklast.value = vbChecked Then
            rs("LastGroup").value = 1
        Else
            rs("LastGroup").value = 0
        End If
        
        rs.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        FillGroupCombo

        If TxtModFlg.text = "E" Then
            'TreeGroups.Nodes.Remove (Trim(Rs("GroupID").Value) & "G")
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").text = rs("GroupName").value
        ElseIf TxtModFlg.text = "N" Then
            Set XNode = TreeGroups.Nodes.Add(Trim(rs("ParentID").value) & "G", tvwChild, Trim(rs("GroupID").value) & "G", Trim(rs("GroupName").value), "Closed_Node", "Open_Node")
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").Selected = True
        End If

        Me.Retrive (XPTxtID.text)
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„Ã„Ê⁄…" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Data was Saved , do you want to enter another data y/n" & Chr(13)
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
    
                TreeGroups.ImageList = mdifrmmain.ImgLstTree
                LoadTreeGroups Me.TreeGroups

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If

        End Select

        TxtModFlg.text = "R"
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
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Can't Save error in entered data " & Chr(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« " & Chr(13)
        Else
            Msg = "Can't save Data , Reasons: Data integrity " & Chr(13)
 
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        rs.CancelUpdate
        Exit Sub
    End If

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Else
                Msg = "Sorry...... Error During Saving Data " & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «· ⁄œÌ·«  " & Chr(13)
            Else
                Msg = "Sorry...... Error During Saving cahanges" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End Select

End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Set GroupReport = New ClsGroupReport
        GroupReport.GroupData XPTxtID.text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FillGroupCombo()
    On Error GoTo ErrTrap
    Dim Num As Integer
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    StrSQL = "SELECT * FROM TblCameraLocations Order By GroupName"
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
    Dim lngCount As Long
        
    With ItemsTree.Nodes

        For lngCount = .count To 1 Step -1
            .Remove lngCount
        Next

    End With
    
    If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "„Ê«ﬁ⁄ «·ﬂ«„Ì—« ", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "Cameras Locations", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    End If

    Me.TreeGroups.Sorted = False
    '''''''''''''''''''''''''''' add group
    My_SQL = " SELECT TblCameraLocations.* "
    My_SQL = My_SQL + "  From TblCameraLocations "
    My_SQL = My_SQL + " where (ParentID =1); "
    Set my_ch_rs = New ADODB.Recordset
    my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    BolDisplayArabic = True

    If BolDisplayArabic = True Then
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "TblCameraLocations", "ParentID")
    Else
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "TblCameraLocations", "ParentID", , 2)
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
                .MenuItems.Add tsMenuCaption, "ÿ»«⁄… ‘Ã—… «·√’‰«›", False, False, 10, , , True, , "PrintGroup", , , , "ÿ»«⁄…"
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
                .MenuItems.Add tsMenuCaption, "Print Items Tree", False, False, 10, , , True, , "PrintGroup", , , , "Print all of the tree of items"
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
            XPCboGroup.BoundText = left(XNode.key, Len(XNode.key) - 1)

        Case "EditGroup"
            Cmd_Click (1)

        Case "DelGroup"
            Cmd_Click (4)

        Case "ClearGroup"

        Case "CutGroup"
            TreeGroups.SelectedItem.backcolor = vbGreen
            TxtCutKey.text = (TreeGroups.SelectedItem.key)

            '        Me.TxtMenuState.Text = "C"
        Case "PasteGroup"
            TreeGroups.Nodes.Remove (TxtCutKey.text)
            Set XNode = TreeGroups.Nodes.Add(Trim(TreeGroups.SelectedItem.key), tvwChild, rs("GroupID") & "G", rs("GroupName"), "Closed_Node", "Open_Node")
            StrSQL = "update Groups set ParentID=" & val(left(TreeGroups.SelectedItem.key, Len(TreeGroups.SelectedItem.key) - 1)) & " where GroupID=" & val(rs("GroupID").value)
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
    lbl(18).Caption = "Validitation"
    Me.Caption = "Items Groups"
    Me.LblHeader.Caption = Me.Caption
    Chklast.Caption = "Last Group"
    Me.lbl(6).Caption = "Group ID"
    Me.lbl(17).Caption = "Group Code"
    Me.lbl(5).Caption = "Group Name"
    Me.lbl(0).Caption = "Parent Group"
    Me.lbl(1).Caption = "Current Record"
    Me.lbl(2).Caption = "NO. Recordes"

    Ele.Caption = "More Information"
    Ele.Font.Bold = True
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
    StrSQL = "Select GroupCode From Groups Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From Groups Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
        IntTemp = val(Mid(StrLastGroupCode, Len(StrParentCode) + 1))
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
    StrSQL = "SELECT Count(Groups.GroupID) AS CountGroupID"
    StrSQL = StrSQL + " From Groups WHERE (Groups.ParentID=" & Lngid & ")"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Me.lbl(10).Caption = "0"
    Else
        Me.lbl(10).Caption = IIf(IsNull(rs("CountGroupID").value), 0, rs("CountGroupID").value)
    End If

    rs.Close
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

