VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEmpSalaryComponentIncres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… «·“Ì«œ« "
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13650
   Icon            =   "frmEmpContractIncrese.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   13650
   Begin VB.TextBox TXTTotal 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Text            =   "0"
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox Contract_ID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Text            =   "Text3"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Text            =   "TxtModFlg"
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox XPTxtEmpName 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox emp_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox emp_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox emp_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox emp_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Basic_salary 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox test_period_no 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox test_period 
      Height          =   315
      ItemData        =   "frmEmpContractIncrese.frx":000C
      Left            =   240
      List            =   "frmEmpContractIncrese.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox emp_code 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13605
      _cx             =   23998
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
      Caption         =   "  «·“Ì«œ«   "
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
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
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
         ButtonImage     =   "frmEmpContractIncrese.frx":0024
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
         TabIndex        =   2
         Top             =   90
         Visible         =   0   'False
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
         ButtonImage     =   "frmEmpContractIncrese.frx":03BE
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
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
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
         ButtonImage     =   "frmEmpContractIncrese.frx":0758
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
         TabIndex        =   4
         Top             =   90
         Visible         =   0   'False
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
         ButtonImage     =   "frmEmpContractIncrese.frx":0AF2
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   4860
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   13440
      _cx             =   23707
      _cy             =   8572
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
      Rows            =   10
      Cols            =   17
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEmpContractIncrese.frx":0E8C
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
   End
   Begin MSComCtl2.DTPicker Issue_date 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   76021761
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   10980
      TabIndex        =   17
      Top             =   8670
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   18
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   2
      Left            =   7200
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   3
      Left            =   5880
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   4
      Left            =   4680
      TabIndex        =   21
      Top             =   8040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   6
      Left            =   3240
      TabIndex        =   22
      Top             =   8670
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   8670
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
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
      Height          =   495
      Index           =   5
      Left            =   5790
      TabIndex        =   24
      Top             =   8670
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
   Begin MSDataListLib.DataCombo Departement 
      Height          =   315
      Left            =   3360
      TabIndex        =   27
      Top             =   1440
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo job 
      Height          =   315
      Left            =   9840
      TabIndex        =   28
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   9
      Left            =   12000
      TabIndex        =   42
      Top             =   7320
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–› ”ÿ—"
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
      ButtonImage     =   "frmEmpContractIncrese.frx":1153
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   7
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–› ”ÿ—"
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
      ButtonImage     =   "frmEmpContractIncrese.frx":16ED
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„›—œ«  «· Ì ÌÕ’· ⁄·ÌÂ« «·„ÊŸ›"
      Height          =   375
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  «·“Ì«œ«   "
      Height          =   375
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
      Height          =   375
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄œœ «·⁄ﬁÊœ"
      Height          =   255
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄ﬁœ «·Õ«·Ì"
      Height          =   255
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      Caption         =   "Label19"
      Height          =   255
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      Caption         =   "XPTxtCurrent"
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—« » «·«”«”Ì"
      Height          =   375
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   " «—ÌŒ „»«‘—… «·⁄„·"
      Height          =   375
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬁ”„"
      Height          =   255
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ÊŸÌ›…"
      Height          =   375
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÂ «·«Œ »«—"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   375
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·„ÊŸ›"
      Height          =   375
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   -120
      Width           =   1695
   End
End
Attribute VB_Name = "frmEmpSalaryComponentIncres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim EmpReport As ClsEmployeeReport
Dim xReport As New CRAXDRT.Report
Dim NO As Double
 Public LngRow  As Double
Public LngCol  As Double
                     
Private objScript As Object
Dim case_id As Integer
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
   
Public Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index
  
        Case 1
    
            TxtModFlg.text = "E"
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1

        Case 2
     
            SaveData
            '        FrmEmployee.updateResults
     
        Case 3
            Undo

        Case 4
  
            Del_ProfData

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If
        
        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

        Case 9

            With Me.VSFlexGrid1
  
                If .Row <= 0 Then Exit Sub
                .RemoveItem .Row
                Me.TxtTotal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows - 1, .ColIndex("value"))
            End With

            ReLineGrid
         
    End Select

    Exit Sub
ErrTrap:

End Sub
 
Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Public Function get_value(operand As String) As Double
    operand = Replace$(operand, "A", "")
    Dim value As Double
    Dim mofrad_count As Integer
    mofrad_count = 0

'    If operand = 1 Then
'        If IsNumeric(Basic_salary.text) Then
'            get_value = 1 * val(Basic_salary.text)
'            Exit Function
'        Else
'            get_value = 0
'            MsgBox "·„ Ì „  ÕœÌœ ﬁÌ„… «·—« » «·«”«”Ì »—Ã«¡  ÕœÌœÂ«"
'            Exit Function
'        End If
'
'    End If

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) = operand Then
                mofrad_count = mofrad_count + 1
              
            End If
        
        Next i

    End With

    If mofrad_count = 0 Then
        MsgBox "«·„›—œ €Ì— „ÊÃÊœ"
        Exit Function
    ElseIf mofrad_count > 1 Then
        MsgBox "«·„›—œ    „Õœœ «ﬂÀ— „‰ „—…"
        Exit Function
    End If

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) = operand Then
                get_value = .TextMatrix(i, .ColIndex("value"))
                Exit Function
            End If
        
        Next i

    End With
 
End Function

Public Function cal_value(src As String) As Double
    On Error GoTo errortrap
    Dim new_pos As Integer
    Dim last_pos As Integer
    Dim cuttent_operand As String
    Dim new_str As String
    Dim objScript As Object
    last_pos = 1
    new_str = ""

    For i = 1 To Len(src)

        If Mid(src, i, 1) = "+" Or Mid(src, i, 1) = "-" Or Mid(src, i, 1) = "*" Or Mid(src, i, 1) = "/" Or Mid(src, i, 1) = "=" Then
            new_pos = i
            cuttent_operand = Mid(src, last_pos, new_pos - last_pos)

            If InStr(cuttent_operand, "A") > 0 Then
                cuttent_operand = get_value(cuttent_operand)
            End If

            new_str = new_str & cuttent_operand & Mid(src, i, 1)

            If i < Len(src) Then
                last_pos = new_pos + 1
            Else
                GoTo LL
            End If
        End If
 
    Next i

LL:
    new_str = Replace$(new_str, "=", "")

    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
 
    cal_value = objScript.Eval(new_str)
    cal_value = Round(cal_value, 2)
    Exit Function
errortrap:
    cal_value = 0

End Function

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "AccountName"
                         '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
             
                StrSQL = " SELECT     *, dbo.mofrad.name, dbo.mofrad.nameE, dbo.mofrad.AddOrDiscount, dbo.mofrad.id"
                StrSQL = StrSQL & " FROM         dbo.mofrdat INNER JOIN"
                StrSQL = StrSQL & "       dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.id "
                StrSQL = StrSQL & "         Where mofrad_code = " & StrAccountCode
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("eq_sys").value), "", rs("eq_sys").value)
                    
                    .TextMatrix(Row, .ColIndex("eq_text")) = IIf(IsNull(rs("eq_text").value), "", rs("eq_text").value)
                    .TextMatrix(Row, .ColIndex("mofrad_type")) = IIf(IsNull(rs("mofrad_type").value), "", rs("mofrad_type").value)
                    .TextMatrix(Row, .ColIndex("AddOrDiscount")) = IIf(IsNull(rs("AddOrDiscount").value), "", rs("AddOrDiscount").value)
                    
                    .TextMatrix(Row, .ColIndex("specific_value")) = IIf(IsNull(rs("specific_value").value), "", rs("specific_value").value)
                    .TextMatrix(Row, .ColIndex("assurance")) = IIf(IsNull(rs("assurance").value), "", rs("assurance").value)
                    .TextMatrix(Row, .ColIndex("percentage")) = IIf(IsNull(rs("percentage").value), "", rs("percentage").value)
                    .TextMatrix(Row, .ColIndex("min_val")) = IIf(IsNull(rs("min_val").value), "", rs("min_val").value)
                    .TextMatrix(Row, .ColIndex("max_val")) = IIf(IsNull(rs("max_val").value), "", rs("max_val").value)
                    .TextMatrix(Row, .ColIndex("is_fixed")) = IIf(IsNull(rs("is_fixed").value), "", rs("is_fixed").value)
                    .TextMatrix(Row, .ColIndex("Monthly")) = IIf(IsNull(rs("Monthly").value), "", rs("Monthly").value)
                   
                End If

                calcnets
           
            Case "value"
                Dim sgl As String
                
                Me.TxtTotal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows - 1, .ColIndex("value"))
                '  Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        calcnets
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.TxtTotal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows - 1, .ColIndex("value"))
    
        '    Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
ErrTrap:
End Sub

Function calcnets()
    Dim Row As Integer

    With VSFlexGrid1

        For Row = 1 To VSFlexGrid1.Rows - 1

            If .TextMatrix(Row, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(Row, .ColIndex("is_fixed"))) = 1 Then
                    .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("specific_value")))
                ElseIf val(.TextMatrix(Row, .ColIndex("is_fixed"))) = 0 Then
                    .TextMatrix(Row, .ColIndex("value")) = cal_value(.TextMatrix(Row, .ColIndex("eq_text")))
                ElseIf val(.TextMatrix(Row, .ColIndex("is_fixed"))) = 2 Then
                    .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("value")))
           
                End If
           
                If val(.TextMatrix(Row, .ColIndex("value"))) < val(.TextMatrix(Row, .ColIndex("min_val"))) And val(.TextMatrix(Row, .ColIndex("min_val"))) > 0 Then
                    .TextMatrix(Row, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("min_val"))
                ElseIf val(.TextMatrix(Row, .ColIndex("value"))) > val(.TextMatrix(Row, .ColIndex("max_val"))) And val(.TextMatrix(Row, .ColIndex("max_val"))) > 0 Then
                    .TextMatrix(Row, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("max_val"))
                End If
           
                If val(.TextMatrix(Row, .ColIndex("AccountCode"))) = 1 Then
                    '          .TextMatrix(Row, .ColIndex("value")) = val(Basic_salary.text)
                End If
            End If

        Next Row

    End With

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
Dim totalvalue As Double
    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                    If .TextMatrix(i, .ColIndex("AddOrDiscount")) = False Then
                     totalvalue = totalvalue + val(.TextMatrix(i, .ColIndex("value")))
    
                .Cell(flexcpBackColor, i, 1, i, 14) = &H80FF80
    
            
                    Else
                    .Cell(flexcpBackColor, i, 1, i, 14) = &H8080FF
                    totalvalue = totalvalue - val(.TextMatrix(i, .ColIndex("value")))
                    End If
            
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                
                TxtTotal.text = totalvalue
            End If

        Next i

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

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
   Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim rdate As Date
  ' Dim frm As FrmGridAddItemComment
    Dim Frm1 As FromRegisterEmpDateIncrease

    'On Error GoTo ErrTrap

    With Me.VSFlexGrid1

        Select Case .ColKey(Col)

                 Case "EntIncresDataM"
                  LngRow = Row
'
                     LngCol = Col
            
                Load FromRegisterEmpDateIncrease
                FromRegisterEmpDateIncrease.show
                  Case "EntIncresDataH"
                                    LngRow = Row

 LngCol = Col
                 Load FromRegisterEmpDateIncrease
               FromRegisterEmpDateIncrease.show
                    
                End Select
                End With
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With VSFlexGrid1

        Select Case .ColKey(Col)
Case "EntIncresDataM"
            .ColComboList(.ColIndex("EntIncresDataM")) = "..."
               Case "EntIncresDataH"
            .ColComboList(.ColIndex("EntIncresDataH")) = "..."
            Case "AccountName"
                'Full Path Display
                 
                StrSQL = " select * from mofrdat "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "eq_sys, *mofrad_name", "mofrad_code")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "eq_sys, *mofrad_namee", "mofrad_code")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_Load()
    system_path = App.path
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    Dim Msg As String
    Set Dcombos = New ClsDataCombos

    Dcombos.GetEmpDepartments Me.Departement
    Dcombos.GetEmpJobsTypes Me.job

    'On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
 
    End If

    '
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Set Dcombos = New ClsDataCombos
 
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    rs.Open "TblEmployee", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    'If OPEN_NEW_SCREEN = True Then
    'Cmd_Click (0)
    'End If

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
    Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    'Exit Sub
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·“Ì«œ«   "
            Else
                Me.Caption = "Increases Data"
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

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·⁄ﬁÊœ ( ”ÃÌ· ”Ã· ÃœÌœ)"
            Else
                Me.Caption = "Contract  Data(Enter New Record)"
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
          
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "„›—œ«  «·„ÊŸ›  (  ⁄œÌ· )"
            Else
                Me.Caption = "Contarct Data(Edit Current Record)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False

    End Select

    Exit Sub
ErrTrap:
End Sub
 
Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
    Else

        If Lngid <> 0 Then
            rs.find "emp_id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
                clear_all Me
                VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                VSFlexGrid1.Rows = 1
            End If
        End If
    End If

    Me.job.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
    Me.Departement.BoundText = IIf(IsNull(rs("DepartmentID").value), "", rs("DepartmentID").value)
    'Contract_ID.text = IIf(IsNull(rs("Contract_ID").value), "", Val(rs("Contract_ID").value))
    Basic_salary.text = IIf(IsNull(rs("Emp_Salary").value), "", rs("Emp_Salary").value)
    Emp_id.text = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
    emp_code.text = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)

    XPTxtEmpName.text = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))

    If SystemOptions.UserInterface = ArabicInterface Then
        emp_name(0).text = IIf(IsNull(rs("Emp_Name1").value), "", Trim(rs("Emp_Name1").value))
        emp_name(1).text = IIf(IsNull(rs("Emp_Name2").value), "", Trim(rs("Emp_Name2").value))
        emp_name(2).text = IIf(IsNull(rs("Emp_Name3").value), "", Trim(rs("Emp_Name3").value))
        emp_name(3).text = IIf(IsNull(rs("Emp_Name4").value), "", Trim(rs("Emp_Name4").value))
    Else
        emp_name(0).text = IIf(IsNull(rs("Emp_Namee1").value), "", Trim(rs("Emp_Namee1").value))
        emp_name(1).text = IIf(IsNull(rs("Emp_Namee2").value), "", Trim(rs("Emp_Namee2").value))
        emp_name(2).text = IIf(IsNull(rs("Emp_Namee3").value), "", Trim(rs("Emp_Namee3").value))
        emp_name(3).text = IIf(IsNull(rs("Emp_Namee4").value), "", Trim(rs("Emp_Namee4").value))

    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
          
    Dim rscomponent As ADODB.Recordset
    Dim sql As String

   
   
   
sql = " SELECT     TOP 100 PERCENT dbo.EmpSalaryComponent.id, dbo.EmpSalaryComponent.Contract_ID, dbo.EmpSalaryComponent.AccountCode, dbo.EmpSalaryComponent.emp_ID,"
sql = sql & "  dbo.EmpSalaryComponent.AccountName, dbo.EmpSalaryComponent.[Value], dbo.EmpSalaryComponent.des, dbo.EmpSalaryComponent.eq_text,"
sql = sql & "  dbo.EmpSalaryComponent.specific_value, dbo.EmpSalaryComponent.assurance, dbo.EmpSalaryComponent.percentage, dbo.EmpSalaryComponent.min_val,"
sql = sql & " dbo.EmpSalaryComponent.max_val, dbo.EmpSalaryComponent.is_fixed, dbo.EmpSalaryComponent.Monthly, dbo.EmpSalaryComponent.mofrad_type,"
sql = sql & " dbo.EmpSalaryComponent.ModDate, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.EmpSalaryComponent.Flagx,"
sql = sql & " dbo.EmpSalaryComponent.EntIncresDataM , dbo.EmpSalaryComponent.EntIncresDataH, dbo.MOFRAD.AddOrDiscount"
sql = sql & " FROM         dbo.mofrdat INNER JOIN"
sql = sql & " dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.id RIGHT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
'Where (dbo.EmpSalaryComponent.Emp_id = 2) And (dbo.EmpSalaryComponent.Flagx = 1)
sql = sql & " Where (dbo.EmpSalaryComponent.Flagx = 1) And (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
    Set rscomponent = New ADODB.Recordset
    rscomponent.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rscomponent.RecordCount > 0 Then

        With Me.VSFlexGrid1
            .Rows = .FixedRows + rscomponent.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rscomponent("AccountCode").value), "", rscomponent("AccountCode").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rscomponent("mofrad_name").value), "", rscomponent("mofrad_name").value)
                Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rscomponent("mofrad_namee").value), "", rscomponent("mofrad_namee").value)
                End If
            .TextMatrix(i, .ColIndex("AddOrDiscount")) = IIf(IsNull(rscomponent("AddOrDiscount").value), 0, rscomponent("AddOrDiscount").value)
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rscomponent("value").value), 0, rscomponent("value").value)
               .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rscomponent("des").value), "", rscomponent("des").value)
                .TextMatrix(i, .ColIndex("eq_text")) = IIf(IsNull(rscomponent("eq_text").value), "", rscomponent("eq_text").value)
                .TextMatrix(i, .ColIndex("specific_value")) = IIf(IsNull(rscomponent("specific_value").value), 0, rscomponent("specific_value").value)
              .TextMatrix(i, .ColIndex("assurance")) = IIf(IsNull(rscomponent("assurance").value), 0, rscomponent("assurance").value) 'rscomponent("assurance").value
           
                .TextMatrix(i, .ColIndex("percentage")) = IIf(IsNull(rscomponent("percentage").value), 0, rscomponent("percentage").value)
                .TextMatrix(i, .ColIndex("min_val")) = IIf(IsNull(rscomponent("min_val").value), 0, rscomponent("min_val").value)
           
                .TextMatrix(i, .ColIndex("max_val")) = IIf(IsNull(rscomponent("max_val").value), 0, rscomponent("max_val").value)
           '
                .TextMatrix(i, .ColIndex("is_fixed")) = IIf(IsNull(rscomponent("is_fixed").value), 0, rscomponent("is_fixed").value) ' rscomponent("is_fixed").value
           
               .TextMatrix(i, .ColIndex("Monthly")) = IIf(IsNull(rscomponent("Monthly").value), 0, rscomponent("Monthly").value) 'rscomponent("Monthly").value
                .TextMatrix(i, .ColIndex("mofrad_type")) = IIf(IsNull(rscomponent("mofrad_type").value), 0, rscomponent("mofrad_type").value)
            If Not (IsNull(rscomponent("EntIncresDataM").value)) Then
                    .TextMatrix(i, .ColIndex("EntIncresDataM")) = Format(rscomponent("EntIncresDataM").value, "yyyy/M/d")
                   
                End If
                .TextMatrix(i, .ColIndex("EntIncresDataH")) = IIf(IsNull(rscomponent("EntIncresDataH").value), "", rscomponent("EntIncresDataH").value)
                    
            
                
                rscomponent.MoveNext
            Next

            Me.TxtTotal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows - 1, .ColIndex("value"))
        End With

    End If
ReLineGrid
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

  '  If Not IsNumeric(Basic_salary.text) Then
  '      Msg = "ÌÃ» «œŒ«· «·—« » «·«”«”Ì ··„ÊŸ›  "
  '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  '      Basic_salary.SetFocus
  '      SelectText Basic_salary
  '      Exit Sub
  '  End If
    
    If val(Emp_id.text) = 0 Then
        Msg = "Â–« «·„ÊŸ› €Ì— „ÊÃÊœ  "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    calcnets

    sql = "delete    EmpSalaryComponent where emp_ID=" & val(Emp_id.text) & " And (flagx = 1)"
    Cn.Execute sql

    Dim rscomponent As ADODB.Recordset

  '  sql = "EmpSalaryComponent"
    Set rscomponent = New ADODB.Recordset
   ' rscomponent.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdTable
     sql = "SELECT     * from dbo.EmpSalaryComponent Where (1 = -1)"
   rscomponent.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 2
       If .TextMatrix(i, .ColIndex("EntIncresDataH")) <> "" Then
            rscomponent.AddNew
            rscomponent("Contract_ID").value = val(Contract_ID.text)
            rscomponent("emp_ID").value = val(Emp_id.text) '„—»Êÿ »—ﬁ„ «·„ÊŸ›
            rscomponent("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", "", .TextMatrix(i, .ColIndex("AccountCode")))
            rscomponent("AccountName").value = IIf((.TextMatrix(i, .ColIndex("AccountName"))) = "", "", .TextMatrix(i, .ColIndex("AccountName")))
            rscomponent("value").value = IIf((.TextMatrix(i, .ColIndex("value"))) = "", 0, .TextMatrix(i, .ColIndex("value")))
            rscomponent("des").value = IIf((.TextMatrix(i, .ColIndex("des"))) = "", "", .TextMatrix(i, .ColIndex("des")))
            rscomponent("eq_text").value = IIf((.TextMatrix(i, .ColIndex("eq_text"))) = "", "", .TextMatrix(i, .ColIndex("eq_text")))
            rscomponent("specific_value").value = IIf((.TextMatrix(i, .ColIndex("specific_value"))) = "", 0, .TextMatrix(i, .ColIndex("specific_value")))
            rscomponent("assurance").value = .TextMatrix(i, .ColIndex("assurance"))
            rscomponent("percentage").value = IIf((.TextMatrix(i, .ColIndex("percentage"))) = "", 0, .TextMatrix(i, .ColIndex("percentage")))
            rscomponent("min_val").value = IIf((.TextMatrix(i, .ColIndex("min_val"))) = "", 0, .TextMatrix(i, .ColIndex("min_val")))
            rscomponent("max_val").value = IIf((.TextMatrix(i, .ColIndex("max_val"))) = "", 0, .TextMatrix(i, .ColIndex("max_val")))
            rscomponent("is_fixed").value = IIf((.TextMatrix(i, .ColIndex("is_fixed"))) = "", 0, .TextMatrix(i, .ColIndex("is_fixed")))
            rscomponent("Monthly").value = IIf((.TextMatrix(i, .ColIndex("Monthly"))) = "", 0, .TextMatrix(i, .ColIndex("Monthly")))
            rscomponent("mofrad_type").value = IIf((.TextMatrix(i, .ColIndex("mofrad_type"))) = "", 0, .TextMatrix(i, .ColIndex("mofrad_type")))
             rscomponent("EntIncresDataM").value = IIf(IsDate(.TextMatrix(i, .ColIndex("EntIncresDataM"))), .TextMatrix(i, .ColIndex("EntIncresDataM")), Date)
              rscomponent("EntIncresDataH").value = IIf((.TextMatrix(i, .ColIndex("EntIncresDataH"))) = "", "", .TextMatrix(i, .ColIndex("EntIncresDataH")))
            rscomponent("Flagx").value = 1
            rscomponent.update
            Else
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ œÊ‰  «—ÌŒ «·“Ì«œ…"
             Exit Sub
            End If
        Next i

    End With

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

    End If
 
    TxtModFlg.text = "R"
    'Retrive val(Me.Emp_id.text)
    'addSalaryComponentToEmployee val(Me.Emp_id.text)
    Unload Me
    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        Cn.RollbackTrans
        BeginTrans = False
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"

        Case "E"
     
            Retrive val(Emp_id.text)
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_ProfData()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Contract_ID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„ÊŸ› —ﬁ„ " & Chr(13)
        Msg = Msg + (Contract_ID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
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

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·⁄ﬁœ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            KeyCode = 0
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

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then
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
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄ﬁœ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄ﬁÊœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄ﬁÊœ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› «·⁄ﬁÊœ „ÊŸ›" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄ﬁœ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄ﬁÊœ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record ..." & Wrap & "Click here to add a new Contract" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the current Contract data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or " & Wrap & "save the edit in the " & Wrap & "current record", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current Contract data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an Contract" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Contract Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
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

Private Sub ChangeLang()

    Me.Caption = "Salary Incres Component"
    EleHeader.Caption = Me.Caption

    Label3.Caption = "Contract #"
    Label5.Caption = "Component"
    'Label8.Caption = "Type"
    'Label9.Caption = "Contract period"
    Label10.Caption = "Exam period"

    Label6.Caption = "Emp Code"
    Label7.Caption = "Emp Name"
    Label11.Caption = "Job"
    Label12.Caption = "Departement"
    Label13.Caption = "Start date"
    Label1.Caption = "Total"
    Cmd(9).Caption = "Remove Line"
    Label18.Caption = "Basic Salary"

    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
        .TextMatrix(0, .ColIndex("AccountName")) = "Breakdown of Remuneration "
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("des")) = "Equation"
        .TextMatrix(0, .ColIndex("Monthly")) = "Monthly"
        .TextMatrix(0, .ColIndex("specific_value")) = "specific value"
        .TextMatrix(0, .ColIndex("min_val")) = "min_val"
        .TextMatrix(0, .ColIndex("max_val")) = "max_val"
        .TextMatrix(0, .ColIndex("EntIncresDataM")) = "Entitlement AD"
        .TextMatrix(0, .ColIndex("EntIncresDataH")) = "Entitlement Hegira"
        .TextMatrix(0, .ColIndex("percentage")) = "Insurance%"
    End With
 
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
End Sub
 
