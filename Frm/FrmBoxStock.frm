VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBoxStock 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ «·Œ“‰…"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "FrmBoxStock.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Caption         =   " À»Ì  Ê⁄—÷ Ã„Ì⁄ «·›∆« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Tag             =   "not"
      Top             =   4530
      Width           =   2265
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ≈÷«›Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1005
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   660
      Width           =   2835
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—’Ìœ «·Œ“‰… Õ Ï «· «—ÌŒ «·„Õœœ"
         Height          =   405
         Index           =   8
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   510
         Width           =   1425
      End
      Begin VB.Label LblBoxName 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label LblBoxAccount 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtRemarks 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   570
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   4920
      Width           =   5175
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2910
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   645
      Width           =   1545
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   540
      TabIndex        =   0
      Top             =   1740
      Width           =   5805
      _cx             =   10239
      _cy             =   4842
      Appearance      =   2
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBoxStock.frx":038A
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
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6375
      _cx             =   11245
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "Ã—œ «·Œ“‰…"
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
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   90
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
         ButtonImage     =   "FrmBoxStock.frx":046F
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
         Top             =   90
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
         ButtonImage     =   "FrmBoxStock.frx":0809
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
         Top             =   90
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
         ButtonImage     =   "FrmBoxStock.frx":0BA3
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
         Top             =   90
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
         ButtonImage     =   "FrmBoxStock.frx":0F3D
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   1365
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   345
      Left            =   3840
      TabIndex        =   8
      Top             =   1005
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      _Version        =   393216
      Format          =   84541441
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   12
      Top             =   6576
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
      Left            =   4770
      TabIndex        =   13
      Top             =   6576
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   4035
      TabIndex        =   14
      Top             =   6576
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
      Left            =   3285
      TabIndex        =   15
      Top             =   6576
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   2505
      TabIndex        =   16
      Top             =   6576
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
      Height          =   372
      Index           =   6
      Left            =   96
      TabIndex        =   17
      Top             =   6576
      Width           =   732
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
      Left            =   4200
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   792
      _ExtentX        =   1402
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
      Left            =   1680
      TabIndex        =   19
      Top             =   6576
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   2580
      TabIndex        =   25
      Top             =   5760
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton XPBtnAdd 
      Height          =   315
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2010
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmBoxStock.frx":12D7
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      LowerToggledContent=   0   'False
   End
   Begin ImpulseButton.ISButton XPBtnRemove 
      Height          =   315
      Left            =   120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2370
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmBoxStock.frx":1671
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      LowerToggledContent=   0   'False
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   2520
      TabIndex        =   38
      Top             =   7080
      Width           =   1155
      _ExtentX        =   2037
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
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   795
      _ExtentX        =   1402
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
      Height          =   372
      Index           =   8
      Left            =   840
      TabIndex        =   40
      Top             =   6576
      Width           =   852
      _ExtentX        =   1508
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   5760
      Width           =   1755
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   5
      Left            =   5790
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4950
      Width           =   555
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   315
      Index           =   1
      Left            =   1860
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5760
      Width           =   675
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   285
      Index           =   6
      Left            =   5370
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5790
      Width           =   945
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   6180
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1380
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6180
      Width           =   705
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   5100
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6180
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6180
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   0
      Left            =   5430
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   660
      Width           =   915
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   255
      Index           =   3
      Left            =   5430
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·⁄„·Ì…"
      Height          =   315
      Index           =   7
      Left            =   5430
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1005
      Width           =   915
   End
End
Attribute VB_Name = "FrmBoxStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch

Private Enum GridDisplayMode
    ShowNonBands
    ShowAllBands
End Enum
Dim m_GridDisplayMode As GridDisplayMode

Private Sub Chk_Click()
    Dim Msg As String

    If Me.Chk.value = vbChecked Then
        m_GridDisplayMode = ShowAllBands
    ElseIf Me.Chk.value = vbUnchecked Then
        m_GridDisplayMode = ShowNonBands
    End If

    SetupGrid m_GridDisplayMode

    If Me.TxtModFlg.text = "R" Then
        Retrive val(Me.XPTxtID.text)
    End If

End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"

            If Me.Chk.value = vbChecked Then
                m_GridDisplayMode = ShowAllBands
            ElseIf Me.Chk.value = vbUnchecked Then
                m_GridDisplayMode = ShowNonBands
            End If

            SetupGrid m_GridDisplayMode
            XPTxtID.text = CStr(new_id("TblBoxStock", "BoxStockID", "", True))
            Me.DCboUserName.BoundText = user_id
            ' XPDtbTrans.SetFocus
            'Me.Fg.Rows = Me.Fg.FixedRows + 1
            Me.LblBoxName = ""
            Me.LblBoxAccount.Caption = 0

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata
        
        Case 2
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

            Load FrmBoxStockSearch
            FrmBoxStockSearch.show vbModal

        Case 6
            Unload Me
            
        Case 8
               print_report
    End Select

    Exit Sub
ErrTrap:
End Sub


Private Sub print_report()

     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT dbo.TblBoxStock.BoxStockDate, dbo.TblBoxStockDetails.CurrencyBandID, dbo.TblBoxStock.BoxStockID, dbo.TblCurrencyBandNames.CurrencyBandName,"
 MySQL = MySQL + "                 dbo.TblCurrencyBandNames.CurrenyBandValue, dbo.TblBoxStockDetails.CurrencyBandCount, dbo.TblBoxStock.BoxID, dbo.TblBoxStock.Remarks,"
MySQL = MySQL + "                        dbo.TblBoxesData.Boxname"
MySQL = MySQL + "      FROM     dbo.TblBoxStock INNER JOIN"
 MySQL = MySQL + "                       dbo.TblBoxStockDetails ON dbo.TblBoxStock.BoxStockID = dbo.TblBoxStockDetails.BoxStockID INNER JOIN"
  MySQL = MySQL + "                      dbo.TblCurrencyBandNames ON dbo.TblBoxStockDetails.CurrencyBandID = dbo.TblCurrencyBandNames.CurrencyBandID LEFT OUTER JOIN"
 MySQL = MySQL + "                       dbo.TblBoxesData ON dbo.TblBoxStock.BoxID = dbo.TblBoxesData.BoxID"
    
   MySQL = MySQL + " where     TblBoxStock.BoxStockID = " & val(XPTxtID.text)
    
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & Report_Folder & "\rpt_BoxStock.rpt"
    Else
            StrFileName = App.path & "\Special\" & Report_Folder & "\rpt_BoxStock.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
            
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 



End Sub


Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (Me.XPTxtID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                    'GetBoxData
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim strsql As String
    Dim BeginTrans As Boolean
    Dim IntRes As Integer
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Trim(Me.DcboBox.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If CheckBoxStockDate(Me.DcboBox.BoundText, Me.XPDtbTrans.value) = False Then
            Exit Sub
        End If

        If val(Me.LblBoxAccount.Caption) <> val(Me.LblTotal.Caption) Then
            Msg = "«‰ »Â ..."
            Msg = Msg & Chr(13) & "ÌÃ» „·«ÕŸ… «‰ ﬁÌ„… «·Ã—œ «·„œŒ·… «·¬‰ " & val(Me.LblTotal.Caption)
            Msg = Msg & Chr(13) & "·«  ”«ÊÏ „⁄ ﬁÌ„… —’Ìœ «·Œ“‰… ›Ï Â–« «·ÌÊ„ " & val(Me.LblBoxAccount.Caption)
            Msg = Msg & Chr(13) & " Ê·ﬂ‰ Ì„ﬂ‰ﬂ Õ›Ÿ ⁄„·Ì… «·Ã—œ Â–Â ›Â·  —Ìœ «·„ «»⁄… ..øøø"
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

            If IntRes = vbNo Then
                Exit Sub
            End If
        End If

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            strsql = "Delete From TblBoxStockDetails Where BoxStockID=" & val(Me.XPTxtID.text) & ""
            Cn.Execute strsql
        End If

        rs("BoxStockID").value = val(XPTxtID.text)
        rs("BoxStockDate").value = XPDtbTrans.value
        rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
        rs("Remarks").value = IIf(Me.TxtRemarks.text = "", Null, Trim(TxtRemarks.text))
        rs("UserID").value = Me.DCboUserName.BoundText
        rs.update
        Set RsTemp = New ADODB.Recordset
        'RsTemp.Open "TblBoxStockDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            strsql = "SELECT     * from dbo.TblBoxStockDetails Where (1 = -1)"
   RsTemp.Open strsql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
     
        With Me.FG

            For i = .FixedRows To .Rows - 1

                If val(.TextMatrix(i, .ColIndex("CurrencyBandCount"))) > 0 Then
                    RsTemp.AddNew
                    RsTemp("BoxStockID").value = val(XPTxtID.text)
                    RsTemp("CurrencyBandID").value = val(.TextMatrix(i, .ColIndex("CurrenyBandName")))
                    RsTemp("CurrencyBandCount").value = val(.TextMatrix(i, .ColIndex("CurrencyBandCount")))
                    RsTemp.update
                End If

            Next i

        End With

        Cn.CommitTrans
        BeginTrans = False
        'GetBoxData
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
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
            rs.find "BoxStockID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim LngFindRow As Long

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
            rs.find "BoxStockID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("BoxStockID").value), "", val(rs("BoxStockID").value))
    TxtRemarks.text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))
    XPDtbTrans.value = IIf(IsNull(rs("BoxStockDate").value), Date, rs("BoxStockDate").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    GetBoxData
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Set RsTemp = New ADODB.Recordset
    strsql = "SELECT BoxStockID, TblBoxStockDetails.CurrencyBandID, CurrencyBandCount, TableID" & ",TblCurrencyBandNames.CurrenyBandValue" & " FROM TblBoxStockDetails INNER JOIN " & " TblCurrencyBandNames ON TblBoxStockDetails.CurrencyBandID = " & "TblCurrencyBandNames.CurrencyBandID "
    strsql = strsql + " Where BoxStockID=" & val(XPTxtID.text) & ""
    strsql = strsql + " Order BY TableID"
    RsTemp.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.BOF Or RsTemp.EOF) Then
        SetupGrid m_GridDisplayMode

        If m_GridDisplayMode = ShowNonBands Then

            With Me.FG
                .Rows = .FixedRows + RsTemp.RecordCount

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("CurrenyBandName")) = IIf(IsNull(RsTemp("CurrencyBandID").value), "", RsTemp("CurrencyBandID").value)
                    .TextMatrix(i, .ColIndex("CurrencyBandCount")) = IIf(IsNull(RsTemp("CurrencyBandCount").value), "", RsTemp("CurrencyBandCount").value)
                    .TextMatrix(i, .ColIndex("CurrenyBandValue")) = IIf(IsNull(RsTemp("CurrenyBandValue").value), "", RsTemp("CurrenyBandValue").value)
                    RsTemp.MoveNext
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        ElseIf m_GridDisplayMode = ShowAllBands Then

            With Me.FG

                For i = 1 To RsTemp.RecordCount
                    LngFindRow = .FindRow(RsTemp("CurrencyBandID").value, .FixedRows, .ColIndex("CurrenyBandName"), False, True)

                    If LngFindRow <> -1 Then
                        .TextMatrix(LngFindRow, .ColIndex("CurrenyBandName")) = IIf(IsNull(RsTemp("CurrencyBandID").value), "", RsTemp("CurrencyBandID").value)
                        .TextMatrix(LngFindRow, .ColIndex("CurrencyBandCount")) = IIf(IsNull(RsTemp("CurrencyBandCount").value), "", RsTemp("CurrencyBandCount").value)
                        .TextMatrix(LngFindRow, .ColIndex("CurrenyBandValue")) = IIf(IsNull(RsTemp("CurrenyBandValue").value), "", RsTemp("CurrenyBandValue").value)
                    End If

                    RsTemp.MoveNext
                Next i

            End With

        End If

        CalculateGrid
    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments XPTxtID, "0812201403"

End Sub

Private Sub DcboBox_Change()
    GetBoxData
End Sub

Private Sub DcboBox_Click(Area As Integer)
    GetBoxData
End Sub

Private Sub GetBoxData()
    Dim DblExistValue As Double

    If Me.DcboBox.BoundText = "" Then Exit Sub

    Me.LblBoxName = Me.DcboBox.text
    Me.LblBoxAccount.Caption = get_balanceFromGl(ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText)))

    Exit Sub
    Me.LblBoxName = Me.DcboBox.text

    If CheckBoxAccount(val(Me.DcboBox.BoundText), 0, Me.XPDtbTrans.value, False, DblExistValue) = True Then
        Me.LblBoxAccount.Caption = DblExistValue
    Else
        Me.LblBoxAccount.Caption = 0
    End If

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Dim LngFoundRow As Long
    Dim LngNewID As Long
    Dim Msg As String

    With Me.FG

        Select Case .ColKey(Col)

            Case "CurrenyBandName"

                If .ComboIndex = -1 Then Exit Sub
                LngNewID = .ComboData(.ComboIndex)
                LngFoundRow = .FindRow(LngNewID, .FixedRows, .ColIndex("CurrenyBandName"), False, True)

                If LngFoundRow <> -1 Then
                    If LngFoundRow <> Row Then
                        Msg = "·«Ì„ﬂ‰  ﬂ—«— «·›∆… ›Ï ‰›” ⁄„·Ì… «·Ã—œ"
                        Msg = Msg & Chr(13) & "«·›∆… «·„Œ «—… ›Ï «·”ÿ— —ﬁ„ " & LngFoundRow
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        .TextMatrix(Row, Col) = ""
                    Else
                        PutItemInGrid LngNewID, Row
                    End If

                Else
                    PutItemInGrid LngNewID, Row
                End If

            Case "CurrencyBandCount"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
                .Cell(flexcpBackColor, Row, Col, Row, Col) = 0
                .CellBorder 0, 0, 0, 0, 0, 0, 0
        End Select

        CalculateGrid
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    With Me.FG

        Select Case .ColKey(Col)

            Case "CurrenyBandName"

                If m_GridDisplayMode = ShowAllBands Then
                    Cancel = True
                ElseIf m_GridDisplayMode = ShowNonBands Then
                    Cancel = False
                End If

            Case "CurrenyBandValue"
                Cancel = True

            Case "CurrencyBandCount"

                If .TextMatrix(Row, .ColIndex("CurrenyBandName")) = "" Then
                    Cancel = True
                End If

            Case "CurrencyValue"
                Cancel = True

            Case Else
        End Select

    End With

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.FG

        Select Case .ColKey(Col)

            Case "CurrencyBandCount"
                .Cell(flexcpBackColor, Row, Col, Row, Col) = vbYellow
            
                .Select Row, Col, Row, Col
                .CellBorder vbRed, 2, 2, 1, 1, 1, 1
        End Select

    End With

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
CmdAttach.Caption = "Attachments"

    'Me.XPTab301.TabCaption(0) = "Items"
    
    'Me.XPTab301.TabCaption(1) = "Notes"

    Me.Caption = "Box Stock"
    EleHeader.Caption = Me.Caption

    lbl(0).Caption = "ID"
    lbl(7).Caption = "Date"
    lbl(3).Caption = "Box"

    lbl(8).Caption = "Balance "
    Fra.Caption = "Info"
    Chk.Caption = "View all Category"

    lbl(5).Caption = "Remarks"

    lbl(6).Caption = "BY"
 
    lbl(1).Caption = " Total:"
 
    lbl(2).Caption = "Curr. Rec."
    lbl(4).Caption = "Rec. Count:"

    With Me.FG
        .TextMatrix(0, .ColIndex("serial")) = "Index"
        .TextMatrix(0, .ColIndex("CurrenyBandName")) = " Band Name"
        .TextMatrix(0, .ColIndex("CurrencyBandCount")) = "Band Count"
        .TextMatrix(0, .ColIndex("CurrencyValue")) = "Currency Value"
    End With

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
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & "ﬂÊœ  «·⁄„·Ì…  " & XPTxtID.text & Chr(13) & "   «· «—ÌŒ  " & XPDtbTrans & Chr(13) & "   «·Œ“Ì‰…  " & DcboBox & Chr(13) & "   «·«Ã„«·Ì  " & LblTotal
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & "Opr Code " & XPTxtID.text & Chr(13) & "   Date  " & XPDtbTrans & Chr(13) & "   Box  " & DcboBox & Chr(13) & "   Total  " & LblTotal
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, , , , XPTxtID
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", , , , XPTxtID
    End If
    
End Function

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim strsql As String
    Dim GtdBack As New ClsBackGroundPic

    Dim StrComboList As String

    On Error GoTo ErrTrap
    ScreenNameArabic = " Ã—œ «·Œ“‰… "
    ScreenNameEnglish = "Box inventory "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"
 
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
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    SetDtpickerDate Me.XPDtbTrans
    Set Me.FG.WallPaper = GtdBack.MoneyWallpaper
    SetupGrid ShowNonBands
    Set rs = New ADODB.Recordset
    strsql = "TblBoxStock"
    'MsgBox "MSG6"
    rs.Open strsql, Cn, adOpenStatic, adLockOptimistic, adCmdTable
    'MsgBox "MSG7"
    XPDtbTrans.value = Date
    'MsgBox "MSG8"
    XPBtnMove_Click 2
    'MsgBox "MSG9"
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    Msg = "ÕœÀ Œÿ« „« «À‰«¡ ⁄„· «·»—‰«„Ã"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub AddTip()
    'Dim Wrap As String
    'On Error GoTo ErrTrap
    'Wrap = Chr(13) + Chr(10)
    'Set TTP = New clstooltip
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(0), _
    '    "ÃœÌœ ..." & Wrap & _
    '    "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(1), _
    '    " ⁄œÌ· ..." & Wrap & _
    '    "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(2), _
    '    "Õ›Ÿ ..." & Wrap & _
    '    "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & _
    '     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(3), _
    '    " —«Ã⁄ ..." & Wrap & _
    '    "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
    '     "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    ' With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(4), _
    '    "Õ–› ..." & Wrap & _
    '    "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(6), _
    '    "Œ—ÊÃ ..." & Wrap & _
    '    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnMove(1), _
    '    "«·√Ê· ..." & Wrap & _
    '    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnMove(0), _
    '    "«·”«»ﬁ ..." & Wrap & _
    '    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnMove(3), _
    '    "«· «·Ì ..." & Wrap & _
    '    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnMove(2), _
    '    "«·√ŒÌ— ..." & Wrap & _
    '    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "”Õ» „‰ «·Œ“‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl CmdHelp, _
    '    "„”«⁄œ… ..." & Wrap & _
    '    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
    '    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
    '    "≈÷€ÿ Â‰«" & Wrap, True
    'End With
    'Exit Sub
    'ErrTrap:
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

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub TxtModFlg_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "Ã—œ «·Œ“‰…"
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
        
            TxtRemarks.locked = True
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

            Me.XPBtnAdd.Enabled = False
            Me.XPBtnRemove.Enabled = False
            Me.FG.Editable = flexEDNone
            Me.Chk.Enabled = True

        Case "N"
            '        Me.Caption = "Ã—œ «·Œ“‰… ( ÃœÌœ )"
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
            TxtRemarks.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date
            Me.XPBtnAdd.Enabled = True
            Me.XPBtnRemove.Enabled = True
        
            Me.FG.Editable = flexEDKbdMouse
            Me.Chk.Enabled = True

        Case "E"
            '        Me.Caption = "Ã—œ «·Œ“‰…(  ⁄œÌ· )"
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
        
            TxtRemarks.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            Me.XPBtnAdd.Enabled = True
            Me.XPBtnRemove.Enabled = True
            Me.FG.Editable = flexEDKbdMouse
            Me.Chk.Enabled = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnAdd_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("CurrenyBandName")) <> "" Then
        FG.Rows = FG.Rows + 1
        'NewGrid.GridDefaultValue Fg.Rows - 1
        FG.Row = FG.Rows - 1
        FG.Col = FG.ColIndex("CurrenyBandName")
        FG.ShowCell FG.Rows - 1, FG.ColIndex("CurrenyBandName")
        FG.SetFocus
    End If

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

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.Rows > 1 Then
        If FG.Rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
            '        NewGrid.Calculate 1, True
            CalculateGrid
        Else

            If FG.Rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Row)
                End If
            End If

            CalculateGrid
            '        NewGrid.Calculate 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub PutItemInGrid(Lngid As Long, _
                         LngRow As Long)
    Dim strsql As String
    Dim rs As ADODB.Recordset

    strsql = "SELECT TblCurrencyBandNames.CurrencyBandID, "
    strsql = strsql + " TblCurrencyBandNames.CurrencyBandName," & "TblCurrencyBandNames.CurrenyBandValue "
    strsql = strsql + " FROM TblCurrencyBandNames "
    strsql = strsql + " Where TblCurrencyBandNames.CurrencyBandID=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FG
            .TextMatrix(LngRow, .ColIndex("CurrenyBandValue")) = IIf(IsNull(rs("CurrenyBandValue").value), 0, rs("CurrenyBandValue").value)
        End With

    End If

    rs.Close
    Set rs = Nothing
    CalculateGrid

End Sub

Private Sub CalculateGrid()
    Dim i As Integer

    With Me.FG

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("CurrencyValue")) = val(.TextMatrix(i, .ColIndex("CurrenyBandValue"))) * val(.TextMatrix(i, .ColIndex("CurrencyBandCount")))
        Next i

        LblTotal.Caption = ""
        LblTotal.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CurrencyValue"), .Rows - 1, .ColIndex("CurrencyValue"))

    End With

End Sub

Private Function CheckBoxStockDate(LngBoxID As Long, _
                                   D_StockDate As Date) As Boolean
    Dim strsql As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    strsql = "SELECT BoxStockID, BoxStockDate, BoxID, Remarks, UserID"
    strsql = strsql + " From TblBoxStock"
    strsql = strsql + " Where BoxID=" & LngBoxID & ""
    strsql = strsql + " AND BoxStockDate=" & SQLDate(D_StockDate, True) & ""

    If Me.TxtModFlg.text = "E" Then
        strsql = strsql + " AND BoxStockID <> " & val(Me.XPTxtID.text) & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Msg = "⁄›Ê« .. Â‰«ﬂ Õ—ﬂ… Ã—œ ·Â–Â «·Œ“‰… "
        Msg = Msg & Chr(13) & "„”Ã· „”»ﬁ« ›Ï ‰›” «· «—ÌŒ «·–Ï  —Ìœ «· ”ÃÌ· ›ÌÂ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckBoxStockDate = False
    Else
        CheckBoxStockDate = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub XPDtbTrans_Change()
    GetBoxData
End Sub

Private Sub SetupGrid(IntMode As GridDisplayMode)
    Dim strsql As String
    Dim RsTemp As ADODB.Recordset
    Dim i As Integer

    If m_GridDisplayMode = ShowNonBands Then

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows + 1
            strsql = "SELECT TblCurrencyBandNames.CurrencyBandID, "
            strsql = strsql + " TblCurrencyBandNames.CurrencyBandName," & "TblCurrencyBandNames.CurrenyBandValue "
            strsql = strsql + " FROM TblCurrencyBandNames "
            Set RsTemp = New ADODB.Recordset
    
            RsTemp.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = .BuildComboList(RsTemp, "CurrencyBandName", "CurrencyBandID")
            .ColComboList(.ColIndex("CurrenyBandName")) = "|" & StrComboList
            .AutoSize 0, .Cols - 1, False
        End With

        Me.XPBtnAdd.Enabled = True
        Me.XPBtnRemove.Enabled = True
    ElseIf m_GridDisplayMode = ShowAllBands Then

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows + 1
            strsql = "SELECT TblCurrencyBandNames.CurrencyBandID, "
            strsql = strsql + " TblCurrencyBandNames.CurrencyBandName," & "TblCurrencyBandNames.CurrenyBandValue "
            strsql = strsql + " FROM TblCurrencyBandNames "
            strsql = strsql + "Order By CurrencyBandID DESC"
            Set RsTemp = New ADODB.Recordset
    
            RsTemp.Open strsql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            .Rows = .FixedRows + RsTemp.RecordCount
        
            StrComboList = .BuildComboList(RsTemp, "CurrencyBandName", "CurrencyBandID")
            .ColComboList(.ColIndex("CurrenyBandName")) = "|" & StrComboList
            RsTemp.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("CurrenyBandName")) = RsTemp("CurrencyBandID").value
                .TextMatrix(i, .ColIndex("CurrenyBandValue")) = RsTemp("CurrenyBandValue").value
                RsTemp.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

        Me.XPBtnAdd.Enabled = False
        Me.XPBtnRemove.Enabled = False
    End If

End Sub

