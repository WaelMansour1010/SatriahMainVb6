VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmItemsDetails2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Œ’Ì’ ”‰œ«  ’—› «·«‰ «Ã "
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   Icon            =   "FrmItemsDetails2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtItemCodeB 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9360
      TabIndex        =   35
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check7 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌœ/«·€«¡ «·ﬂ·"
      Height          =   255
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   855
      Left            =   4080
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Text            =   "FrmItemsDetails2.frx":000C
      Top             =   7320
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      Caption         =   "„·«ÕŸ…"
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   -2040
      Width           =   4575
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«»œ „‰  ›⁄Ì· «·«·Ê«‰ Ê«·„ﬁ«”«  Ê«·›—“ „‰ „œÌ— «·‰Ÿ«„"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtArrivalDate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtUnitName 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtFoxyNo 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14400
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox TxtUnitID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtItemID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14640
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtItemName 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox TxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5220
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   9555
      _cx             =   16854
      _cy             =   9208
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
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmItemsDetails2.frx":0012
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
   Begin ALLButtonS.ALLButton CmdRemove 
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      Tag             =   "Delete Row"
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmItemsDetails2.frx":03FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "Delete Row"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«÷«›… ”ÿ—"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmItemsDetails2.frx":041A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   540
      Index           =   1
      Left            =   -240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8400
      Width           =   13575
      _cx             =   23945
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
         Index           =   1
         Left            =   10605
         TabIndex        =   15
         Top             =   90
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
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
         Height          =   375
         Index           =   3
         Left            =   7635
         TabIndex        =   16
         Top             =   90
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
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
         Height          =   375
         Index           =   4
         Left            =   6000
         TabIndex        =   17
         Top             =   90
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
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
         Height          =   375
         Index           =   5
         Left            =   4500
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   30
         TabIndex        =   19
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   2985
         TabIndex        =   20
         Top             =   90
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         ButtonStyle     =   1
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   1500
         TabIndex        =   21
         Top             =   90
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
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
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   0
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   9555
      _cx             =   16854
      _cy             =   1349
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
      BackColor       =   16777152
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmItemsDetails2.frx":0436
      Caption         =   " Œ’Ì’ ”‰œ«  ’—› «·«‰ «Ã "
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
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   6720
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   6720
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·»«—ﬂÊœ"
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
      ButtonImage     =   "FrmItemsDetails2.frx":1110
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton SearchCashCustomer 
      Height          =   315
      Left            =   9000
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
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
      ButtonImage     =   "FrmItemsDetails2.frx":14AA
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»«—ﬂÊœ"
      Height          =   255
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "5"
      Height          =   255
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   -1800
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "ÊÕœÂ"
      Height          =   255
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·’‰›"
      Height          =   375
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   -1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·«„—"
      Height          =   375
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "FrmItemsDetails2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim workwithsize As Boolean
Dim workwithcolor As Boolean
Dim itemsWorkWithDates As Boolean
 
Dim itemsWorkWithClass As Boolean
Dim itemcodePart1 As Integer
Dim itemcodePart2 As Integer
Dim itemcodePart3 As Integer
Dim itemcodeSeperator1 As String
Dim itemcodeSeperator2 As String
Dim itemcodePart1NoOFDigit As Integer
Dim itemcodePart2NoOFDigit As Integer
Dim itemcodePart3NoOFDigit As Integer
Dim codeNoofDigit  As Integer
Dim SizeNoofDigit As Integer
Dim ColorNoofDigit As Integer
Dim itemcodePlace As Integer
Dim ColorPlace  As Integer
Dim SizePlace  As Integer
Public FG As VSFlex8UCtl.vsFlexGrid

Public LngRow As Long

Public LngCol As Long

Public AllDate As String
Public AllIDS As String
Public Allline As String

Private Sub Check7_Click()
    
    Dim i As Integer
   Grid.Enabled = True
    If Check7.value = vbChecked Then

        With Me.Grid
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = True
            Next i

        End With

    Else

        With Me.Grid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = False
            Next i

        End With

    End If


End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 1
         '   Me.Grid.Rows = Me.Grid.Rows + 1

        Case 2
         '   SaveData


       If Me.FG.ColIndex("ItemsDetailsNewidea") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("ItemsDetailsNewidea")) = AllIDS
        End If

    If Me.FG.ColIndex("Count") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("Count")) = val(txtqty.text)
  
        End If
  Unload Me

        Case 6
            Unload Me
    End Select

End Sub

 

Public Sub Retrive(Optional FoxyNo As String)
     
 

End Sub

Private Sub CmdRemove_Click()
    RemoveGridRow
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub Command1_Click()
    Retrive
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
   
AllIDS = ""
txtqty.text = 0
 Dim total As Double

total = 0
LBLInsWages = 0
lBLnOoFsTNES = 0
    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("count")) <> "" Then
            Allline = ""
                IntCounter = IntCounter + 1
'               .TextMatrix(i, .ColIndex("Ser")) = IntCounter
        '        AllDate = AllDate & .TextMatrix(i, .ColIndex("MaDate")) & ","
        




       Allline = .TextMatrix(i, .ColIndex("ItemFullcode")) & "@@" & .TextMatrix(i, .ColIndex("ParrtNoCode")) & "@@"
        
     Allline = Allline & .TextMatrix(i, .ColIndex("count")) & "@@"
        Allline = Allline & (.TextMatrix(i, .ColIndex("unitid"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("ColorID"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("sizeid"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("ClassId"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("ProductionDate"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("ExpiryDate"))) & "@@"
         AllIDS = AllIDS & Allline & "&&"
         
        ' .TextMatrix(i, .ColIndex("Total")) = (val(.TextMatrix(i, .ColIndex("price"))) * val(.TextMatrix(i, .ColIndex("weight")))) + (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("InstallPrice"))))
        ' lBLnOoFsTNES = lBLnOoFsTNES + val(.TextMatrix(i, .ColIndex("Count")))
        ' LBLInsWages = LBLInsWages + (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("InstallPrice"))))
        ' total = total + val(.TextMatrix(i, .ColIndex("Total")))
         txtqty = val(txtqty) + val(.TextMatrix(i, .ColIndex("count")))
            End If

        Next i
' lblTotals.Caption = total + val(txtWages)
 
    End With
    
Text1.text = AllIDS
End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim strFilterText1 As String
      Dim Unitname As String
      Dim sizename As String
      Dim colorname As String
      Dim classname As String
    Dim ttypename As String
     Dim typename As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim inty As Integer
    Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
 
    
     Dim astrSplitItems1() As String
     If AllIDS = "" Then
     Exit Sub
     End If
     
    strFilterText = "&&"
         strFilterText1 = "@@"
    astrSplitItems = Split(Me.AllIDS, strFilterText)
    Grid.Rows = UBound(astrSplitItems) + 1
 
    For intX = 0 To UBound(astrSplitItems) - 1
    
    astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
           

 

                Grid.TextMatrix(intX + 1, Grid.ColIndex("ItemFullcode")) = astrSplitItems1(0)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("ParrtNoCode")) = astrSplitItems1(1)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("count")) = astrSplitItems1(2)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("ActCount")) = Grid.TextMatrix(intX + 1, Grid.ColIndex("count"))
'                    GRID.TextMatrix(intX + 1, GRID.ColIndex("uniteid")) = astrSplitItems1(2)
        '    GRID.Cell(flexcpData, intX + 1, GRID.ColIndex("uniteid")) = GRID.TextMatrix(intX + 1, GRID.ColIndex("uniteid"))
        
        Grid.TextMatrix(intX + 1, Grid.ColIndex("unitid")) = astrSplitItems1(3)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("ColorID")) = astrSplitItems1(4)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("sizeid")) = astrSplitItems1(5)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("ClassId")) = astrSplitItems1(6)
                Grid.TextMatrix(intX + 1, Grid.ColIndex("ProductionDate")) = astrSplitItems1(7)
   Grid.TextMatrix(intX + 1, Grid.ColIndex("ExpiryDate")) = astrSplitItems1(7)
  
     
  ' ttypename As String, Optional ByRef typename
   
  GetiItemsNewDetails val(Grid.TextMatrix(intX + 1, Grid.ColIndex("unitid"))), val(Grid.TextMatrix(intX + 1, Grid.ColIndex("sizeid"))) _
  , val(Grid.TextMatrix(intX + 1, Grid.ColIndex("ColorID"))), val(Grid.TextMatrix(intX + 1, Grid.ColIndex("ClassId"))), Unitname, sizename, colorname, classname
 
   Grid.TextMatrix(intX + 1, Grid.ColIndex("UnitName")) = Unitname
  Grid.TextMatrix(intX + 1, Grid.ColIndex("colorname")) = colorname
  Grid.TextMatrix(intX + 1, Grid.ColIndex("sizename")) = sizename
  Grid.TextMatrix(intX + 1, Grid.ColIndex("classname")) = classname
  
  'GRID.TextMatrix(intX + 1, GRID.ColIndex("TType")) = ttypename
  
 
    Next
     
     
     Grid.Rows = Grid.Rows + 1
     
      ReLineGrid
     
     
ErrTrap:
End Sub

Private Sub ReLineGridx()
    Dim i As Integer
    Dim IntCounter As Integer

    With Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemFullcode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    IntCounter = 0
 
End Sub

Function CreateItemCodePerLine(Optional itemcode As String, Optional itemsize As String, Optional itemcolor As String) As String
 
    With Me.Grid
 .TextMatrix(.Row, .ColIndex("UnitID")) = 1
     '   unitid = .TextMatrix(.Row, .ColIndex("UnitID"))
        itemcolor = .TextMatrix(.Row, .ColIndex("ColorName"))
        itemsize = .TextMatrix(.Row, .ColIndex("SizeName"))
        itemcode = lbl(4).Caption
SizeNoofDigit = 5
ColorNoofDigit = 5
        If Len(itemsize) > SizeNoofDigit Then
            MsgBox "errr"
            Exit Function
        End If

        If Len(itemcolor) > ColorNoofDigit Then
            MsgBox "errr"
            Exit Function
        End If

        If Len(itemsize) < SizeNoofDigit Then
       '     itemsize = Format(itemsize, String(SizeNoofDigit, "0"))
        End If

        If Len(itemcolor) < ColorNoofDigit Then
       '     itemcolor = Format(itemcolor, String(ColorNoofDigit, "0"))
        End If

        If workwithcolor = False Then
            itemcolor = ""
        End If

        If workwithsize = False Then
            itemsize = ""
        End If
 
        Dim fullcode As String
itemcodePlace = 1
SizePlace = 3

        If itemcodePlace = 1 Then

            If SizePlace = 2 Then
                fullcode = itemcode & itemcodeSeperator1 & itemsize & itemcodeSeperator2 & itemcolor
            ElseIf SizePlace = 3 Then
                fullcode = itemcode & itemcodeSeperator1 & itemcolor & itemcodeSeperator2 & itemsize
            End If

        ElseIf itemcodePlace = 2 Then

            If SizePlace = 1 Then
                fullcode = itemsize & itemcodeSeperator1 & itemcode & itemcodeSeperator2 & itemcolor
            ElseIf SizePlace = 3 Then
                fullcode = itemcolor & itemcodeSeperator1 & itemcode & itemcodeSeperator2 & itemsize
            End If

        ElseIf itemcodePlace = 3 Then

            If SizePlace = 1 Then
                fullcode = itemsize & itemcodeSeperator1 & itemcolor & itemcodeSeperator2 & itemcode
            ElseIf SizePlace = 2 Then
                fullcode = itemcolor & itemcodeSeperator1 & itemsize & itemcodeSeperator2 & itemcode
            End If
        End If

        .TextMatrix(.Row, .ColIndex("ItemFullcode")) = fullcode
       ' .TextMatrix(.Row, .ColIndex("ParrtNoCode")) = fullcode
        
        .TextMatrix(.Row, .ColIndex("ArrivalDate")) = TxtArrivalDate.text

        CreateItemCodePerLine = fullcode
    End With

End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim unitid As String
    Dim sizeid As String
    Dim ColorID As String
    Dim itemcode As String
    Dim fullcode As String
    Dim code As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
            Case "ColorName"

                code = .ComboData
                '  LngRow = .FindRow(Code, .FixedRows, .ColIndex("ColorID"), False, True)
                .TextMatrix(Row, .ColIndex("ColorID")) = code
                .TextMatrix(Row, .ColIndex("ColorName")) = .ComboItem
             
            Case "SizeName"
                code = .ComboData
                '      LngRow = .FindRow( Code , .FixedRows, .ColIndex("SizeId"), False, True)
                .TextMatrix(Row, .ColIndex("SizeId")) = code
                .TextMatrix(Row, .ColIndex("SizeName")) = .ComboItem
             
            Case "ClassName"
                code = .ComboData
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("ClassId"), False, True)
                .TextMatrix(Row, .ColIndex("ClassId")) = code
                .TextMatrix(Row, .ColIndex("ClassName")) = .ComboItem
       
        End Select

       fullcode = CreateItemCodePerLine(itemcode, sizeid, ColorID)
   '     fullcode = CreateItemCodePerLine(ItemCode, ColorID, sizeid)
        
    End With
ReLineGrid
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim MyStrList As String
    Dim LngItemID As Long

    With Me.Grid

        '    If Me.GridTrans = InvoiceTransaction Then
        Select Case .ColKey(Col)
           
            Case "ClassName"
                StrSQL = "SELECT  SizeId,SizeName "
                StrSQL = StrSQL + " FROM TblItemsclasses  "
                StrSQL = StrSQL + " Order BY SizeId "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    MyStrList = .BuildComboList(rs, "SizeName", "SizeId")
                    '                    Grid.ColComboList = MyStrList
                    Grid.ColComboList(.ColIndex("ClassName")) = "|" & MyStrList
                Else
                    Cancel = True
                End If
                
            Case "UnitName"
                '           LngItemID = Val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))
                LngItemID = val(TxtItemID.text)

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
            Case "ColorName"
            
                StrSQL = "SELECT  ColorID,ColorName "
                StrSQL = StrSQL + " FROM TblItemsColors  "
                StrSQL = StrSQL + " Order BY ColorName "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    MyStrList = .BuildComboList(rs, "ColorName", "ColorID")
                    '                    Grid.ColComboList = MyStrList
                    Grid.ColComboList(.ColIndex("ColorName")) = "|" & MyStrList
                Else
                    Cancel = True
                End If

            Case "SizeName"
            
                StrSQL = "SELECT  SizeId,SizeName "
                StrSQL = StrSQL + " FROM TblItemsSizes  "
                StrSQL = StrSQL + " Order BY SizeId "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    MyStrList = .BuildComboList(rs, "SizeName", "SizeId")
                    '                    Grid.ColComboList = MyStrList
                    Grid.ColComboList(.ColIndex("SizeName")) = "|" & MyStrList
                Else
                    Cancel = True
                End If
            
        End Select
        
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
            
            
                      
    
              
        
    
            
            
        End If

        ReLineGrid
    
        '    End If
    End With

End Sub
 
Private Sub ISButton1_Click()
  Dim str As String

    Dim RowNum As Integer
    Dim ItemCount As Integer
    str = "Delete  TblPrintBarCode"
    Cn.Execute str
DoEvents
Dim LngItemID As Long
Dim LngUnitID As Long
    'cBarcode.AddItem
    ' cBarcode.ClearItems
  

    LngItemID = val(TxtItemID.text)
    LngUnitID = val(TxtUnitID.text)
    For RowNum = 1 To Grid.Rows - 1

        If Grid.Cell(flexcpChecked, RowNum, Grid.ColIndex("Print")) = flexChecked Then
            If Not IsNull(Grid.TextMatrix(RowNum, Grid.ColIndex("ActCount"))) Then
           
      addtotable val(Grid.TextMatrix(RowNum, Grid.ColIndex("ActCount"))), Grid.TextMatrix(RowNum, Grid.ColIndex("ParrtNoCode")), GetItemPrice(LngItemID, 1, LngUnitID), Grid.TextMatrix(RowNum, Grid.ColIndex("ItemFullcode")), lbl(5).Caption, lbl(5).Caption, _
      Grid.TextMatrix(RowNum, Grid.ColIndex("ColorName")), Grid.TextMatrix(RowNum, Grid.ColIndex("SizeName")), Grid.TextMatrix(RowNum, Grid.ColIndex("ClassName"))
          
            End If
        End If

    Next RowNum

    printCodes WindowTarget


End Sub
Public Sub printCodes(m_PrintTarget As PrintTarget)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo

    If Dir(App.path & "\Reports\Inventory\" & "BarCode1.rpt") = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

        MySQL = " "

 MySQL = MySQL & "  SELECT     dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color,"
 MySQL = MySQL & "       dbo.TblPrintBarCode.[size] , dbo.TblPrintBarCode.Class, dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ItemID, dbo.TblItems.ItemNamee"
 MySQL = MySQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
 MySQL = MySQL & "     dbo.ItemsDetails ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId"
 MySQL = MySQL & " RIGHT OUTER JOIN"
  MySQL = MySQL & "  dbo.TblPrintBarCode ON dbo.ItemsDetails.ParrtNoCode = dbo.TblPrintBarCode.Code"

  MySQL = "SELECT     Code, PartNo, Cost, Name, Color, [size], class, Code AS itemcode, Name AS itemname FROM         dbo.TblPrintBarCode"
  
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
  
    Else
       
        Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & "BarCode1.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        
    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & "BarCode1.rpt"

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub






Function addtotable(NoOfRow As Integer, code As String, cost As Double, Optional PartNo As String = "", Optional name As String = "" _
, Optional NameE As String, Optional Color As String, Optional size As String, Optional Class As String, Optional itemcode As String)
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim i As Integer

    str = "select * from   TblPrintBarCode where 1=-1"
    
   rs.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 1 To NoOfRow
        rs.AddNew
   
        rs("PartNo").value = PartNo
        rs("code").value = code
        rs("cost").value = val(cost)
        rs("Name").value = name
'        rs("NameE").value = NameE
        rs("Color").value = Color
        rs("size").value = size
        rs("class").value = Class
        rs.update
    Next i
'
End Function



 

Private Sub SearchCashCustomer_Click()
 
     Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 3
        FrmItemSearch2.show
 

End Sub

 Private Sub TxtItemCodeB_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 '   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
Dim colorname As String
Dim sizename As String
  Dim classname As String
  Dim Unitname As String

  Dim LngRow As Integer
'    StrSQL = " SELECT     ItemDetailedCode, ParrtNoCode, ProductionDate, ExpireDate, ColorID, UnitID, SizeID, ClassId, ItemId"
'StrSQL = StrSQL & " from dbo.ItemsDetails"
'StrSQL = StrSQL & " GROUP BY ItemDetailedCode, ParrtNoCode, ColorID, UnitID, SizeID, ClassId, ProductionDate, ExpireDate, ItemId"
'StrSQL = StrSQL & "  HAVING        (ParrtNoCode = '" & TxtItemCodeB & "')"
     StrSQL = " SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ColorID, "
 StrSQL = StrSQL & "       dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.ItemId, dbo.TblItemsclasses.SizeName AS classname,"
  StrSQL = StrSQL & "   dbo.TblItemsSizes.sizename , dbo.TblItemsColors.colorname, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"
StrSQL = StrSQL & "  FROM         dbo.ItemsDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "     dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   StrSQL = StrSQL & "    dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
  StrSQL = StrSQL & "      dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "     dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "   GROUP BY dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID,"
StrSQL = StrSQL & "    dbo.ItemsDetails.ClassId, dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ItemId, dbo.TblItemsclasses.SizeName,"
StrSQL = StrSQL & "   dbo.TblItemsSizes.sizename , dbo.TblItemsColors.colorname, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"
'StrSQL = StrSQL & "  HAVING      (dbo.ItemsDetails.ParrtNoCode = '883884777264')"

 StrSQL = StrSQL & "  HAVING        (ParrtNoCode = '" & TxtItemCodeB & "')"
     
       
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
  '      LngItemID = val(DCboItemCode.BoundText)
        LngItemID = IIf(IsNull(RsTemp("itemid").value), 0, RsTemp("ItemID").value)
  LngUnitID = IIf(IsNull(RsTemp("UnitID").value), 0, RsTemp("UnitID").value)
  ColorID = IIf(IsNull(RsTemp("ColorID").value), 0, RsTemp("ColorID").value)
  sizeid = IIf(IsNull(RsTemp("SizeID").value), 0, RsTemp("SizeID").value)
  ClassId = IIf(IsNull(RsTemp("ClassId").value), 0, RsTemp("ClassId").value)
    LngUnitID = IIf(IsNull(RsTemp("UnitID").value), 0, RsTemp("UnitID").value)


  colorname = IIf(IsNull(RsTemp("colorname").value), 0, RsTemp("colorname").value)
  sizename = IIf(IsNull(RsTemp("sizename").value), 0, RsTemp("sizename").value)
  classname = IIf(IsNull(RsTemp("classname").value), 0, RsTemp("classname").value)
   If SystemOptions.UserInterface = ArabicInterface Then
      Unitname = IIf(IsNull(RsTemp("Unitname").value), 0, RsTemp("Unitname").value)
    Else
    Unitname = IIf(IsNull(RsTemp("UnitNamee").value), 0, RsTemp("UnitNamee").value)
    End If
    
  

            
            
 
    ParrtNoCode = IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        ItemDetailedCode = IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID <> 0 Then


        
    With Grid

    
      
 LngRow = .Rows - 1
            .TextMatrix(LngRow, .ColIndex("ColorID")) = ColorID
            .TextMatrix(LngRow, .ColIndex("sizeid")) = sizeid
            .TextMatrix(LngRow, .ColIndex("Classid")) = ClassId
       .TextMatrix(LngRow, .ColIndex("unitid")) = LngUnitID
       
                 .TextMatrix(LngRow, .ColIndex("ColorName")) = colorname
            .TextMatrix(LngRow, .ColIndex("SizeName")) = sizename
            .TextMatrix(LngRow, .ColIndex("ClassName")) = classname
                .TextMatrix(LngRow, .ColIndex("Unitname")) = Unitname
       .TextMatrix(LngRow, .ColIndex("Count")) = 1
            .TextMatrix(LngRow, .ColIndex("ParrtNoCode")) = ParrtNoCode
         .TextMatrix(LngRow, .ColIndex("ItemFullcode")) = ItemDetailedCode
          ReLineGrid
  .Rows = .Rows + 1

RsTemp.Close
Set rs = Nothing
     End With
           
           
 
         'DCboItemCode_KeyDown vbKeyReturn, 0
          Me.TxtItemCodeB.text = ""
      '    Unload FrmItemSearch2
      Me.TxtItemCodeB.SetFocus
         
    Else
        Exit Sub
    End If
    
    
End If

End If
End Sub



