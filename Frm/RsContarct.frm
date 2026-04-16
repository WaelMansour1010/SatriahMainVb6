VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RSContract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄ﬁÊœ"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16605
   Icon            =   "RsContarct.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   16605
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Caption         =   "«· √ÀÌÀ"
      Height          =   3855
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   -4320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.ComboBox DcbFurnishing 
         Height          =   315
         ItemData        =   "RsContarct.frx":57E2
         Left            =   2880
         List            =   "RsContarct.frx":57EC
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   240
         Width           =   3255
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2565
         Left            =   0
         TabIndex        =   59
         Top             =   720
         Width           =   6885
         _cx             =   12144
         _cy             =   4524
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"RsContarct.frx":5800
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label3"
         Height          =   15
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«· √ÀÌÀ"
         Height          =   285
         Index           =   29
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "⁄—÷ «·œ›⁄« "
      Height          =   375
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox DcbContType 
      Height          =   315
      ItemData        =   "RsContarct.frx":5910
      Left            =   -720
      List            =   "RsContarct.frx":591A
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Õ”«» «·œ›⁄« "
      Height          =   375
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtContNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   -60
      Width           =   945
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1125
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   10560
      Visible         =   0   'False
      Width           =   10605
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "RsContarct.frx":5932
         Left            =   2280
         List            =   "RsContarct.frx":5942
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   870
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   3405
      Left            =   21840
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   6405
      _cx             =   11298
      _cy             =   6006
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"RsContarct.frx":595B
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
   Begin ImpulseButton.ISButton BtnPrint 
      Height          =   525
      Left            =   9000
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   10200
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      ButtonStyle     =   1
      ButtonPositionImage=   2
      Caption         =   "ÿ»«⁄Â «·⁄ﬁœ"
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
      ButtonImage     =   "RsContarct.frx":5A08
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   450
      Left            =   6360
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   7080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«·œ›⁄« "
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "RsContarct.frx":5DA2
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic ELe 
      Height          =   9795
      Index           =   10
      Left            =   0
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   0
      Width           =   16605
      _cx             =   29289
      _cy             =   17277
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
      Align           =   5
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
      Begin VB.CommandButton Command15 
         Caption         =   " ÊÀÌﬁ «·ﬂ·"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   288
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox ChkAccredit 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «· ÊÀÌﬁ"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   283
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtDiscountValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11280
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   282
         Top             =   3480
         Width           =   705
      End
      Begin VB.TextBox txtDiscountPercent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13740
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   279
         Top             =   3540
         Width           =   705
      End
      Begin VB.CommandButton Command14 
         Caption         =   " ÕœÌÀ"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   277
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtRemark2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   272
         Top             =   8280
         Width           =   2370
      End
      Begin VB.CommandButton Command13 
         Caption         =   "≈‰‘«¡ «—ﬁ«„ ”‰œ«  «·œ›⁄« "
         Height          =   255
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   242
         Top             =   9480
         Width           =   2040
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   239
         Top             =   600
         Width           =   1815
         Begin VB.OptionButton RdRTypeDate 
            Alignment       =   1  'Right Justify
            Caption         =   "ÂÃ—Ì"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton RdRTypeDate 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ì·«œÌ"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   240
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CommandButton CMDSENDSMS 
         Caption         =   "«—”«· —”«·Â"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   238
         Top             =   9480
         Width           =   975
      End
      Begin VB.TextBox TxtOldID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   232
         Top             =   0
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   215
         Top             =   600
         Width           =   1935
         Begin VB.OptionButton ComResid 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ«÷⁄"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton ComResid 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— Œ«÷⁄"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.OptionButton FrmContractOldData 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -960
         RightToLeft     =   -1  'True
         TabIndex        =   212
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtOthersRules 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4560
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   7560
         Width           =   2370
      End
      Begin VB.TextBox TxtNotID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "ÃœÌœ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   15000
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "«›  «ÕÌ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   13920
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14280
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox ChKOutContract 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄ﬁœ Œ«—ÃÌ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   12720
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox ChkRenew 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «·‰ÃœÌœ"
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox ChKEndContract 
         Alignment       =   1  'Right Justify
         Caption         =   " „ «· ’›Ì…"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox ChkEmployeecontract 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊŸ› ‘—ﬂ…"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtEmpCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   1080
         Width           =   1065
      End
      Begin VB.CheckBox ChKLegalIssue 
         Alignment       =   1  'Right Justify
         Caption         =   "‘∆Ê‰ ﬁ«‰Ê‰Ì…"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtNotVal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox TxtNotSreail1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   960
         Width           =   1515
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   0
         Width           =   16635
         Begin VB.CheckBox chkIsNotCreateEntry 
            Alignment       =   1  'Right Justify
            Caption         =   "·« Ì‰‘√ ﬁÌœ „Õ«”»Ï"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   330
            Width           =   2055
         End
         Begin VB.TextBox TXTNewNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   284
            Top             =   120
            Width           =   2985
         End
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   -1590
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   70
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   -585
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "«·„” Œœ„"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   13
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   45
               Width           =   855
            End
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Text            =   "modflag"
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   510
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtContNoOld 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   15120
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   3120
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":613C
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":64D6
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":6870
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":6C0A
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":6FA4
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":733E
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":76D8
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RsContarct.frx":7C72
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   90
            TabIndex        =   72
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":800C
            ColorButton     =   14871017
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   555
            TabIndex        =   73
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":83A6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1155
            TabIndex        =   74
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":8740
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1620
            TabIndex        =   75
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":8ADA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «·„ÊÕœ"
            ForeColor       =   &H000000FF&
            Height          =   270
            Index           =   79
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   285
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label400 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " „"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   525
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   275
            Top             =   -150
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·⁄ﬁÊœ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   0
            Width           =   3990
         End
         Begin VB.Label lblnew 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ÃœÌœ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   -120
            Width           =   2295
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   7920
            Picture         =   "RsContarct.frx":8E74
            Stretch         =   -1  'True
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   4200
            Picture         =   "RsContarct.frx":AAE2
            Stretch         =   -1  'True
            Top             =   -30
            Width           =   525
         End
      End
      Begin MSComCtl2.DTPicker ContDate 
         Height          =   270
         Left            =   10560
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         Format          =   64225283
         CurrentDate     =   41640
      End
      Begin Dynamic_Byte.NourHijriCal RecorddateH 
         Height          =   255
         Left            =   12120
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   4440
         TabIndex        =   88
         Top             =   600
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   89
         Top             =   1080
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   375
         Left            =   7320
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
         Top             =   960
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   661
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
         ButtonImage     =   "RsContarct.frx":E74A
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
      End
      Begin MSComCtl2.DTPicker allowdate 
         Height          =   315
         Left            =   600
         TabIndex        =   90
         Top             =   1200
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64225281
         CurrentDate     =   41640
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   3375
         Index           =   0
         Left            =   0
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1440
         Width           =   11055
         _cx             =   19500
         _cy             =   5953
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
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   "√Ê· ﬁ”ÿ"
            Height          =   240
            Index           =   4
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   465
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ— ﬁ”ÿ"
            Height          =   240
            Index           =   3
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   465
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            Caption         =   "ÌœÊÌ"
            Height          =   240
            Index           =   2
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   465
            Width           =   1095
         End
         Begin VB.CheckBox chkDivElectric 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁ”Ì„ «·ﬂÂ—»«¡ ⁄·Ï «·œ›⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   120
            Width           =   2295
         End
         Begin VB.CheckBox chkDivWater 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁ”Ì„ «·„Ì«Â ⁄·Ï «·œ›⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   2055
         End
         Begin VB.TextBox TxtPaymentCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1065
         End
         Begin VB.TextBox TxtPeriods 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   1065
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "RsContarct.frx":EB47
            Left            =   7320
            List            =   "RsContarct.frx":EB54
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   465
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker FristPaymentDate 
            Height          =   255
            Left            =   4560
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal FirstInstallDateH 
            Height          =   240
            Left            =   6000
            TabIndex        =   36
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   20
            Left            =   480
            TabIndex        =   43
            Top             =   345
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈÷«›…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":EB67
            DrawFocusRectangle=   0   'False
         End
         Begin C1SizerLibCtl.C1Tab TabMain 
            Height          =   2595
            Left            =   60
            TabIndex        =   243
            Top             =   720
            Width           =   10980
            _cx             =   19368
            _cy             =   4577
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
            BackColor       =   12648447
            ForeColor       =   -2147483630
            FrontTabColor   =   14871017
            BackTabColor    =   12648447
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   16711680
            Caption         =   "«·œ›⁄«  |«·œ›⁄«  ﬁ»· «· ⁄œÌ·| Ê«—ÌŒ «· ⁄œÌ·«  ⁄·Ï «·œ›⁄« "
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
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   2220
               Index           =   12
               Left            =   45
               TabIndex        =   244
               TabStop         =   0   'False
               Top             =   45
               Width           =   10890
               _cx             =   19209
               _cy             =   3916
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
               Begin VB.CommandButton cmdSavePayment 
                  Caption         =   "Õ›Ÿ  ⁄œÌ·«  «·œ›⁄« "
                  Height          =   255
                  Left            =   8820
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   1860
                  Width           =   1995
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   2220
                  Index           =   1
                  Left            =   12630
                  TabIndex        =   245
                  Top             =   570
                  Width           =   10830
                  _cx             =   19103
                  _cy             =   3916
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
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"RsContarct.frx":EF01
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
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   1770
                  Left            =   0
                  TabIndex        =   250
                  Top             =   0
                  Width           =   10860
                  _cx             =   19156
                  _cy             =   3122
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
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   12
                  Cols            =   77
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"RsContarct.frx":EFC1
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
               Begin VB.Label lblOldValue 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   269
                  Top             =   1875
                  Width           =   1290
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ »ﬁÌ ”«»ﬁ"
                  Height          =   300
                  Index           =   75
                  Left            =   5850
                  RightToLeft     =   -1  'True
                  TabIndex        =   268
                  Top             =   1875
                  Width           =   825
               End
               Begin VB.Label LBLRemain 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   264
                  Top             =   1920
                  Width           =   1590
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ »ﬁÌ"
                  Height          =   300
                  Index           =   74
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   263
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.Label LblActulaPyaed 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   262
                  Top             =   1920
                  Width           =   1290
               End
               Begin VB.Label LblTotalQasts 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   6555
                  RightToLeft     =   -1  'True
                  TabIndex        =   249
                  Top             =   1890
                  Width           =   1110
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "œ›⁄«  «·⁄ﬁœ"
                  Height          =   300
                  Index           =   34
                  Left            =   7815
                  RightToLeft     =   -1  'True
                  TabIndex        =   248
                  Top             =   1890
                  Width           =   825
               End
               Begin VB.Label LblNotPayed 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   247
                  Top             =   1650
                  Visible         =   0   'False
                  Width           =   1590
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " „”œœ"
                  Height          =   300
                  Index           =   36
                  Left            =   3885
                  RightToLeft     =   -1  'True
                  TabIndex        =   246
                  Top             =   1890
                  Width           =   630
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   2220
               Index           =   11
               Left            =   11625
               TabIndex        =   251
               TabStop         =   0   'False
               Top             =   45
               Width           =   10890
               _cx             =   19209
               _cy             =   3916
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
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments2 
                  Height          =   1770
                  Left            =   0
                  TabIndex        =   256
                  Top             =   0
                  Width           =   10860
                  _cx             =   19156
                  _cy             =   3122
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
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   12
                  Cols            =   61
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"RsContarct.frx":FC1D
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
               Begin VB.Label LblTotalQasts2 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   255
                  Top             =   1890
                  Width           =   1590
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·œ›⁄« "
                  Height          =   300
                  Index           =   72
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   254
                  Top             =   1890
                  Width           =   1890
               End
               Begin VB.Label LblNotPayed2 
                  Alignment       =   2  'Center
                  Caption         =   "0"
                  Height          =   270
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   253
                  Top             =   1890
                  Width           =   1590
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "€Ì— „”œœ"
                  Height          =   300
                  Index           =   71
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   252
                  Top             =   1890
                  Width           =   1410
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   2220
               Index           =   13
               Left            =   11925
               TabIndex        =   257
               TabStop         =   0   'False
               Top             =   45
               Width           =   10890
               _cx             =   19209
               _cy             =   3916
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
               Begin VSFlex8UCtl.VSFlexGrid grdHistory 
                  Height          =   2070
                  Left            =   5370
                  TabIndex        =   258
                  Top             =   60
                  Width           =   5535
                  _cx             =   9763
                  _cy             =   3651
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
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"RsContarct.frx":10588
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
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   37
            Left            =   5160
            TabIndex        =   101
            Top             =   465
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   270
            Index           =   8
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
            Height          =   270
            Index           =   9
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   270
            Index           =   11
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   465
            Width           =   1410
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   5415
         Index           =   18
         Left            =   11040
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5535
         _cx             =   9763
         _cy             =   9551
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
         Begin VB.TextBox TxtFATYou22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3030
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   293
            Top             =   2040
            Width           =   495
         End
         Begin VB.CheckBox WaterElecValueInVAT 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„Ì«Â Ê«·ﬂÂ—»«¡ ›Ï «·›« "
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   278
            Top             =   870
            Width           =   2505
         End
         Begin VB.CheckBox InsurValueInVAT 
            Alignment       =   1  'Right Justify
            Caption         =   "«· √„Ì‰ ÌœŒ· »«·›« "
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   276
            Top             =   870
            Width           =   1695
         End
         Begin VB.TextBox TxtFATYou2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   270
            Top             =   2910
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox Contract_period_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   266
            Top             =   2760
            Width           =   495
         End
         Begin VB.ComboBox Contract_period 
            Height          =   315
            ItemData        =   "RsContarct.frx":10627
            Left            =   3000
            List            =   "RsContarct.frx":10631
            RightToLeft     =   -1  'True
            TabIndex        =   265
            Top             =   2760
            Width           =   975
         End
         Begin VB.CheckBox CommiValueInVAT 
            Alignment       =   1  'Right Justify
            Caption         =   "«·”⁄Ì ÌœŒ· »«·›« "
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   229
            Top             =   1110
            Width           =   1695
         End
         Begin VB.TextBox TxtTotalValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   227
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox TxtIncresYearRate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   2400
            Width           =   705
         End
         Begin VB.TextBox TxtIncresYearValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   3120
            Width           =   705
         End
         Begin VB.TextBox TxtFATValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   222
            Top             =   2040
            Width           =   1065
         End
         Begin VB.TextBox TxtFATYou 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3570
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TxtNetValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   2040
            Width           =   1185
         End
         Begin VB.TextBox TxtMiniRentValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   5040
            Width           =   1140
         End
         Begin VB.TextBox TxtEnternet 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   4080
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox TxtPayAmini 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   3360
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TxtTotalContract 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1110
            Width           =   705
         End
         Begin VB.TextBox TxtInsuranceValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1110
            Width           =   705
         End
         Begin VB.TextBox TxtWater 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1470
            Width           =   705
         End
         Begin VB.TextBox TxtElectricity 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1470
            Width           =   705
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   210
            Width           =   705
         End
         Begin VB.TextBox TxtCommiValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   1110
            Width           =   705
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1470
            Width           =   705
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   600
            Width           =   705
         End
         Begin VB.TextBox TxtOutOffice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1470
            Width           =   705
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   315
            Left            =   480
            TabIndex        =   8
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   240
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker StrDate 
            Height          =   270
            Left            =   1920
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal FromdateH 
            Height          =   255
            Left            =   3360
            TabIndex        =   18
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   600
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdCus 
            Height          =   345
            Left            =   0
            TabIndex        =   106
            Top             =   240
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   609
            ButtonStyle     =   1
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
            ButtonImage     =   "RsContarct.frx":1063F
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker EndDate 
            Height          =   270
            Left            =   1920
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal TodateH 
            Height          =   255
            Left            =   3360
            TabIndex        =   20
            Top             =   3120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   1335
            Index           =   3
            Left            =   0
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   3480
            Width           =   5535
            _cx             =   9763
            _cy             =   2355
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
            Begin VB.TextBox TxtOldRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3600
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtoldCommi 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   585
            End
            Begin VB.TextBox TxtOldElectric 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   240
               Width           =   585
            End
            Begin VB.TextBox TxtOldWater 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   240
               Width           =   585
            End
            Begin VB.TextBox balanceDes 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               Top             =   600
               Width           =   2490
            End
            Begin MSComCtl2.DTPicker balanceDate 
               Height          =   255
               Left            =   3360
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   960
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               Format          =   64225283
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal balanceDateH 
               Height          =   255
               Left            =   3360
               TabIndex        =   26
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂÂ—»«¡"
               Height          =   195
               Index           =   40
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   240
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Ì«Â"
               Height          =   195
               Index           =   41
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   240
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰  «—ÌŒ"
               Height          =   285
               Index           =   43
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   840
               Width           =   690
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   195
               Index           =   44
               Left            =   2475
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   720
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Œœ„« "
               Height          =   195
               Index           =   39
               Left            =   435
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   240
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·⁄ﬁœ"
               Height          =   195
               Index           =   38
               Left            =   4635
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   240
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—’Ìœ „ »ﬁÌ ⁄·Ï «·„” «Ã—"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   59
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   0
               Width           =   2550
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   615
            Index           =   4
            Left            =   240
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   4800
            Width           =   2655
            _cx             =   4683
            _cy             =   1085
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
            AutoSizeChildren=   7
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
            Begin VB.TextBox TxtOldInsurance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " √„Ì‰"
               Height          =   195
               Index           =   42
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   240
               Width           =   510
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—’Ìœ ··„” «Ã— „”œœ „”»ﬁ«"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   60
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   0
               Width           =   1950
            End
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "„œÂ «·⁄ﬁœ "
            Height          =   375
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   267
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   68
            Left            =   525
            RightToLeft     =   -1  'True
            TabIndex        =   224
            Top             =   1800
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   67
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   223
            Top             =   1800
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   66
            Left            =   3165
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   65
            Left            =   4365
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   1800
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "√ﬁ· ﬁÌ„…  √ÃÌ—Ì…"
            Height          =   195
            Index           =   55
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   5040
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„« "
            Height          =   195
            Index           =   28
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   3960
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„« "
            Height          =   195
            Index           =   24
            Left            =   465
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   3600
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Â«ÌÂ  «·⁄ﬁœ"
            Height          =   405
            Index           =   23
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·«ÌÃ«—"
            Height          =   195
            Index           =   6
            Left            =   4665
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   1110
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " √„Ì‰"
            Height          =   195
            Index           =   19
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   1110
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„Ì«Â"
            Height          =   195
            Index           =   20
            Left            =   2265
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   1470
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÂ—»«¡/€«“"
            Height          =   315
            Index           =   21
            Left            =   4725
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   1470
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·„” √Ã—"
            Height          =   285
            Index           =   5
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "»œ«ÌÂ «·⁄ﬁœ"
            Height          =   285
            Index           =   22
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   2400
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”⁄Ì "
            Height          =   285
            Index           =   25
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   1440
            Visible         =   0   'False
            Width           =   450
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„« "
            Height          =   195
            Index           =   27
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„…"
            Height          =   285
            Index           =   31
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "“Ì«œ… ”‰ÊÌ…%"
            Height          =   195
            Index           =   30
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   2400
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰œÊ»"
            Height          =   285
            Index           =   37
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”⁄Ì Œ«—ÃÌ"
            Height          =   435
            Index           =   47
            Left            =   795
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1350
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " »Ì«‰«  «·⁄ﬁœ"
            ForeColor       =   &H00C00000&
            Height          =   405
            Index           =   61
            Left            =   4050
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   30
            Width           =   1290
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   2265
         Index           =   15
         Left            =   11040
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1230
         Width           =   5535
         _cx             =   9763
         _cy             =   3995
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
         AutoSizeChildren=   7
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
         Begin VB.TextBox TxtElectAccount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   1065
            Width           =   2025
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   135
            MaxLength       =   50
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   141
            Top             =   1800
            Width           =   4410
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1425
            Width           =   825
         End
         Begin VB.TextBox TxtMeterValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   1065
            Width           =   1305
         End
         Begin VB.TextBox TxtMeterCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   705
            Width           =   1305
         End
         Begin VB.ComboBox DcbRentType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "RsContarct.frx":109D9
            Left            =   120
            List            =   "RsContarct.frx":109E3
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   705
            Width           =   2025
         End
         Begin VB.TextBox TxtSearch 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   60
            Width           =   825
         End
         Begin MSDataListLib.DataCombo DcbIqara 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
            Top             =   60
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitNo 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   405
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitType 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   405
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcsupplier 
            Height          =   315
            Left            =   120
            TabIndex        =   143
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   1425
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ Õ”«» «·ﬂÂ—»«¡"
            Height          =   435
            Index           =   48
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   945
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Ê’›  «·ÊÕœ…"
            Height          =   285
            Index           =   33
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   1815
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·„ —"
            Height          =   195
            Index           =   17
            Left            =   4245
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   1065
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·«„ «—"
            Height          =   195
            Index           =   18
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   705
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÊÕœ…"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   14
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   405
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ÊÕœ…"
            Height          =   285
            Index           =   15
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «· √ÃÌ—"
            Height          =   195
            Index           =   16
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   705
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄ﬁ«—"
            Height          =   195
            Index           =   4
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   60
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·„«·ﬂ"
            Height          =   165
            Index           =   1
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   1425
            Width           =   810
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   2415
         Index           =   1
         Left            =   5040
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   4800
         Width           =   6015
         _cx             =   10610
         _cy             =   4260
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
         Begin VSFlex8UCtl.VSFlexGrid UnitsGrid 
            Height          =   1605
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   5685
            _cx             =   10028
            _cy             =   2831
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   25
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"RsContarct.frx":10A00
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   0
            Left            =   4890
            TabIndex        =   154
            Top             =   2040
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":10DCC
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÊÕœ«  «·„œ„Ã…"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   45
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   120
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   2415
         Index           =   2
         Left            =   0
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   4800
         Width           =   5055
         _cx             =   8916
         _cy             =   4260
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
         Begin VB.Frame Frame4 
            Caption         =   "› —… «·ﬁÌ«”"
            Enabled         =   0   'False
            Height          =   645
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   290
            Top             =   0
            Visible         =   0   'False
            Width           =   11415
            Begin MSComCtl2.DTPicker txtDateK 
               Height          =   285
               Left            =   3570
               TabIndex        =   291
               TabStop         =   0   'False
               Top             =   180
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               Format          =   64225283
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker txtDateK2 
               Height          =   285
               Left            =   90
               TabIndex        =   292
               TabStop         =   0   'False
               Top             =   180
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               Format          =   64225283
               CurrentDate     =   41640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
            Height          =   1605
            Left            =   120
            TabIndex        =   157
            Top             =   360
            Width           =   4845
            _cx             =   8546
            _cy             =   2831
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   31
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"RsContarct.frx":11366
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   13
            Left            =   3960
            TabIndex        =   158
            Top             =   2040
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":117FB
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰«œÌ»"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   53
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   120
            Width           =   1770
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1020
         Left            =   7320
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   8730
         Width           =   5280
         _cx             =   9313
         _cy             =   1799
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
         BorderWidth     =   1
         ChildSpacing    =   1
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
         Begin VB.CheckBox CheckFp 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ»«⁄Â »’„Â"
            Height          =   195
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   274
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   4365
            TabIndex        =   161
            Top             =   555
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":11D95
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   2640
            TabIndex        =   162
            Top             =   555
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":1212F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   3525
            TabIndex        =   163
            Top             =   555
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":124C9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   1815
            TabIndex        =   164
            Top             =   555
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":12863
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   990
            TabIndex        =   165
            Top             =   555
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":12BFD
            ColorButton     =   14871017
            Alignment       =   0
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   120
            TabIndex        =   166
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   555
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":13197
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   5085
            TabIndex        =   167
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   105
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "RsContarct.frx":13531
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   555
            TabIndex        =   168
            Top             =   1035
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   582
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
            ButtonImage     =   "RsContarct.frx":138CB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   1290
            TabIndex        =   289
            Top             =   0
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   150
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   375
            Width           =   510
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   1500
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   345
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   345
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   1335
         Index           =   6
         Left            =   0
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   7200
         Width           =   4455
         _cx             =   7858
         _cy             =   2355
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
         Begin VB.CommandButton Command12 
            Caption         =   "≈‰‘«¡ «·ﬁÌœ"
            Height          =   255
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Õ–› «·ﬁÌœ"
            Height          =   255
            Left            =   1140
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "ﬂ‘› Õ”«»"
            Height          =   255
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   960
            Width           =   960
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   255
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   240
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "RsContarct.frx":13C65
            Height          =   315
            Left            =   0
            TabIndex        =   228
            Top             =   -240
            Visible         =   0   'False
            Width           =   3450
            _ExtentX        =   6085
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
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   270
            Left            =   1050
            TabIndex        =   233
            TabStop         =   0   'False
            Top             =   960
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker FrmDate 
            Height          =   270
            Left            =   2760
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   195
            Index           =   70
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   960
            Width           =   270
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   195
            Index           =   69
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   235
            Top             =   960
            Width           =   270
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   62
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   0
            Width           =   1890
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   1095
         Index           =   7
         Left            =   3120
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   8640
         Width           =   4215
         _cx             =   7435
         _cy             =   1931
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
         Begin VB.CommandButton Command6 
            Caption         =   "⁄—÷ ”‰œ«  «·ﬁ»÷ «·Œ«’Â »«·⁄ﬁœ"
            Height          =   495
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton Command7 
            Caption         =   " ‰«“·"
            Height          =   375
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton CmDRENEW 
            Caption         =   " ÃœÌœ/ „œÌœ"
            Height          =   375
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "«Œ·«¡"
            Height          =   375
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   " ÃœÌœ/ „œÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "ÿ»«⁄Â «·⁄ﬁœ"
            Height          =   375
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   120
            Width           =   975
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   11
            Left            =   2040
            TabIndex        =   186
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   615
         Index           =   9
         Left            =   12720
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   9120
         Width           =   3855
         _cx             =   6800
         _cy             =   1085
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
         AutoSizeChildren=   7
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
         Begin MSComCtl2.DTPicker FromdateO 
            Height          =   270
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal FromdateHO 
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   270
            Left            =   0
            TabIndex        =   213
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   64225283
            CurrentDate     =   41640
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "»œ«Ì… «·⁄ﬁœ «·«’·Ì"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   64
            Left            =   870
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   0
            Width           =   2835
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "»œ«ÌÂ «·⁄ﬁœ"
            Height          =   285
            Index           =   46
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   240
            Width           =   810
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   1455
         Index           =   5
         Left            =   6960
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   7320
         Width           =   4095
         _cx             =   7223
         _cy             =   2566
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
         Begin VB.TextBox TxtFATValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   260
            Top             =   1080
            Width           =   705
         End
         Begin VB.TextBox TxtServce 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   285
            Width           =   705
         End
         Begin VB.TextBox TxtElectricityValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   675
            Width           =   705
         End
         Begin VB.TextBox TxtWaterValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   675
            Width           =   705
         End
         Begin VB.TextBox TxtCommValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   285
            Width           =   705
         End
         Begin VB.TextBox TxtRetValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   675
            Width           =   705
         End
         Begin VB.TextBox TxtInstrunceValue2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… „÷«›…"
            Height          =   210
            Index           =   73
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   261
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Œœ„« "
            Height          =   210
            Index           =   56
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÂ—»«¡"
            Height          =   225
            Index           =   54
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„Ì«Â"
            Height          =   225
            Index           =   52
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”⁄Ì"
            Height          =   210
            Index           =   51
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«ÌÃ«—"
            Height          =   225
            Index           =   50
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " √„Ì‰"
            Height          =   210
            Index           =   49
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " ›«’Ì· «·⁄—»Ê‰ "
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   58
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   0
            Width           =   2970
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   975
         Index           =   8
         Left            =   0
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   8520
         Width           =   3135
         _cx             =   5530
         _cy             =   1720
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
         Begin VB.CommandButton Command10 
            Caption         =   "› Õ ”‰œ ﬁ»÷"
            Height          =   255
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   287
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Optx 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ»ﬁ« ··„‰œÊ»"
            Height          =   195
            Index           =   4
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Optx 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ»ﬁ« ··„” √Ã—"
            Height          =   195
            Index           =   3
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Optx 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ»ﬁ« ··„«·ﬂ"
            Height          =   195
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Optx 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ»ﬁ« ··⁄ﬁ«—"
            Height          =   195
            Index           =   1
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Optx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
            Height          =   195
            Index           =   0
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ŒÌ«—«  «·⁄—÷"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   63
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   0
            Width           =   990
         End
      End
      Begin ImpulseButton.ISButton BtnUpdate6 
         Height          =   315
         Left            =   3600
         TabIndex        =   214
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   600
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "RsContarct.frx":13C7A
         DrawFocusRectangle=   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   270
         Left            =   6840
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         Format          =   64225283
         CurrentDate     =   41640
      End
      Begin MSDataListLib.DataCombo AccountVat2 
         Height          =   315
         Left            =   12840
         TabIndex        =   271
         Top             =   720
         Visible         =   0   'False
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„Â «·Œ’„"
         Height          =   195
         Index           =   78
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   281
         Top             =   3480
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„"
         Height          =   195
         Index           =   77
         Left            =   14565
         RightToLeft     =   -1  'True
         TabIndex        =   280
         Top             =   3540
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ«  «·ﬁÌœ"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   76
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   273
         Top             =   8040
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‘—Êÿ Œ«’Â ··⁄ﬁœ"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   211
         Top             =   7320
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·⁄ﬁœ"
         Height          =   195
         Index           =   3
         Left            =   15480
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·⁄ﬁœ"
         Height          =   285
         Index           =   7
         Left            =   15360
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·⁄ﬁœ"
         Height          =   285
         Index           =   12
         Left            =   13200
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   1020
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·›—⁄"
         Height          =   195
         Index           =   32
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁÌ„… «·⁄—»Ê‰"
         Height          =   285
         Index           =   61
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ”‰œ «·⁄—»Ê‰"
         Height          =   285
         Index           =   60
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‘—Êÿ Œ«’Â ··⁄ﬁœ"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   57
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Index           =   26
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   2520
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·œ›⁄« "
      Height          =   285
      Index           =   10
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "RSContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Dim commisiontype As Integer
Dim ScreenNameArabic As String
Dim ScreenNameEnglish As String
Dim Subvat As Double
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim Iqar As Double
Dim ownerid As Double
Dim AmolaValues As Double
Dim cSearch  As clsDCboSearch
Dim Account_Code_dynamic80 As String
Dim Account_Code_dynamic81 As String
Dim Account_Code_dynamic82 As String
Dim Account_Code_dynamic83 As String


Dim Account_Code_dynamic84 As String
Dim Account_Code_dynamic85 As String
Dim Account_Code_dynamic92 As String
Dim Account_Code_dynamic59 As String
Dim Account_Code_dynamic123 As String
Dim Account_Code_dynamic125 As String

Dim Account_Code_dynamic154 As String
Dim Account_Code_dynamic155 As String
Dim Account_Code_dynamic156 As String

Public RereivID As Double
Dim InstalNo As Integer
Dim UonitStatus As Integer
Dim hijriorJerojian As Integer
Dim FlagContrNew As Boolean
Dim FlagContrNew2 As Boolean
Dim mCreateEntryManual As Boolean
Dim mchkAllowEditPaymentCont As Boolean
Dim mCanEdit As Boolean
Function checkallocation2(ContNo As Double, Optional des As String) As Boolean
Dim str As String
Dim RsDetails1 As ADODB.Recordset
Dim i As Double
 des = ""
str = " SELECT     TOP 100 PERCENT dbo.TblContractInstallments.ContNo, dbo.TblContract.NoteSerial1, dbo.tblContractInsAllocations1.transID"
str = str & "   FROM         dbo.tblContractInsAllocations1 INNER JOIN"
str = str & "                        dbo.tblContractInsAllocationsDetails2 ON dbo.tblContractInsAllocations1.transID = dbo.tblContractInsAllocationsDetails2.transID INNER JOIN"
str = str & "                        dbo.TblContractInstallments ON dbo.tblContractInsAllocationsDetails2.Installid = dbo.TblContractInstallments.id INNER JOIN"
str = str & "                        dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
str = str & "  GROUP BY dbo.TblContractInstallments.ContNo, dbo.TblContract.NoteSerial1, dbo.tblContractInsAllocations1.transID"
str = str & "  Having (dbo.TblContractInstallments.ContNo = " & ContNo & ")"
str = str & "  ORDER BY dbo.tblContractInsAllocations1.transID"
Set RsDetails1 = New ADODB.Recordset
    RsDetails1.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If RsDetails1.RecordCount > 0 Then
    For i = 1 To RsDetails1.RecordCount
    des = des & " , " & (IIf(IsNull(RsDetails1("transID").value), "", RsDetails1("transID").value))
    RsDetails1.MoveNext
    Next i
         checkallocation2 = True
         Else
         checkallocation2 = False
    End If



End Function

Function checkAllocations(ContNo As Double, Optional des As String) As Boolean
Dim str As String
Dim RsDetails1 As ADODB.Recordset
Dim i As Double
 
des = ""
str = "SELECT     TOP 100 PERCENT dbo.TblContractInstallments.ContNo, dbo.tblContractInsAllocations.transID"
str = str & "  FROM         dbo.tblContractInsAllocations INNER JOIN"
str = str & "                        dbo.tblContractInsAllocationsDetails ON dbo.tblContractInsAllocations.transID = dbo.tblContractInsAllocationsDetails.transID INNER JOIN"
str = str & "                        dbo.TblContractInstallments ON dbo.tblContractInsAllocationsDetails.Installid = dbo.TblContractInstallments.id"
str = str & "  Where (dbo.TblContractInstallments.ContNo = " & ContNo & ")"
str = str & "  ORDER BY dbo.tblContractInsAllocations.transID"
Set RsDetails1 = New ADODB.Recordset
    RsDetails1.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If RsDetails1.RecordCount > 0 Then
    For i = 1 To RsDetails1.RecordCount
    des = des & " , " & (IIf(IsNull(RsDetails1("transID").value), "", RsDetails1("transID").value))
    RsDetails1.MoveNext
    Next i
         checkAllocations = True
         Else
         checkAllocations = False
    End If

End Function

Sub GetUonitStatus()

    Dim RsDetails1 As ADODB.Recordset
    Dim StrSQL As String

    Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT   Status  from  TblAqarDetai where id =" & val(DcbUnitNo.BoundText) & ""
    RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If RsDetails1.RecordCount > 0 Then
        UonitStatus = val(IIf(IsNull(RsDetails1("Status").value), "", RsDetails1("Status").value))
    End If
    
End Sub
Sub SaveUoitInformation()

    Dim RsDetails1 As ADODB.Recordset
    Dim StrSQL, Msg As String
    
    Msg = ""
 
    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = Msg & "Work was guaranteed catch filtering No."
        Msg = Msg & CHR(13) & TxtNoteSerial1.Text
        Msg = Msg & "Start From"
        Msg = Msg & CHR(13) & FromDateH.value & "Approved" & StrDate.value
        Msg = Msg & CHR(13)
        Msg = Msg & "End With "
        Msg = Msg & CHR(13) & todateH.value & "Approved" & EndDate.value
        Msg = Msg & CHR(13)
        Msg = Msg & "Contract Value"
        Msg = Msg & TxtTotalContract.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " COmm value "
        Msg = Msg & TxtCommiValue.Text
        Msg = Msg & CHR(13)
        Msg = Msg & "Insurance Value "
        Msg = Msg & TxtInsuranceValue.Text
        Msg = Msg & CHR(13)
        Msg = Msg & "  Water Value "
        Msg = Msg & TxtWater.Text
        Msg = Msg & CHR(13)
        Msg = Msg & "Electricity Value"
        Msg = Msg & TxtElectricity.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " Services Value "
        Msg = Msg & TxtPhone.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " Comm Out Value "
        Msg = Msg & TxtOutOffice.Text
        Msg = Msg & CHR(13)
    Else
        Msg = Msg & "   „ ⁄„· ⁄ﬁœ —ﬁ„  "
        Msg = Msg & CHR(13) & TxtNoteSerial1.Text
        Msg = Msg & "Ì»œ√ „‰  "
        Msg = Msg & CHR(13) & FromDateH.value & "«·„Ê«›ﬁ" & StrDate.value
        Msg = Msg & CHR(13)
        Msg = Msg & "ÊÌ‰ ÂÌ "
        Msg = Msg & CHR(13) & todateH.value & "«·„Ê«›ﬁ" & EndDate.value
        Msg = Msg & CHR(13)
        Msg = Msg & "ﬁÌ„… «·⁄ﬁœ "
        Msg = Msg & TxtTotalContract.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… «·”⁄Ì "
        Msg = Msg & TxtCommiValue.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… «· «„Ì‰ "
        Msg = Msg & TxtInsuranceValue.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… «·„Ì«Â "
        Msg = Msg & TxtWater.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… «·ﬂÂ—»«¡ "
        Msg = Msg & TxtElectricity.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… «·Œœ„«  "
        Msg = Msg & TxtPhone.Text
        Msg = Msg & CHR(13)
        Msg = Msg & " ﬁÌ„… ”⁄Ì Œ«—ÃÌ "
        Msg = Msg & TxtOutOffice.Text
        Msg = Msg & CHR(13)
    End If
        
    Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
    RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    RsDetails1.AddNew
    RsDetails1("BranchId").value = val(Dcbranch.BoundText)
    RsDetails1("CusID").value = val(dcCustomer.BoundText)
    RsDetails1("UnitNo").value = val(DcbUnitNo.BoundText)
    RsDetails1("UnitStatus").value = UonitStatus
    RsDetails1("Des").value = Msg
    RsDetails1("RecDate").value = ContDate.value
    RsDetails1("RecDateH").value = RecorddateH.value
    RsDetails1("NoteID").value = Null
    RsDetails1("ContNo").value = val(Me.TxtContNo.Text)
    RsDetails1("FilterNo").value = Null
    RsDetails1("OrderMaint").value = Null
    RsDetails1("LocOrderMaint").value = Null
    RsDetails1.update

End Sub
Function saveinstdetailforpart2()
    If commisiontype = 1 Then Exit Function
        
        Dim StrSQL  As String
        Dim RsDetails1 As New ADODB.Recordset
        Dim Countsofall As Double
        Dim j As Integer
        Dim SngAllValue As Single
 
        Dim IntNoOFQast As Integer
        Dim IntRes As Integer
        Dim SngOnePor As Single
        Dim FirstDate As Date
        Dim PreDate As Date
        Dim NewDate As Date
        Dim DateInterval As String
        Dim NewDateH As String
        Dim endpartdays As Integer
        Dim PreDateH As String
        Dim hijriorJerojian As Integer
        Dim LastDate As Date
        Dim LastDateH As String
        Dim FirstDate1 As Date
        Dim FirstDateH1 As String
        Dim DateNumber As Integer
 
        Dim watervalue As Double
        Dim Electricity As Double
        Dim noOfRemaindays As Integer
        Dim noOfRemaindays1 As Integer
        Dim MonthLastDay1 As Date
        Dim onedayvale As Double
        Dim onedayRentValue As Double
        Dim onedayCommissions As Double
        Dim onedayInsurance As Double
        Dim onedayWater As Double
        Dim onedayElectric As Double
        Dim onedayTelandNet As Double
  
        StrSQL = "Delete From tblContractInsAllocationsDetails1 Where transid is null and  ContractFlag=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
        StrSQL = "SELECT  *  from dbo.tblContractInsAllocationsDetails1 Where (1 = -1)"
        RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim noOfInstallments As Integer
If SystemOptions.WorkWithFirstInstallOnly = True Then
noOfInstallments = 1
Else
noOfInstallments = GridInstallments.Rows - 1

End If
        If Opt(0).value = False Then Exit Function
            Dim i As Integer
            If GridInstallments.Rows = 1 Then Exit Function
                With Me.GridInstallments
                
                    For i = 1 To noOfInstallments
                    'And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked
                       If val(.TextMatrix(i, .ColIndex("value"))) <> 0 Then
                            
                            
                            
If val(TxtPaymentCount) = 1000 Then 'œ›⁄Â Ê«Õœ…
 

RsDetails1.AddNew
                                RsDetails1("ContractFlag").value = Me.TxtContNo.Text
                                RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
                                hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
                                hijriorJerojian = 1
                                RsDetails1("NoteSerial").value = TxtNoteSerial1.Text
                                RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
                                RsDetails1("Installdate").value = (.TextMatrix(i, .ColIndex("Due_Date")))
                                RsDetails1("InstalldateH").value = (.TextMatrix(i, .ColIndex("Due_DateH")))
                                RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
                                RsDetails1("installValue").value = val(.TextMatrix(i, .ColIndex("value")))
                                RsDetails1("RentValue").value = val(.TextMatrix(i, .ColIndex("RentValue")))
                                RsDetails1("Commissions").value = val(.TextMatrix(i, .ColIndex("Commissions")))
                                RsDetails1("Insurance").value = val(.TextMatrix(i, .ColIndex("Insurance")))
                                RsDetails1("Water").value = val(.TextMatrix(i, .ColIndex("Water")))
                                RsDetails1("Electric").value = val(.TextMatrix(i, .ColIndex("Electric")))
                                RsDetails1("TelandNet").value = val(.TextMatrix(i, .ColIndex("TelandNet")))
                                RsDetails1("CusID").value = val(dcCustomer.BoundText)
                                RsDetails1.update
      
      Exit Function
End If
                            
                            
                            
    
                            Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall")))
                            
                        If DcbPeriodsID.ListIndex = 0 Then 'day
                           
                           Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall"))) / 30
                           
                            ElseIf DcbPeriodsID.ListIndex = 1 Then 'month
                               Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall")))
                             ElseIf DcbPeriodsID.ListIndex = 2 Then 'year
                             
                                Countsofall = val(.TextMatrix(i, .ColIndex("Countsofall"))) * 12
                             End If
                             
                  '   Countsofall = 12
                             
                            VBA.Calendar = vbCalGreg
                            LastDate = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_Date"))))
                            LastDate = DateAdd("d", -1, LastDate)
                            VBA.Calendar = vbCalHijri
                            LastDateH = DateAdd("M", Countsofall, (.TextMatrix(i, .ColIndex("Due_DateH"))))
                            LastDateH = DateAdd("d", -1, LastDateH)
                            '«· √ﬂœ «‰ «· «—ÌŒ ·Ì” «Ê· «·‘Â—
                            hijriorJerojian = 1
                            If hijriorJerojian = 1 Then 'jorjian
                                VBA.Calendar = vbCalGreg
                                FirstDate1 = dhFirstDayInMonth(.TextMatrix(i, .ColIndex("Due_Date")))
                                noOfRemaindays1 = DateDiff("D", .TextMatrix(i, .ColIndex("Due_Date")), FirstDate1)
                            Else
                                VBA.Calendar = vbCalHijri
                                FirstDateH1 = dhFirstDayInMonth(.TextMatrix(i, .ColIndex("Due_DateH")))
                                noOfRemaindays1 = DateDiff("D", .TextMatrix(i, .ColIndex("Due_DateH")), FirstDateH1)
                            End If
                            If noOfRemaindays1 = 0 Then GoTo ll
                            hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
                            hijriorJerojian = 1
                            If hijriorJerojian = 1 Then 'jorjian
                                VBA.Calendar = vbCalGreg
                                noOfRemaindays = DateDiff("D", .TextMatrix(i, .ColIndex("Due_Date")), MonthLastDay(.TextMatrix(i, .ColIndex("Due_Date"))))
                            Else
                                VBA.Calendar = vbCalHijri
                                noOfRemaindays = DateDiff("D", .TextMatrix(i, .ColIndex("Due_DateH")), MonthLastDay(.TextMatrix(i, .ColIndex("Due_DateH"))))
                            End If
                            If noOfRemaindays > 0 Then
                                Countsofall = Countsofall - 1
                            End If
                            
                           Dim newDivision As Integer
                           
                           If DcbPeriodsID.ListIndex = 0 Then 'day
                           
                           newDivision = TxtPeriods / 30
                           
                            ElseIf DcbPeriodsID.ListIndex = 1 Then 'month
                               newDivision = TxtPeriods
                             ElseIf DcbPeriodsID.ListIndex = 2 Then 'year
                             
                                newDivision = TxtPeriods * 12
                             End If
                             
                            endpartdays = 30 - noOfRemaindays
                            
                            
                            onedayvale = val(.TextMatrix(i, .ColIndex("value"))) / val(newDivision) / 30
                            
                            onedayRentValue = val(.TextMatrix(i, .ColIndex("RentValue"))) / val(newDivision) / 30
                            onedayCommissions = val(.TextMatrix(i, .ColIndex("Commissions"))) / val(newDivision) / 30
                            onedayInsurance = val(.TextMatrix(i, .ColIndex("Insurance"))) / val(newDivision) / 30
                            onedayWater = val(.TextMatrix(i, .ColIndex("Water"))) / val(newDivision) / 30
                            onedayElectric = val(.TextMatrix(i, .ColIndex("Electric"))) / val(newDivision) / 30
                            onedayTelandNet = val(.TextMatrix(i, .ColIndex("TelandNet"))) / val(newDivision) / 30
                            
                            
                            '*****************part one of month
                            If noOfRemaindays > 0 Then
                                VBA.Calendar = vbCalGreg
                                NewDate = (.TextMatrix(i, .ColIndex("Due_Date")))
                                NewDateH = Format((.TextMatrix(i, .ColIndex("Due_DateH"))), "DD/MM/YYYY")
                                RsDetails1.AddNew
                                RsDetails1("ContractFlag").value = Me.TxtContNo.Text
                                RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
                                hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
                                hijriorJerojian = 1
                                RsDetails1("NoteSerial").value = TxtNoteSerial1.Text
                                RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
                                RsDetails1("Installdate").value = (NewDate)
                                RsDetails1("InstalldateH").value = NewDateH
                                RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
                                RsDetails1("installValue").value = Round(onedayvale * noOfRemaindays, 2)
                                RsDetails1("RentValue").value = Round(onedayRentValue * noOfRemaindays, 2)
                                RsDetails1("Commissions").value = Round(onedayCommissions * noOfRemaindays, 2)
                                RsDetails1("Insurance").value = Round(onedayInsurance * noOfRemaindays, 2)
                                RsDetails1("Water").value = Round(onedayWater * noOfRemaindays, 2)
                                RsDetails1("Electric").value = Round(onedayElectric * noOfRemaindays, 2)
                                RsDetails1("TelandNet").value = Round(onedayTelandNet * noOfRemaindays, 2)
                                RsDetails1("CusID").value = val(dcCustomer.BoundText)
                                RsDetails1.update
                            End If
                            '***********************end of first part*******************************
                            VBA.Calendar = vbCalGreg
                            NewDate = MonthLastDay(.TextMatrix(i, .ColIndex("Due_Date")))
                            VBA.Calendar = vbCalHijri
                            NewDateH = MonthLastDay(.TextMatrix(i, .ColIndex("Due_DateH")))
                             
                            VBA.Calendar = vbCalGreg
                             
                            NewDate = DateAdd("D", 1, NewDate)
                            VBA.Calendar = vbCalHijri
                            NewDateH = DateAdd("D", 1, NewDateH)
ll:
                            If noOfRemaindays = 0 Then
                                VBA.Calendar = vbCalGreg
                                NewDate = (.TextMatrix(i, .ColIndex("Due_Date")))
                                NewDateH = Format((.TextMatrix(i, .ColIndex("Due_DateH"))), "DD/MM/YYYY")
                            End If
         
                            For j = 1 To Countsofall
                                RsDetails1.AddNew
                                RsDetails1("ContractFlag").value = Me.TxtContNo.Text

                                RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
                                hijriorJerojian = 1
                                'hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
                                If j = 1 Then
                                Else
                                    VBA.Calendar = vbCalGreg
                                    PreDate = NewDate
                                    If hijriorJerojian = 1 Then 'jorijan
                                        VBA.Calendar = vbCalGreg
                                        NewDate = DateAdd("m", 1, NewDate)
                                        NewDateH = ToHijriDate(NewDate)
                                    End If
                                    PreDateH = NewDateH
                                    If hijriorJerojian = 0 Then 'hijri
                                        VBA.Calendar = vbCalHijri
                                        NewDateH = (DateAdd("m", 1, NewDateH))
                                        VBA.Calendar = vbCalGreg
                                        NewDate = ToGregorianDate(NewDateH)
                                    End If
                                End If
                                RsDetails1("NoteSerial").value = TxtNoteSerial1.Text
                                RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
                                VBA.Calendar = vbCalGreg
                                RsDetails1("Installdate").value = (NewDate)
                                RsDetails1("InstalldateH").value = NewDateH
                                RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
                                
                                If DcbPeriodsID.ListIndex = 0 Then 'day
                           
                      '     .TextMatrix(i, .ColIndex("Countsofall")) = val(.TextMatrix(i, .ColIndex("Countsofall"))) / 30
                           
                            ElseIf DcbPeriodsID.ListIndex = 1 Then 'month
                       '        .TextMatrix(i, .ColIndex("Countsofall")) = val(.TextMatrix(i, .ColIndex("Countsofall")))
                             ElseIf DcbPeriodsID.ListIndex = 2 Then 'year
                             
                      '          .TextMatrix(i, .ColIndex("Countsofall")) = val(.TextMatrix(i, .ColIndex("Countsofall"))) * 12
                             End If
                          
                          Dim increasersalim As Integer
                          If noOfRemaindays1 = 0 Then
                          increasersalim = 0
                          Else
                          increasersalim = 1
                          End If
                                
                                RsDetails1("installValue").value = Round(val(.TextMatrix(i, .ColIndex("value"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("RentValue").value = Round(val(.TextMatrix(i, .ColIndex("RentValue"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("Commissions").value = Round(val(.TextMatrix(i, .ColIndex("Commissions"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("Insurance").value = Round(val(.TextMatrix(i, .ColIndex("Insurance"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("Water").value = Round(val(.TextMatrix(i, .ColIndex("Water"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("Electric").value = Round(val(.TextMatrix(i, .ColIndex("Electric"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("TelandNet").value = Round(val(.TextMatrix(i, .ColIndex("TelandNet"))) / (Countsofall + increasersalim), 2)
                                RsDetails1("CusID").value = val(dcCustomer.BoundText)
                                RsDetails1.update
                            Next j
                            '*****************  Last part of month
                            If noOfRemaindays1 = 0 Then GoTo xx
                            If noOfRemaindays > 0 Then
                                If hijriorJerojian = 1 Then ' jorjia then
                                    VBA.Calendar = vbCalGreg
                                    NewDate = DateAdd("m", 1, NewDate)
                                    NewDateH = ToHijriDate(NewDate)
                                Else
                                    VBA.Calendar = vbCalHijri
                                    NewDateH = DateAdd("m", 1, NewDateH)
                                    VBA.Calendar = vbCalGreg
                                    NewDate = ToGregorianDate(NewDateH)
                                End If
                                'Calendar = vbCalGreg
                                'NewDateH = ToHijriDate(NewDate)
                                If hijriorJerojian = 1 Then 'jorjian
                                    VBA.Calendar = vbCalGreg
                                    noOfRemaindays = DateDiff("D", NewDate, LastDate)
                                Else
                                    VBA.Calendar = vbCalHijri
                                    noOfRemaindays = DateDiff("D", NewDateH, LastDateH)
                                End If
                                noOfRemaindays = noOfRemaindays + 1
                                RsDetails1.AddNew
                                RsDetails1("ContractFlag").value = Me.TxtContNo.Text
                                RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
                                hijriorJerojian = (.TextMatrix(i, .ColIndex("hijri")))
                                RsDetails1("NoteSerial").value = TxtNoteSerial1.Text
                                RsDetails1("Installid").value = val(.TextMatrix(i, .ColIndex("Installid")))
                                VBA.Calendar = vbCalGreg
                                RsDetails1("Installdate").value = NewDate
                                RsDetails1("InstalldateH").value = NewDateH
                                RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
                                RsDetails1("installValue").value = Round(onedayvale * endpartdays, 2)
                                RsDetails1("RentValue").value = Round(onedayRentValue * endpartdays, 2)
                                RsDetails1("Commissions").value = Round(onedayCommissions * endpartdays, 2)
                                RsDetails1("Insurance").value = Round(onedayInsurance * endpartdays, 2)
                                RsDetails1("Water").value = Round(onedayWater * endpartdays, 2)
                                RsDetails1("Electric").value = Round(onedayElectric * endpartdays, 2)
                                RsDetails1("TelandNet").value = Round(onedayTelandNet * endpartdays, 2)
                                RsDetails1("CusID").value = val(dcCustomer.BoundText)
                                RsDetails1.update
                            End If
                            '*********************************************************************
xx:
                        Else
                            'Cn.Execute " update  TblContractInstallments set  allocations=0 where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
                        End If
                    Next i
                    RsDetails1.Close
                End With
                '**********************************************************************************************

End Function

Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim i As Integer
    
    On Error GoTo ErrTrap


Dim StrSQL As String
Dim des As String

If checkallocation2(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «Ì—«œ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If


If checkAllocations(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «” Õﬁ«ﬁ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If




    If ChekClodePeriod(StrDate.value) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
        Else
            MsgBox "Please Change Date Becouse This is Period is Closed"
        End If
        Exit Sub
    End If
              

    DTPicker1.value = Date
    Dim FDate As String
    FDate = ToHijriDate(DTPicker1.value)
    If ChkRenew.value = vbChecked Then
        MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
        Exit Sub
    End If

    If checkContractTransactions(val(TxtContNo.Text)) = True Then
        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰ Õ–›…", vbCritical
        Exit Sub
    End If

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    'Dim StrSQL  As String
    If TxtContNo.Text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        If MSGType = vbYes Then
            If TxtContNoOld.Text <> "" Then
                Cn.Execute "  update TblContract  Set Renew =0    Where ContNo =" & val(TxtContNoOld.Text)
            End If
            With UnitsGrid
                For i = 1 To .Rows - 1
                    If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
                        Cn.Execute "  update TblAqarDetai  Set FilterDateH='" & FDate & "',FilterDate=" & SQLDate(DTPicker1.value, True) & ", Status = 0 ,customerid=0  Where id =" & val(.TextMatrix(i, .ColIndex("id")))
                     End If
                Next i
            End With
            DleteUnit
            If val(TxtNotID.Text) > 0 Then
                Cn.Execute "Update Notes set PayedOrBon=Null where NoteID=" & val(TxtNotID.Text) & ""
            End If
            DeleteJE
            Cn.Execute "Update TblContract set Renew=0 where ContNo=" & val(TxtContNoOld.Text) & ""
            StrSQL = "Delete From TblUnitNoInformation Where ContNo =" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblIqrMerg Where Cont=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblCOntractSales Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From tblContractInsAllocationsDetails1 Where ContractFlag=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblContractDet Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblContractInstallments Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "delete From Notes where NoteID=" & val(Me.TxtNoteID.Text) ' Val(rs("Transaction_ID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute "  update TblAqarDetai  Set FilterDateH='" & FDate & "',FilterDate=" & SQLDate(DTPicker1.value, True) & ", Status = 0 ,customerid=0  Where id =" & val(DcbUnitNo.BoundText)
            Cn.Execute "  update TblAqarDetai  Set ContID=" & val(TxtContNo.Text) & "  Where id =" & val(DcbUnitNo.BoundText)
            RsSavRec.find "ContNo=" & val(Me.TxtContNo.Text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            RsSavRec.Resync
            CuurentLogdata ("D")
            BtnLast_Click
           ' BtnNext_Click
            FillGridWithData
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtContNo.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
           LabCurrRec.Caption = 0
    LabCountRec.Caption = 0
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    DCboUserName.BoundText = user_id
    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Public Sub BtnLast_Click()
    
    On Error GoTo ErrTrap
    
    Dim My_SQL As String
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtContNo.Text)
        Me.TxtModFlg.Text = "R"
    End If
    
    My_SQL = " select * from TblContract "
    If SystemOptions.usertype = UserAdminAll Then
        My_SQL = My_SQL & " where   1<>-1"
    Else
        My_SQL = My_SQL & " where   Branch_NO=" & Current_branch
    End If
    
'    If RereivID <> 0 Then
'        My_SQL = My_SQL & "  and ContNo=" & RereivID & ""
'    End If
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If RereivID <> 0 Then
'        FindRec RereivID
'    End If
    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
  Dim FirstPeriod As Date
  getFirstPeriodDateInthisYear FirstPeriod
  FrmDate.value = FirstPeriod
  ToDate.value = Date
BegnieWork:
    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
Dim StrSQL As String
Dim des As String

If checkallocation2(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «Ì—«œ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If

If checkAllocations(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «” Õﬁ«ﬁ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If

 


DcbUnitNo.Enabled = False
DcbUnitType.Enabled = False
DcbIqara.Enabled = False
    
    If ChekClodePeriod(StrDate.value) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
        Else
            MsgBox "Please Change Date Becouse This is Period is Closed"
        End If
        Exit Sub
    End If
    
    Dim Msg As String
 '   If mchkAllowEditPaymentCont Then
 '       TxtModFlg = "E"
 '   End If
    
    If (ChkRenew Or checkContractTransactions(val(TxtContNo.Text))) Then
        mCanEdit = True
        
    Else
        mCanEdit = False
    End If
    mCanEdit = True
    If ChkRenew.value = vbChecked Then
        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
        Exit Sub
    End If

    If checkContractTransactions(val(TxtContNo.Text)) = True Then
        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰  ⁄œÌ·…", vbCritical
        Exit Sub
    
    End If
    
        If TxtNoteSerial.Text <> "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                         MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
                Else
                          MsgBox "Please Delete JE"
                End If
            CuurentLogdata "E"
        Exit Sub
        End If
    If DoPremis(Do_Edit, Me.Name, True) = False Then
      Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtContNo.Text <> "" Then
        TxtModFlg = "E"
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        UnitsGrid.Rows = UnitsGrid.Rows + 1
        VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
        Frm2.Enabled = True
        ReloadUonit
       ' Me.TxtVacName.SetFocus
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select

End Sub
Private Sub btnNew_Click()
    
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    clear_all Me
    chkDivWater.value = vbChecked
    lblremain.Caption = ""
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    DcbUnitNo.Enabled = True
    DcbIqara.Enabled = True
    DcbUnitType.Enabled = True
    '-----------------------------------
    'Me.TxtVac_ID.text = ""
    'Me.TxtVacName.text = ""
    '-----------------------------------
    RdRTypeDate(0).value = True
    TxtModFlg.Text = "N"
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    UnitsGrid.Clear flexClearScrollable, flexClearEverything
    UnitsGrid.Rows = 2
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    Me.Dcbranch.BoundText = Current_branch
    TxtPaymentCount.Text = 2
    TxtPeriods.Text = 6
    DcbPeriodsID.ListIndex = 1
    Opt(0).value = True
    Opt(4).value = True
    ReloadUonit
    'My_SQL = "TblContract"
    'rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    'If rs.RecordCount > 0 Then
        'TxtContNo.text = rs.RecordCount + 1
    'Else
        'TxtContNo.text = 1
    'End If
    'rs.Close
    'CmbType.ListIndex = 0
    'TxtVacName.SetFocus
    ComResid(1).value = True
    ComResid_Click (0)
    RecorddateH.value = ToHijriDate(Date)
    FromDateH.value = ToHijriDate(Date)
    todateH.value = ToHijriDate(Date)
    FirstInstallDateH.value = ToHijriDate(Date)
    ContDate.value = Date
    StrDate.value = Date
    EndDate.value = Date
    FristPaymentDate.value = Date
    Me.LblTotalQasts.Caption = 0
    Opt(2).value = True
    DCboUserName.BoundText = user_id
    ClculteVAT
    RecorddateH.SetFocus
Contract_period.ListIndex = 1
Contract_period_no.Text = 1
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    
    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtContNo.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    
    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtContNo.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnQuery_Click()
    'Load FrmIqarContractSearch
    FrmIqarContractSearch.m_RetrunType = 0
    FrmIqarContractSearch.show vbModal
End Sub
Private Sub btnSave_Click()
   On Error GoTo ErrTrap
    If 1 = 1 Then
        Dim Msg As String
        Dim StrVacCode As String
        Dim StrVacName As String
        Dim CtrlTxt As Control
        checkdates
        
        If ChekClodePeriod(StrDate.value) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
            Else
                MsgBox "Please Change Date Becouse This is Period is Closed"
            End If
            Exit Sub
        End If
    Dim AccountVATDept As String
    Dim account As String
    
   If SystemOptions.MustEnterNewNo = True And TXTNewNO.Text = "" Then
    
                              If SystemOptions.UserInterface = ArabicInterface Then
                                   MsgBox "Ì—ÃÏ «œŒ«· «·—ﬁ„ «·„ÊÕœ ", vbCritical
                            Else
                                  MsgBox "Please  Enter Ministry of housing  No", vbCritical
                            End If
                            Exit Sub
       
   End If
    
    If SystemOptions.OpenVATAccountOwner = True And commisiontype = 1 Then
    
    
    Else
    PercentgValueAddedAccount_Transec StrDate.value, 8, 1, account
   AccountVat.BoundText = account
    If AccountVat.BoundText = "" And True = True And ComResid(1).value = True And CheckAnyVAT(StrDate.value) = True Then
    
    MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
    Exit Sub
    End If
    End If
    Dim Account_Code_dynamic As String
     If (val(TxtCommiValue.Text)) > 0 And SystemOptions.DueComm = True Then
              Account_Code_dynamic = get_account_code_branch(153, my_branch)
              If Account_Code_dynamic = "NO branch" Then
              MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
                Else
               If Account_Code_dynamic = "NO account" Then
                  MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «” Õﬁ«ﬁ «·”⁄Ì", vbCritical
                 Exit Sub
        
               End If
                End If
    
     End If
        Iqar = val(DcbIqara.BoundText)
        commisiontype = AqarCommisionType(Iqar, AmolaValues, ownerid)
               
        '---------------------- check if data Vaclete -----------------------
    
        For Each CtrlTxt In Me.Controls
            If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
                If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                    MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                    CtrlTxt.SetFocus
                    Exit Sub
                End If
            End If
        Next
        
        With VSFlexGrid2
            If val(.Rows) >= 2 Then
                If val(.TextMatrix(1, .ColIndex("id"))) = 0 Then
                    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
                    Exit Sub
                End If
            Else
                MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
                .Rows = .Rows + 1
                Exit Sub
            End If
        End With
        
        If val(dcCustomer.BoundText) = 0 Then
            MsgBox "ÌÃ» «Œ Ì«— «”„ «·„” «Ã—"
            dcCustomer.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
              
        If val(TxtTotalContract.Text) = 0 Then
            MsgBox "ÌÃ»   «œŒ«· ﬁÌ„… «·«ÌÃ«—"
            TxtTotalContract.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
    
        If val(TxtPaymentCount) = 0 Then
            MsgBox "·«»œ „‰  ÕœÌœ «·› —… »Ì‰ «·œ›⁄« "
            TxtPaymentCount.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
    
   Dim s As String
        s = "Select * from TblIqarDiscountTrans2 Where UnitNo = " & val(DcbUnitNo.BoundText) & " and unittype = " & val(DcbUnitType.BoundText)
        s = s & " and Iqar = " & val(DcbIqara.BoundText) '& " and BranchID = " & val(Dcbranch.BoundText)
        Dim rsDummy As New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
          '  txtDiscountPercent.Text = rsDummy!DiscountPercent & ""
            txtDiscountPercent.Tag = rsDummy!DiscountPercent & ""
        End If
        If val(txtDiscountPercent) > val(txtDiscountPercent.Tag) Then
            MsgBox "·« Ì„ﬂ‰  Ã«Ê“ ‰”»… «·Œ’„ «·„Õœœ…"
            txtDiscountPercent.SetFocus
            Exit Sub
        End If
        
        Dim i As Integer
        Dim NpayedValue As Double
        Dim contracttotals As Double
        
        NpayedValue = 0
        contracttotals = 0
        With GridInstallments
            NpayedValue = .Aggregate(flexSTSum, .FixedRows, .ColIndex("NpayedValue"), .Rows - 1, .ColIndex("NpayedValue"))
            'contracttotals = val(TxtTotalContract) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtOldRent) + val(TxtOldWater) - NpayedValue
            contracttotals = val(TxtTotalContract) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtFATValue.Text) - NpayedValue
            For i = .FixedRows To .Rows - 1
                If Opt(4).value = True And i = 1 Then
                    .TextMatrix(i, .ColIndex("RentValue")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + (contracttotals - val(LblTotalQasts.Caption))
                End If
                If Opt(3).value = True And i = (.Rows - 1) Then
                    .TextMatrix(i, .ColIndex("RentValue")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + (contracttotals - val(LblTotalQasts.Caption))
                End If
            Next i
        End With
          
        If checkistallment = False Then
            Exit Sub
        End If
        
        If val(Me.Dcbranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If
            Dcbranch.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    
        If val(DcbIqara.BoundText) = 0 Then
            MsgBox "ÌÃ» «Œ Ì«— «”„ «·⁄ﬁ«—"
            DcbIqara.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        If val(dcsupplier.BoundText) = 0 Then
            MsgBox "ÌÃ» «Œ Ì«— «”„ «·„«·ﬂ"
            dcsupplier.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        If val(DcbUnitType.BoundText) = 0 Then
            MsgBox "ÌÃ» «Œ Ì«—   ‰Ê⁄ «·ÊÕœ…"
            DcbUnitType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If val(DcbUnitNo.BoundText) = 0 Then
            MsgBox "ÌÃ» «Œ Ì«—   —ﬁ„ «·ÊÕœ…"
            DcbUnitNo.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    
        Dim SUM As Double
        SUM = 0
        
        If VSFlexGrid2.Rows > 1 Then
            'fg2.Rows = fg2.Rows - 1
            With VSFlexGrid2
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, .ColIndex("empname")) <> "" Then
                        SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
                    End If
                Next i
                If SUM > 100 Or SUM < 100 Then
                    MsgBox " ·« Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» Ì”«ÊÌ 100%"
                    Exit Sub
                End If
            End With
        End If
    
    
    
        '------------------------------ check if Empcode exist ----------------------
        'StrVacName = IsRecExist("TblContract", "GovernmentName", Trim(TxtVacName.text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")
        'If StrVacName <> "" Then
        'Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        'MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        'TxtVacName.SetFocus
        'Exit Sub
        'End If
    my_branch = val(Me.Dcbranch.BoundText)
        If CheckAcconts = False Then Exit Sub
            If TxtNoteSerial.Text = "" And Opt(0).value = True Then    'ÃœÌœ ›ﬁÿ
            my_branch = val(Me.Dcbranch.BoundText)
                If Notes_coding(val(my_branch), ContDate.value) = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                my_branch = val(Me.Dcbranch.BoundText)
                    If Notes_coding(val(my_branch), ContDate.value) = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                End If
            End If
        End If
    
        If Opt(1).value = True Then
            TxtNoteSerial = ""
        End If
    
        Dim TxtNoteSerial1str As String
        
        my_branch = val(Me.Dcbranch.BoundText)
        If TxtNoteSerial1.Text = "" Then
            TxtNoteSerial1str = Voucher_coding(val(my_branch), ContDate.value, 60, 60)
            If TxtNoteSerial1str = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…         ⁄ﬁœ ÃœÌœ  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                If TxtNoteSerial1str = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  ⁄ﬁœ ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    'TxtNoteSerial1.text = TxtNoteSerial1str
                    'txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                End If
            End If
        End If
    End If
  
       
       If ComResid(1).value = True And commisiontype = 1 And ComResid(1).value = True Then
         PercentgValueAddedAccount_Transec StrDate, 21, 1, account
            AccountVat2.BoundText = account
       If AccountVat2.BoundText = "" Then
       MsgBox "Ì—ÃÏ  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
       Exit Sub
       End If
       End If
       
    Select Case Me.TxtModFlg.Text
        Case "N"
            AddNewRec
            'BtnLast_Click
        Case "E"
            FiLLRec
    End Select
SendMessage (1)
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
Private Sub BtnUndo_Click()
    Me.TxtModFlg.Text = "R"
    FindRec val(TxtContNo.Text)
    FlagContrNew2 = False
End Sub
Private Sub BtnUpdate_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click

    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub BtnUpdate6_Click()
    Dim RsDetails1 As ADODB.Recordset
    Dim StrSQL As String
    If val(Dcbranch.BoundText) <> 0 Then
        StrSQL = "update TblContract set Branch_NO = " & val(Dcbranch.BoundText) & " where ContNo=" & val(TxtContNo.Text) & " "
        Cn.Execute StrSQL, , adExecuteNoRecords
        MsgBox " „ «· ÕœÌÀ"
    End If
    
End Sub
Function checkdates()
        If DateDiff("D", StrDate.value, EndDate.value) < 0 Then
        MsgBox "·« Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ ‰Â«ÌÂ «·⁄ﬁœ ﬁ»· »œ«Ì …"
       Exit Function
        End If
End Function
Private Sub Cmd_Click(Index As Integer)

   On Error Resume Next
    Dim MSGType As Integer
    Select Case Index
        Case 0
            RemoveGridRow2
        Case 11
            If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            ShowAttachments TxtContNo.Text & TxtNoteSerial1.Text, "270120153"
        Case 20
        
checkdates
        
        If TxtNotSreail1 <> "" Then
RtriveInfoOrbon val(TxtNotID.Text)
End If
        If FlagContrNew2 = False Then
        If TxtNoteSerial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If
End If
            If Me.TxtModFlg.Text <> "R" Then
                If Opt(4).value = False And Opt(3).value = False And Opt(2).value = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
                    Else
                        MsgBox "Please Select Method Number of decimal"
                    End If
                    Exit Sub
                End If
                If val(TxtTotalContract.Text) < val(TxtMiniRentValue.Text) Then
                    MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·«Ã«— «ﬁ· „‰ «ﬁ· ﬁÌ„…  «ÃÌ—ÌÂ"
                    TxtTotalContract.SetFocus
                    Exit Sub
                End If
                If val(TxtPaymentCount) = 0 Then
                    MsgBox "·«»œ „‰  ÕœÌœ «·› —… »Ì‰ «·œ›⁄« "
                    TxtPaymentCount.SetFocus
                    'SendKeys "{F4}"
                     Exit Sub
                End If
                If CheckJE() = True Then
                 MSGType = MsgBox("”Ê› Ì „ Õ–› ﬁÌœ «·œ›⁄«  ", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
                 If MSGType = vbNo Then
                 Exit Sub
                 End If
                End If
                DeleteJE
                Calculations
            End If
        Case 13
            RemoveGridRow
    End Select
End Sub
Function CheckJE() As Boolean
Dim i As Integer
CheckJE = False
With GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("NoteId"))) <> 0 Then
CheckJE = True
Exit Function
End If
Next i
End With
End Function
'Function GetMaxInstal() As Double
'Dim Rs8 As ADODB.Recordset
'Set Rs8 = New ADODB.Recordset
'Dim sql As String
'sql = " SELECT     MAX(dbo.TblContractInstallments.InstallNo) AS maxinstal"
'sql = sql & " FROM         dbo.TblContractInstallments RIGHT OUTER JOIN"
'sql = sql & "                       dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
'sql = sql & "  WHERE     (dbo.TblContract.NoteSerial1 = N'" & TxtNoteSerial1.text & "')"
'Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Rs8.RecordCount > 0 Then
'GetMaxInstal = IIf(IsNull(Rs8("maxinstal").value), 0, Rs8("maxinstal").value)
'Else
'GetMaxInstal = 0
'End If
'End Function
Private Sub Calculations(Optional WithMsg As Boolean = True)
                    Dim Percetage2 As Double
'    On Error GoTo ErrTrap
    Dim SngAllValue As Single
    Dim i  As Integer
    Dim IntNoOFQast As Integer
    Dim IntRes As Integer
    Dim SngOnePor As Single
    Dim FirstDate As Date
    Dim PreDate As Date
    Dim NewDate As Date
    Dim DateInterval As String
    Dim NewDateH As String
    Dim PreDateH As String
    Dim InstalNew As Double
    Dim DateNumber As Integer
    Dim Msg As String
    Dim ActulaPyaed As Double
Dim watervalue As Double
Dim Electricity As Double
ActulaPyaed = 0
    If TxtPaymentCount.Text = "" Then
   
            Msg = "ÌÃ» ≈œŒ«· ⁄œœ «·√ﬁ”«ÿ"

                        If WithMsg = True Then
                            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            TxtPaymentCount.SetFocus
                        End If

            Exit Sub
  End If
  
 If chkDivWater.value = vbChecked Then
 If val(TxtPaymentCount.Text) > 0 Then
 watervalue = Round(val(TxtWater.Text) \ val(TxtPaymentCount.Text), 2)
 Else
 watervalue = 0
 End If
 Else
 watervalue = val(TxtWater.Text)
 End If

 If chkDivElectric.value = vbChecked Then
  If val(TxtPaymentCount.Text) > 0 Then
 Electricity = Round(val(TxtElectricity.Text) \ val(TxtPaymentCount.Text), 2)
 Else
 Electricity = 0
 End If
Else
Electricity = val(TxtElectricity.Text)
 End If



    If DcbPeriodsID.ListIndex = -1 Then
   
            Msg = "ÌÃ» ≈œŒ«·   «·› —… »Ì‰ «·«ﬁ”«ÿ"

                        If WithMsg = True Then
                            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            DcbPeriodsID.SetFocus
                        End If

            Exit Sub
  End If
  
        If Not IsNumeric(TxtPaymentCount.Text) Then
            Msg = " ⁄œœ «·√ﬁ”«ÿ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"

                    If WithMsg = True Then
                        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                         TxtPaymentCount.SetFocus
                    End If

            Exit Sub
        End If
    SngAllValue = Round((val(TxtTotalContract)) / val(TxtPaymentCount), 2)
    SngAllValue = SngAllValue + val(watervalue) + val(Electricity) + val(TxtEnternet)
    IntNoOFQast = val(TxtPaymentCount)
    SngOnePor = SngAllValue

   ' If val(Me.TxtPaymentCount.text) > 0 Then
   '  '   IntNoOFQast = SngAllValue \ val(Me.TxtPaymentCount.text)
   '  ' SngOnePor = val(Me.TxtPaymentCount.text)
   '     SngOnePor = SngAllValue / IntNoOFQast
   ' Else
   '     SngOnePor = SngAllValue / IntNoOFQast
   ' End If
'
    If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
    End If

    NewDate = FristPaymentDate.value
    NewDateH = FirstInstallDateH.value
     
     DateNumber = val(TxtPeriods.Text)

    'End If
    
   If FlagContrNew2 = True Then
  InstalNew = InstalNo + 1
   End If
    Dim notpayed As Double
    notpayed = 0
    With Me.GridInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + IntNoOFQast

        For i = 1 To IntNoOFQast

            DoEvents
            If FlagContrNew2 = False Then
            .TextMatrix(i, .ColIndex("InstallNo")) = i
            Else
            .TextMatrix(i, .ColIndex("InstallNo")) = InstalNew
            .TextMatrix(i, .ColIndex("TempInstal")) = i
             InstalNew = InstalNew + 1
            End If
            
            
              .TextMatrix(i, .ColIndex("Countsofall")) = val(TxtPeriods.Text)
           
            
            If i = 1 Then
           ''// ''19 08 2015 NetRent
           .TextMatrix(i, .ColIndex("Rent1")) = Round((val(TxtTotalContract.Text)) / val(TxtPaymentCount.Text), 2)
           .TextMatrix(i, .ColIndex("RentArbon")) = val(TxtRetValue2.Text)
           .TextMatrix(i, .ColIndex("VATArboon")) = val(TxtFATValue2.Text)
           .TextMatrix(i, .ColIndex("NetRent")) = val(.TextMatrix(i, .ColIndex("Rent1")))
           .TextMatrix(i, .ColIndex("VATValue")) = val(TxtFATValue.Text) / val(TxtPaymentCount.Text)
           If ComResid(1).value = True Then 'Œ«÷⁄
            .TextMatrix(i, .ColIndex("VATValue")) = .TextMatrix(i, .ColIndex("VATValue")) + 1
           End If
           
           .TextMatrix(i, .ColIndex("Commissions1")) = val(TxtCommiValue.Text)
           .TextMatrix(i, .ColIndex("CommissionsArbon")) = val(TxtCommValue2.Text)
           .TextMatrix(i, .ColIndex("NetCommissions")) = val(TxtCommiValue.Text) - val(TxtCommValue2.Text)
           .TextMatrix(i, .ColIndex("ServiceArbon")) = val(TxtServce.Text)
           
           .TextMatrix(i, .ColIndex("Insurance1")) = val(TxtInsuranceValue.Text)
           .TextMatrix(i, .ColIndex("InsuranceArbon")) = val(TxtInstrunceValue2.Text)
           .TextMatrix(i, .ColIndex("NetInsurance")) = val(TxtInsuranceValue.Text) - val(TxtInstrunceValue2.Text)
           
           If chkDivWater.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Water1")) = Round((val(TxtWater.Text)) / IntNoOFQast, 2)
             .TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.Text)
           Else
           .TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.Text)
            .TextMatrix(i, .ColIndex("Water1")) = val(TxtWater.Text)
          End If
    .TextMatrix(i, .ColIndex("NetWater")) = val(.TextMatrix(i, .ColIndex("Water1"))) '- val(.TextMatrix(i, .ColIndex("WaterArbon")))
    .TextMatrix(i, .ColIndex("Water")) = val(.TextMatrix(i, .ColIndex("NetWater")))
      If chkDivElectric.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Electric1")) = Round((val(TxtElectricity.Text)) / IntNoOFQast, 2)
             .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.Text)
           Else
            .TextMatrix(i, .ColIndex("Electric1")) = val(TxtElectricity.Text)
             .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.Text)
            
          End If
               .TextMatrix(i, .ColIndex("NetElectric")) = val(.TextMatrix(i, .ColIndex("Electric1"))) ' - val(.TextMatrix(i, .ColIndex("WaterArbon")))
           .TextMatrix(i, .ColIndex("Electric")) = val(.TextMatrix(i, .ColIndex("NetElectric")))
           ''//
            .TextMatrix(i, .ColIndex("TelandNet")) = val(TxtPhone)
            If val(txtDiscountPercent.Text) > 0 Then
            Dim RentDiscount As Double
             RentDiscount = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2) * val(txtDiscountPercent.Text) * 0.01
             .TextMatrix(i, .ColIndex("RentValue")) = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2) - RentDiscount
             Else
             
             .TextMatrix(i, .ColIndex("RentValue")) = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2)
             End If
              .TextMatrix(i, .ColIndex("Commissions")) = val(TxtCommiValue)
             .TextMatrix(i, .ColIndex("Insurance")) = val(TxtInsuranceValue)
             .TextMatrix(i, .ColIndex("Commissions")) = val(TxtCommiValue)
           .TextMatrix(i, .ColIndex("Value")) = Round(SngOnePor, Decimal_Places1) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtPhone.Text) + val(.TextMatrix(i, .ColIndex("VATValue")))
         
         
         
         
         .TextMatrix(i, .ColIndex("hijri")) = hijriorJerojian
If chkDivWater.value = vbChecked Then
    .TextMatrix(i, .ColIndex("Water")) = Round((val(TxtWater) / IntNoOFQast), 2)
 Else
    .TextMatrix(i, .ColIndex("Water")) = val(TxtWater)
 End If

 If chkDivElectric.value = vbChecked Then
 .TextMatrix(i, .ColIndex("Electric")) = Round((val(TxtElectricity) / IntNoOFQast), 2)
 Else
 .TextMatrix(i, .ColIndex("Electric")) = val(TxtElectricity)
 End If
           
      
         
            
            
            Else
    '        .TextMatrix(i, .ColIndex("Value")) = Round(SngOnePor, Decimal_Places1)
            
'            If chkDivWater.value = vbChecked Then
'    .TextMatrix(i, .ColIndex("Water")) = val(TxtWater) / IntNoOFQast
' Else
'    .TextMatrix(i, .ColIndex("Water")) = 0
' End If
       If chkDivWater.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Water1")) = Round((val(TxtWater.Text)) / IntNoOFQast, 2)
             '.TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.text)
           Else
           .TextMatrix(i, .ColIndex("WaterArbon")) = 0
            .TextMatrix(i, .ColIndex("Water1")) = 0
          End If
          .TextMatrix(i, .ColIndex("VATValue")) = val(TxtFATValue.Text) / IntNoOFQast
    .TextMatrix(i, .ColIndex("NetWater")) = val(.TextMatrix(i, .ColIndex("Water1")))
    .TextMatrix(i, .ColIndex("Water")) = val(.TextMatrix(i, .ColIndex("NetWater")))
    
       '      .TextMatrix(i, .ColIndex("RentValue")) = Round((val(TxtTotalContract)) / IntNoOFQast, 2)
       If val(txtDiscountPercent.Text) > 0 Then
            RentDiscount = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2) * val(txtDiscountPercent.Text) * 0.01
             .TextMatrix(i, .ColIndex("RentValue")) = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2) - RentDiscount
             Else
             
             .TextMatrix(i, .ColIndex("RentValue")) = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2)
             End If
             
 If chkDivElectric.value = vbChecked Then
 .TextMatrix(i, .ColIndex("Electric")) = Round(val(TxtElectricity) / IntNoOFQast, 2)
 Else
 .TextMatrix(i, .ColIndex("Electric")) = 0
 End If
 
   If chkDivElectric.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Electric1")) = Round((val(TxtElectricity.Text)) / IntNoOFQast, 2)
            ' .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.text)
           Else
            .TextMatrix(i, .ColIndex("Electric1")) = 0
           '  .TextMatrix(i, .ColIndex("ElectricArbon")) = 0
            
          End If
               .TextMatrix(i, .ColIndex("NetElectric")) = val(.TextMatrix(i, .ColIndex("Electric1"))) ' - val(.TextMatrix(i, .ColIndex("WaterArbon")))
          
          .TextMatrix(i, .ColIndex("Electric")) = val(.TextMatrix(i, .ColIndex("NetElectric")))
            End If
            
          
            
            If i = 1 Then
                NewDate = NewDate
                NewDateH = NewDateH
            
            Else
                PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("Due_Date"))))
                
                If hijriorJerojian = 1 Then 'jorijan
                NewDate = DateAdd(DateInterval, DateNumber, PreDate)
                NewDateH = ToHijriDate(NewDate)
                End If
                
                     PreDateH = (Trim(.TextMatrix(i - 1, .ColIndex("Due_DateH"))))
     Dim mVatPercent2 As Double
If hijriorJerojian = 0 Then 'hijri
                NewDateH = (DateAdd(DateInterval, DateNumber, PreDateH))
NewDate = ToGregorianDate(NewDateH)
End If
                
                
                
            End If
   
   
  
   
   If lblnew.Visible = True Then
 ' currentvalue = .TextMatrix(i, .ColIndex("Value"))
 '  increasrate = currentvalue * val(TxtIncresYearRate) / 100
 '  currentvalue = currentvalue + increasrate
 '    .TextMatrix(i, .ColIndex("Value")) = currentvalue
   End If
   
  
   
            .TextMatrix(i, .ColIndex("Due_Date")) = Format(NewDate, "yyyy/MM/dd")
            .TextMatrix(i, .ColIndex("Due_DateH")) = Format(NewDateH, "yyyy/MM/dd")
                   If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then
           notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
        End If
        
        ActulaPyaed = ActulaPyaed + val(.TextMatrix(i, .ColIndex("Payed")))
        
     
            
            Due_Date = Format(NewDate, "yyyy/M/d")
        Next i
        
Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String

Dim mCountDay1 As Integer
Dim mCountDay2 As Integer
Dim mCountDaysTotal As Integer
Dim mCostDay As Double
Dim mVATValue1 As Double
Dim mVATValue2 As Double
Dim mVatPercent As Double

Dim mCountDay1Com As Integer
Dim mCountDay2Com As Integer
Dim mCountDaysTotalCom As Integer
Dim mCostDayCom As Double
Dim mVATValue1Com As Double
Dim mVATValue2Com As Double
Dim mVatPercentCom As Double



Dim currentvalue  As Double
Dim increasrate  As Double
Dim mRentValue As Double
 txtDateK.value = CDate("2020-05-14")
 txtDateK2.value = CDate("2020-06-30")
        Dim mPecr1 As Double
        Dim mPecr2 As Double
        Dim mStrDate As Date
        For i = 1 To .Rows - 1
            
            
         commisiontype = AqarCommisionType(val(DcbIqara.BoundText), AmolaValues, val(dcsupplier.BoundText))
        Dim commission As Double
        If commisiontype = 1 Then

            commission = val(.TextMatrix(i, .ColIndex("RentValue"))) * AmolaValues / 100
            
            If SystemOptions.CommissionDue = True Then
                
            End If

        
        End If
           
            
            
            If .TextMatrix(i, .ColIndex("Due_Date")) = "" Then Exit Sub
            mStrDate = .TextMatrix(i, .ColIndex("Due_Date"))
            mRentValue = val(.TextMatrix(i, .ColIndex("RentValue")))
            
            
            If DateDiff("d", "30-06-2020", mStrDate) <= 0 Then
            mVatPercent = 5
            End If
            
            
            If DateDiff("d", "30-06-2020", mStrDate) <= 0 Then
            mVatPercent = 5
            End If
            
            If ContDate.value <= txtDateK.value Then
            mVatPercent = 5
            End If
            mVatPercent2 = 0
            'If ContDate.Value  > txtDateK.value And ContDate.Value  <= txtDateK2.value And mStrDate <= txtDateK2.value Then
            '    mVatPercent = 5
            ' End If
            '
            'If ContDate.Value  > txtDateK.value And ContDate.Value  <= txtDateK2.value And mStrDate <= txtDateK2.value Then
            '    mVatPercent = 5
            '    mVatPercent2 = 0
            ' ElseIf ContDate.Value  > txtDateK.value And ContDate.Value  <= txtDateK2.value And mStrDate > txtDateK2.value Then
            '    mVatPercent = 5
            '    mVatPercent2 = 15
            '
            ' End If
            

            
            mVATValue2 = 0
            mCountDay1 = 0
            mCountDay2 = 0
            mCountDaysTotal = 0
            mCostDay = 0
            mVATValue1 = 0
            mVATValue2 = 0
            mVATValue1Com = 0
            mVATValue2Com = 0
                
                
                If i = .Rows - 1 Then
                    newinstallNo = val(.TextMatrix(i, .ColIndex("InstallNo")))
                    nextinstalldate = EndDate.value
                Else
                    newinstallNo = val(.TextMatrix(i + 1, .ColIndex("InstallNo")))
                    nextinstalldate = .TextMatrix(i + 1, .ColIndex("Due_Date"))
                End If
                
              '  getnextDate newinstallNo, nextinstalldate, nextinstalldateH
'            Dim mVATValue1Com As Double
'Dim mVATValue2Com
                
            
            If year(nextinstalldate) < 1900 Then
            nextinstalldate = Time
            End If
            mCountDaysTotal = DateDiff("D", mStrDate, nextinstalldate) '+ 1
            If mCountDaysTotal = 0 Then mCountDaysTotal = 1
            mCostDay = val(mRentValue) / mCountDaysTotal
            mCostDayCom = val(commission) / mCountDaysTotal
            mVATValue2 = 0
            
           ' If (SQLDate(ContDate.value, False)) > SQLDate(txtDateK.value, False) And (SQLDate(ContDate.value, False)) <= SQLDate(txtDateK2.value, False) Then
         
            
            PercentgValueAddedAccount_Transec StrDate.value, 51, 1, , mVatPercentCom
            
            If DateDiff("d", CDate(mStrDate), txtDateK2.value) < 0 And DateDiff("d", txtDateK2.value, nextinstalldate) >= 0 Then
                mCountDay1 = mCountDaysTotal
                mVatPercent = 15
                mVatPercent2 = 0
                mVATValue1 = Round(val(val(mRentValue) * mVatPercent / 100), 4)
                mVATValue1Com = Round(val(val(commission) * mVatPercentCom / 100), 4) '
            ElseIf DateDiff("d", CDate(mStrDate), txtDateK2.value) >= 0 And DateDiff("d", txtDateK2.value, nextinstalldate) < 0 Then
                mCountDay1 = mCountDaysTotal
                mVatPercent = 5
                mVatPercent2 = 0
                mVATValue1 = Round(val(val(mRentValue) * mVatPercent / 100), 4)
                mVATValue1Com = Round(val(val(commission) * mVatPercentCom / 100), 4)
            ElseIf DateDiff("d", mStrDate, txtDateK2.value) >= 0 And DateDiff("d", txtDateK2.value, nextinstalldate) > 0 Then
                mVatPercent = 5
                mVatPercent2 = 15
                mCountDay1 = DateDiff("D", mStrDate, txtDateK2.value) '+ 1
                mCountDay2 = mCountDaysTotal - mCountDay1
                
                mVATValue1 = Round(val(mCostDay * mCountDay1 * mVatPercent / 100), 2)
                mVATValue2 = Round(val(mCostDay * mCountDay2 * mVatPercent2 / 100), 2)
                
                mVATValue1Com = Round(val(mCostDayCom * mCountDay1 * mVatPercentCom / 100), 2)
                mVATValue2Com = Round(val(mCostDayCom * mCountDay2 * mVatPercentCom / 100), 2)
                
            End If
            
            mCountDay2 = (mCountDaysTotal - mCountDay1)
            
            
            
            
            
            
'
'            ElseIf ContDate.value <= txtDateK.value Then
'                mVatPercent = 5
'                mVatPercent2 = 0
'                mVATValue1 = Round(val(mRentValue) * mVatPercent / 100, 2)
'                mVATValue1Com = Round(val(commission) * mVatPercent / 100, 2)
'                mCountDay1 = mCountDaysTotal
'            ElseIf ContDate.value > txtDateK2.value Then
'            mVatPercent = 15
'            mVatPercent2 = 0
'            mVATValue1 = Round(val(mRentValue) * mVatPercent / 100, 2)
'            mVATValue1Com = Round(val(mRentValue) * mVatPercent / 100, 2)
'            mCountDay1 = mCountDaysTotal
            'End If
            
            
            .TextMatrix(i, .ColIndex("CountDay1")) = mCountDay1
            .TextMatrix(i, .ColIndex("CountDay2")) = mCountDay2
            
            .TextMatrix(i, .ColIndex("VATYou1")) = mVatPercent
            .TextMatrix(i, .ColIndex("VATYou2")) = mVatPercent2
             
             If mPecr1 = 0 Then
                If mVatPercent <> 0 Then mPecr1 = mVatPercent
             End If
            
             If mPecr2 = 0 Then
                If mVatPercent2 <> 0 Then mPecr2 = mVatPercent2
             End If
            
             
     '    If CommiValueInVAT.value = vbChecked Then
                   
     '             PercentgValueAddedAccount_Transec StrDate.value, 21, 1, , Percetage2
     '
     '                            mVATValue1 = mVATValue1 + (val(TxtCommiValue) * Percetage2) / 100
     '
     '             End If
           
            
             
            
            If ComResid(1).value = True Then
            .TextMatrix(i, .ColIndex("VATValue1")) = mVATValue1
            .TextMatrix(i, .ColIndex("VATValue2")) = mVATValue2
            .TextMatrix(i, .ColIndex("VATValue")) = mVATValue1 + mVATValue2
            If i = 1 Then
             
            
            .TextMatrix(i, .ColIndex("VATValue1")) = mVATValue1 + Subvat
            .TextMatrix(i, .ColIndex("VATValue")) = .TextMatrix(i, .ColIndex("VATValue")) + Subvat
            End If
                
                
                
                
                
                .TextMatrix(i, .ColIndex("VATValue1Com")) = mVATValue1Com
                .TextMatrix(i, .ColIndex("VATValue2Com")) = mVATValue2Com
                
                
                .TextMatrix(i, .ColIndex("VATValueCom")) = mVATValue1Com + mVATValue2Com

           
                          
            Else
            If i = 1 Then
                .TextMatrix(i, .ColIndex("VATValue1")) = Subvat
                .TextMatrix(i, .ColIndex("VATValue")) = Subvat
                
             Else
             .TextMatrix(i, .ColIndex("VATValue1")) = 0
                .TextMatrix(i, .ColIndex("VATValue2")) = 0
                
                
                .TextMatrix(i, .ColIndex("VATValue")) = 0
                   
                 .TextMatrix(i, .ColIndex("VATValueCom")) = 0
                
                
                
                End If
            
            End If
            
            

        
           ' .TextMatrix(i, .ColIndex("DiffAmount")) = Round(mVATValue1 + mVATValue2 - val(.TextMatrix(.Rows - 1, .ColIndex("VATValueOld"))), 2)
            .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) '- commission

        Next
LblNotPayed.Caption = notpayed
LblActulaPyaed.Caption = ActulaPyaed
lblremain = val(LblTotalQasts) - val(LblActulaPyaed)
         .AutoSize 1, .Cols - 1, False
        Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        
        Dim mTotalVat1 As Double
        Dim mTotalVat2 As Double
        If ComResid(1).value = True Then
            TxtFATYou = mPecr1
            TxtFATYou22 = mPecr2
            mTotalVat1 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VATValue1"), .Rows - 1, .ColIndex("VATValue1"))
            mTotalVat2 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VATValue2"), .Rows - 1, .ColIndex("VATValue2"))
            TxtFATValue = mTotalVat1 + mTotalVat2
        Else
            TxtFATYou = ""
            TxtFATYou22 = ""
            mTotalVat1 = 0
            mTotalVat2 = 0
            TxtFATValue = ""
            
        End If
        End With
        Calculte True
 ReLineGrid

 
    'BolQastCal = True
    Exit Sub
ErrTrap:
End Sub
Private Sub CmdCus_Click()
If Me.TxtModFlg.Text <> "R" Then
RsCustomers.Index = 1
Load RsCustomers
RsCustomers.show
End If
End Sub

Private Sub cmdDisplayOldPayment_Click()
 'Dim Frm As New FrmOldInstallments
 'Frm.ContNo = Me.TxtContNo
 'Frm.show vbModal
End Sub

Private Sub CmDRENEW_Click()
 Dim Msg As String
Dim temp As Integer
Dim i As Integer
Dim increasevalue As Double
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

'  On Error GoTo ErrTrap

If ChkRenew.value = vbChecked Then

MsgBox "·« Ì„ﬂ‰  ÃœÌœ Â–« «·⁄ﬁœ ·«‰… „Ãœœ »«·›⁄·"
Exit Sub
End If

InstalNo = 0
    If TxtContNo.Text <> "" Then
    With GridInstallments
    For i = 1 To .Rows - 1
   ' MsgBox val(.TextMatrix(i, .ColIndex("InstallNo")))
    If InstalNo < val(.TextMatrix(i, .ColIndex("InstallNo"))) Then
    DTPicker2.value = CDate(.TextMatrix(i, .ColIndex("Due_Date")))
    InstalNo = val(.TextMatrix(1, .ColIndex("InstallNo")))
    End If
    Next i
    End With
   ' ClculteVAT
   
        TxtContNoOld.Text = val(TxtContNo.Text)
        FromdateHO.value = FromDateH.value
         FromdateO.value = StrDate.value
        
    FrmContractOldData.Visible = True
    TxtNotSreail1.Text = ""
TxtNotID.Text = ""
TxtNotVal.Text = ""
   ' ChkRenew.value = vbChecked
    lblnew.Visible = True
    lblnew.Caption = "Ã«—Ì «· ÃœÌœ"
        TxtModFlg = "N"
    VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        Frm2.Enabled = True
   TxtContNo.Text = ""
   'TxtNoteSerial1.text = ""
      'TxtNoteSerial.text = ""
      DcbIqara_Click (0)
increasevalue = val(TxtTotalContract) * val(TxtIncresYearRate.Text) / 100
TxtTotalContract = TxtTotalContract + increasevalue
Dim noOfMonth As Integer
increasevalue = val(TxtPhone) * val(TxtIncresYearRate.Text) / 100
TxtPhone = val(TxtPhone) + val(increasevalue)
'TxtInsuranceValue.text = 0
TxtRetValue2.Text = 0
TxtFATValue2.Text = 0
TxtInstrunceValue2.Text = 0
TxtWaterValue2.Text = 0
TxtCommValue2.Text = 0
  TxtCommValue2.Text = 0
     TxtServce.Text = 0
     TxtOldRent.Text = 0
     TxtElectricityValue2.Text = 0
     txtOldInsurance.Text = val(TxtInsuranceValue.Text)
     TxtInsuranceValue.Text = 0
     TxtPhone.Text = 0
     FlagContrNew = True
     FlagContrNew2 = True
'increasevalue = val(TxtCommiValue) * val(TxtIncresYearRate.text) / 100       'val(TxtCommiValue.text)
'TxtCommiValue = TxtCommiValue + increasevalue
TxtCommiValue.Text = 0
                   VBA.Calendar = vbCalHijri
                    FromDateH.value = todateH.value ' DateAdd("YYYY", 1, TodateH.value)
                    todateH.value = DateAdd("YYYY", 1, FromDateH.value)
                    
   VBA.Calendar = vbCalGreg
Fromdateh_LostFocus
ToDateH_LostFocus


Cmd_Click (20)
If TxtNoteSerial1.Text = "" Or val(TxtNoteSerial1.Text) = 0 Then
TxtNoteSerial1 = Voucher_coding(val(Dcbranch.BoundText), DTPicker2.value, 60, 60)
  End If
       ' Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select
End Sub

Private Sub cmdSavePayment_Click()
  Dim Msg As String
'    If mchkAllowEditPaymentCont Then
'        TxtModFlg = "E"
'    End If
    
    If (ChkRenew Or checkContractTransactions(val(TxtContNo.Text))) And mchkAllowEditPaymentCont Then
        mCanEdit = True
        
    Else
        mCanEdit = False
    End If
    
    If ChkRenew.value = vbChecked And Not mchkAllowEditPaymentCont Then
        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
        Exit Sub
    End If


    If checkContractTransactions(val(TxtContNo.Text)) = True And Not mchkAllowEditPaymentCont Then
        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰  ⁄œÌ·…", vbCritical
        Exit Sub
    
    End If
    
            Dim s As String
        Dim RsDetails2 As New ADODB.Recordset
        s = "Select * from TblContractInstallmentsHist Where ContNo = " & Trim(TxtContNo.Text)
        
 
    
        RsDetails2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        'If Not RsDetails2.EOF Then
            RsDetails2.AddNew
            RsDetails2!UserID = user_id
            RsDetails2!EditDate = Now
            RsDetails2!ContNo = val(TxtContNo)
            RsDetails2.update
        'End If
       
       SaveGridPayment False
       SaveGridPayment True
       MsgBox " „ Õ›Ÿ  ⁄œÌ·«  «·œ›⁄« "
       RetriveOldPayment
       
    
End Sub

Private Sub Command1_Click()
    'RSContractInstallments.show
End Sub

Private Sub Command10_Click()
    If checkApility("FrmCashing1") = False Then
                Exit Sub
            End If
FrmCashing1.show
FrmCashing1.newrecord
FrmCashing1.DCboCashType.ListIndex = 8
 FrmCashing1.TxtContNo.Text = val(TxtContNo.Text)
  FrmCashing1.TxtContractNo.Text = (TxtNoteSerial1.Text)
  
  
          '  OpenScreen CashingDataScreen
End Sub

Private Sub Command11_Click()
Dim StrSQL As String
Dim des As String

If checkallocation2(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «Ì—«œ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If

If checkAllocations(val(TxtContNo), des) = True Then
MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «” Õﬁ«ﬁ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
Exit Sub
End If





       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
       Cn.Execute "Update TblContract set NoteID=null,NoteSerial=null where ContNo=" & val(TxtContNo.Text) & " "
       
TxtNoteSerial.Text = ""
TxtNoteID.Text = 0
MsgBox " „ Õ–› «·ﬁÌœ"
RsSavRec.Resync adAffectCurrent
End Sub

Private Sub Command12_Click()
    
    
   If DoPremis(Do_Edit, Me.Name, True) = False Then
      Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtContNo.Text <> "" Then
        TxtModFlg = "E"
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        UnitsGrid.Rows = UnitsGrid.Rows + 1
        VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
        Frm2.Enabled = True
        ReloadUonit
       ' Me.TxtVacName.SetFocus
    End If
 '   Exit Sub

    
  '  If Not mCreateEntryManual Then
  '      MsgBox "«·ﬁÌœ Ì‰‘√ ¬·Ì« „⁄ «·Õ›Ÿ"
  '      Exit Sub
  '  End If
    If ChekClodePeriod(StrDate.value) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
        Else
            MsgBox "Please Change Date Becouse This is Period is Closed"
        End If
        Exit Sub
    End If
    
    
    

If TxtNoteSerial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If

   
    Dim StrSQL As String
    
    StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where ContNo=" & val(Me.TxtContNo.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    If CheckAcconts = False Then Exit Sub
    createVoucher
     TxtModFlg = "R"



    SendMessage 1
    Exit Sub
ErrTrap:
Dim Msg As String
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select

    
End Sub

Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0)
End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim Opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", Opt)
  If Opt = currentOpt Then
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '«·„” √Ã—
 Msg = "‘ﬂ—« ·«Œ Ì«—ﬂ„ " & CompanyName & "  ··„·«ÕŸ«    " & companyphone
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(val(dcCustomer.BoundText)))



DoEvents
 Msg = "  „ ⁄„· ⁄ﬁœ —ﬁ„ " & TxtNoteSerial1 & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.Text & "    ··⁄ﬁ«— —ﬁ„ " & DcbIqara.Text
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(val(dcsupplier.BoundText)))


 

DoEvents



MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function

Private Sub Command13_Click()
CreateSearial
End Sub

Private Sub Command14_Click()
 

 Dim StrSQL As String
 If ChKLegalIssue.value = vbChecked Then
 
  StrSQL = "    update  dbo.TblContract  set LegalIssue =1  where ContNo=" & val(TxtContNo.Text)
  Else
  StrSQL = "    update  dbo.TblContract  set LegalIssue =null  where ContNo=" & val(TxtContNo.Text)
  End If
  
  Cn.Execute StrSQL
  
  
   If ChkAccredit.value = vbChecked Then
 
  StrSQL = "    update  dbo.TblContract set  NewNO='" & (TXTNewNO) & "' , Accredit =1  where ContNo=" & val(TxtContNo.Text)
  Else
  StrSQL = "    update  dbo.TblContract  set  NewNO='" & (TXTNewNO) & "' ,Accredit =null  where ContNo=" & val(TxtContNo.Text)
  End If
  
  Cn.Execute StrSQL
  
  
  RsSavRec.Resync adAffectCurrent
  MsgBox " „ «· ÕœÌÀ"


End Sub

Private Sub Command15_Click()
  Dim StrSQL As String
   If ChkAccredit.value = vbChecked Then
 
  StrSQL = "    update  dbo.TblContract  set Accredit =1  where CusID=" & val(Me.dcCustomer.BoundText)
  StrSQL = StrSQL & "and UnitNo =" & val(Me.DcbUnitNo.BoundText)
  Else
  StrSQL = "    update  dbo.TblContract  set Accredit =null  where CusID=" & val(Me.dcCustomer.BoundText)
  StrSQL = StrSQL & "and UnitNo =" & val(Me.DcbUnitNo.BoundText)
  End If
  
  
   
  Cn.Execute StrSQL
    RsSavRec.Resync adAffectCurrent
  MsgBox " „ «· ÊÀÌﬁ ·ﬂ· ⁄ﬁÊœ «·„” √Ã—"


  
End Sub

Private Sub Command2_Click()
If Me.TxtModFlg = "R" Then
    If TxtContNo <> "" Then
                print_report TxtContNo
                SendMessage (2)
            End If
End If
End Sub

Private Sub Command3_Click()
    'RSContractInstallments.show
End Sub

Private Sub Command5_Click()
If Me.TxtModFlg.Text = "R" And val(TxtContNo.Text) <> 0 Then
   If checkApility("FrmWaiverSettlement") = False Then
                Exit Sub
            End If
           Load FrmWaiverSettlement
             FrmWaiverSettlement.show
FrmWaiverSettlement.Cmd_Click (0)

   
   FrmWaiverSettlement.DcbIqara2.BoundText = val(DcbIqara.BoundText)
FrmWaiverSettlement.DcbUnitType2.BoundText = val(DcbUnitType.BoundText)
FrmWaiverSettlement.DcbUnitNo2.BoundText = val(DcbUnitNo.BoundText)
FrmWaiverSettlement.dcCustomer2.BoundText = val(dcCustomer.BoundText)
'
 If FrmWaiverSettlement.chek(TxtContNo.Text) = False Then
' FrmWaiverSettlement.TxtContNo.Text = TxtContNo.Text

'FrmWaiverSettlement.TxtOrder = TxtNoteSerial1
' FrmWaiverSettlement.GetContract val(TxtNoteSerial1)


    FrmWaiverSettlement.DcbIqara2.BoundText = val(DcbIqara.BoundText)
FrmWaiverSettlement.DcbUnitType2.BoundText = val(DcbUnitType.BoundText)
FrmWaiverSettlement.DcbUnitNo2.BoundText = val(DcbUnitNo.BoundText)
FrmWaiverSettlement.dcCustomer2.BoundText = val(dcCustomer.BoundText)

  FrmWaiverSettlement.DcbIqara.BoundText = val(DcbIqara.BoundText)
  FrmWaiverSettlement.DcbUnitType.BoundText = val(DcbUnitType.BoundText)
  FrmWaiverSettlement.DcbUnitNo.BoundText = val(DcbUnitNo.BoundText)
  FrmWaiverSettlement.dcCustomer.BoundText = val(dcCustomer.BoundText)

FrmWaiverSettlement.TxtOrder = TxtNoteSerial1
FrmWaiverSettlement.TxtContNo.Text = TxtContNo.Text

 
 ' FrmWaiverSettlement.GetContract TxtNoteSerial1.Text
  'FrmWaiverSettlement.TxtContNo.Text = val(TxtContNo.Text)
    
   ' FrmWaiverSettlement.RetriveOrder val(TxtContNo.Text)
    
   
   
' FrmWaiverSettlement.RetriveOrder val(TxtContNo.Text)
 
 End If
 End If

End Sub
Sub DleteUnit()
Dim StrSQL As String
Dim i As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
          Cn.Execute "  update TblAqarDetai  Set Status = 0   ,customerid=null  Where id =" & val(DcbUnitNo.BoundText)
            StrSQL = " SELECT     dbo.TblIqrMerg.UntID"
            StrSQL = StrSQL & "          FROM         dbo.TblIqrMerg INNER JOIN"
            StrSQL = StrSQL & "          dbo.TblContract ON dbo.TblIqrMerg.Cont = dbo.TblContract.ContNo"
            StrSQL = StrSQL & " Where (dbo.TblIqrMerg.cont = " & val(TxtContNo.Text) & ") And (dbo.TblContract.CusID =" & val(dcCustomer.BoundText) & ")"
            Rs7.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs7.RecordCount > 0 Then
            Rs7.MoveFirst
            For i = 1 To Rs7.RecordCount
             Cn.Execute "  update TblAqarDetai  Set Status = 0   ,customerid=null  Where id =" & IIf(IsNull(Rs7("UntID").value), 0, Rs7("UntID").value)
             Rs7.MoveNext
             Next i
             End If
           
End Sub
Private Sub Command6_Click()
If Me.TxtModFlg.Text = "R" Then
Unload FrmSanadatOFContract
Load FrmSanadatOFContract
FrmSanadatOFContract.Indx = 0
FrmSanadatOFContract.Label1(0).Caption = TxtNoteSerial1.Text
FrmSanadatOFContract.TxtNotID.Text = val(TxtNotID.Text)
FrmSanadatOFContract.TxtContNo.Text = val(TxtContNo.Text)
FrmSanadatOFContract.show
End If
End Sub

Private Sub Command7_Click()
If Me.TxtModFlg.Text = "R" Then

Unload FrmWaiver
Load FrmWaiver
FrmWaiver.show
FrmWaiver.Cmd_Click (0)
FrmWaiver.TxtContNo.Text = val(TxtContNo.Text)
End If
End Sub

Private Sub Command8_Click()
Dim StrTempAccountCode As String
                   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
 
            ShowReport StrTempAccountCode, dcCustomer.Text, FrmDate.value, ToDate.value

End Sub

Private Sub Command9_Click()
       ShowGL_cc Me.TxtNoteSerial.Text, , 200
       
End Sub

Private Sub CommiValueInVAT_Click()
Calculte
End Sub

Private Sub txtDiscountPercent_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtDiscountValue.Text = val(TxtTotalContract) * val(txtDiscountPercent) * 0.01
    Calculte
End If
End Sub

Private Sub WaterElecValueInVAT_Click()
Calculte
End Sub
Private Sub InsurValueInVAT_Click()
Calculte
End Sub
Private Sub ComResid_Click(Index As Integer)
ClculteVAT
End Sub
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim account2 As String
Dim account As String
Dim Percetage2 As Double
If ComResid(1).value = True Then
PercentgValueAddedAccount_Transec StrDate.value, 8, 1, account, Percetage
commisiontype = AqarCommisionType(val(DcbIqara.BoundText), , val(dcsupplier.BoundText))
PercentgValueAddedAccount_Transec StrDate.value, 21, 1, account2, Percetage2
AccountVat2.BoundText = account2
TxtFATYou2.Text = Percetage2
If SystemOptions.OpenVATAccountOwner = True And commisiontype = 1 Then
TxtFATYou.Text = 0
AccountVat.BoundText = ""
Else
TxtFATYou.Text = Percetage
AccountVat.BoundText = account
End If
Else
TxtFATYou.Text = 0
AccountVat.BoundText = ""
End If
Calculte
End If
End Sub
Private Sub ContDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     RecorddateH.value = ToHijriDate(ContDate.value)
         datetype
    If ChekSanNumber(Current_branch, 60) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial.Text = ""
End If
End Sub

Private Sub ContDate_GotFocus()
hijriorJerojian = 1
End Sub

Private Sub Contract_period_Change()
CalcContractIntervalAuto
End Sub

Private Sub Contract_period_Click()
CalcContractIntervalAuto
End Sub

Private Sub Contract_period_no_Click()
CalcContractIntervalAuto
End Sub

Private Sub DcbIqara_Change()
DcbUnitType_Change
DcbIqara_Click (0)
Calculte
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then dcsupplier.BoundText = 0: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
    dcsupplier.BoundText = ownerid
    Calculte
    'DcbUnitType_Change
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 1
FrmAqarSearch.show


End If


If KeyCode = vbKeyF5 Then
ReloadCombos
End If

End Sub

Private Sub DcboEmp_Change()
 'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If
DcboEmp_Click (0)
End Sub


Private Sub DcboEmp_Click(Area As Integer)
Dim i As Integer
If val(Me.DcboEmp.BoundText) <> 0 Then

With VSFlexGrid2
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("id")) = Me.DcboEmp.BoundText Then
Exit Sub
End If
Next i
If .Rows = 2 Then
.TextMatrix(.Rows - 1, .ColIndex("rate")) = 100
End If
If .Rows <> 1 Then
.TextMatrix(.Rows - 1, .ColIndex("id")) = Me.DcboEmp.BoundText
.TextMatrix(.Rows - 1, .ColIndex("empname")) = Me.DcboEmp.Text
End If
.Rows = .Rows + 1

End With
End If
End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub DcboEmpName_Change()
   If val(DcboEmpName.BoundText) = 0 Then TxtEmpCode.Text = "":  Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtEmpCode.Text = EmpCode
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 DcboEmpName_Change
    
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)



    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 20
      Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
        
  
    End If
    
    
End Sub
Private Sub Dcbranch_Click(Area As Integer)
    If ChekSanNumber(Current_branch, 60) = True Then
        TxtNoteSerial1.Text = ""
    End If
    TxtNoteSerial.Text = ""
End Sub

Private Sub DcbRentType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub DcbUnitNo_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Integer
   Set Dcombos = New ClsDataCombos

If val(DcbUnitType.BoundText) > 0 Then
idd = val(DcbUnitNo.BoundText)
Dim meterPrice As Double
Dim lengh As Double
Dim customerid As Integer
Dim rentType As Integer
Dim ElectAccount As String
Dim MiniRentValue As Double
Dim Typed As Integer
 Me.TxtRemarks = GetIqarUnitData(idd, , meterPrice, lengh, customerid, rentType, , , , , , ElectAccount, MiniRentValue, Typed)
 TxtElectAccount.Text = ElectAccount
 DcbRentType.ListIndex = IIf(rentType < 0, 0, rentType - 1)
 TxtMeterValue.Text = meterPrice
 TxtMeterCount.Text = lengh
 TxtMiniRentValue.Text = MiniRentValue
 If Typed = 1 Then
 ComResid(1).value = True
 Else
 ComResid(0).value = True
 End If
 ReLineGrid
 ' dcCustomer.BoundText = customerid
  
End If

If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    Dim s As String
    s = "Select * from TblIqarDiscountTrans2 Where UnitNo = " & val(DcbUnitNo.BoundText) & " and unittype = " & val(DcbUnitType.BoundText)
    s = s & " and Iqar = " & val(DcbIqara.BoundText) '& " and BranchID = " & val(Dcbranch.BoundText)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        txtDiscountPercent.Text = rsDummy!DiscountPercent & ""
        txtDiscountPercent.Tag = rsDummy!DiscountPercent & ""
    End If
End If
End Sub

Private Sub DcbUnitNo_Click(Area As Integer)
DcbUnitNo_Change
End Sub

Private Sub DcbUnitNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub DcbUnitType_Change()
ReloadUonit
End Sub
Sub ReloadUonit(Optional flg As Integer = 0)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
Dim StrSQL As String
Set Dcombos = New ClsDataCombos
     StrSQL = " or id in(Select UntID from  TblIqrMerg where cont =" & val(TxtContNo.Text) & ")"
     StrSQL = StrSQL & " or id in (Select UnitNo from  TblContract    Where ContNo =" & val(TxtContNo.Text) & ")"
If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)
idd1 = val(DcbUnitType.BoundText)
If Me.TxtModFlg = "R" Or flg = 1 Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
ElseIf Me.TxtModFlg = "N" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
ElseIf Me.TxtModFlg = "E" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "E", StrSQL
End If
End If
End Sub


Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub DcbUnitType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub dcCustomer_Change()
  If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
     GetTblCustemersCode , , dcCustomer.BoundText, EmpCode, 56
    Me.Text15.Text = EmpCode

End Sub

Private Sub dcCustomer_Click(Area As Integer)
 dcCustomer_Change
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 1215
        FrmCustemerSearch.show vbModal

    End If
 

If KeyCode = vbKeyF5 Then
ReloadCombos
End If

End Sub

Private Sub dcCustomer_LostFocus()
    If ChecStopeCustomer(val(dcCustomer.BoundText)) = True Then
    MsgBox " Â–« «·„” «Ã— ›Ì «·ﬁ«∆„… «·”Êœ«¡ ·«Ì„ﬂ‰ «· «⁄„· „⁄Â"
    dcCustomer.BoundText = 0
    Text15.Text = ""
    End If
End Sub

Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
   If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text1.Text = EmpCode
    ClculteVAT
End Sub

Private Sub ENDDATE_Change()
'If Me.TxtModFlg.Text <> "R" Then
         todateH.value = ToHijriDate(EndDate.value)
       hijriorJerojian = 1
'End If
End Sub

Private Sub FirstInstallDateH_GotFocus()

hijriorJerojian = 0
End Sub
Private Sub FirstInstallDateH_LostFocus()
        
        If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            FristPaymentDate.value = ToGregorianDate(FirstInstallDateH.value)
               
        End If

End Sub
 
Private Sub Form_Load()

    On Error GoTo ErrTrap
ScreenNameArabic = "‘«‘Â ⁄ﬁÊœ «·«ÌÃ«—"
ScreenNameEnglish = " Real Estate Contract    "

    RereivID = 0
    
    Dim i As Integer
    Dim My_SQL As String
    'wael
    If SystemOptions.CanAcreditRsContract = True Then
    ChkAccredit.Enabled = True
    Command14.Enabled = True
    Else

    ChkAccredit.Enabled = False
     Command14.Enabled = False
    End If
'
    
    '   If SystemOptions.SpecialVersion = True Then
'Ele(6).Visible = False
'   End If
   
    lblnew.Visible = False
    'My_SQL = "TblContract"
    If SystemOptions.TypeContractAutoFromIqar = True Then
       ComResid(0).Enabled = False
       ComResid(1).Enabled = False
    Else
       ComResid(0).Enabled = True
       ComResid(1).Enabled = True
    End If
    
  If SystemOptions.NoCreatJLInRentContract = True Then
  Command12.Enabled = False
  Else
  Command12.Enabled = True
  End If
    Dim RsOpt As ADODB.Recordset
    Set RsOpt = New ADODB.Recordset
    RsOpt.Open "select IsNull(CreateEntryManual,0) as  CreateEntryManual ,isNull(chkAllowEditPaymentCont,0) as chkAllowEditPaymentCont from TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not RsOpt.EOF Then
        mCreateEntryManual = RsOpt!CreateEntryManual
        mchkAllowEditPaymentCont = True ' RsOpt!chkAllowEditPaymentCont
    End If
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
   
    If SystemOptions.UserInterface = ArabicInterface Then
        UnitsGrid.ColComboList(UnitsGrid.ColIndex("namerentType")) = "#1; «·ﬁÌ„… «·«ÌÃ«—Ì…|#2; »«·„ —"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        UnitsGrid.ColComboList(UnitsGrid.ColIndex("namerentType")) = "#1;Rental value |#2;meter "
    End If
    'RsSavRec.CursorLocation = adUseClient
    'RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
       
    My_SQL = " select * from TblContract where 1=-1"
    'If SystemOptions.usertype = UserAdminAll Then
    'Else
    'My_SQL = My_SQL & " where   Branch_NO=" & Current_branch
    'End If
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    Set cSearch = New clsDCboSearch
    'Set cSearch.Client = Me.DcboGovernmentID
    'ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("GovernmentID"), Me.DcboGovernmentID
   ' FillGridWithData
    If ChekSanNumber(Current_branch, 60) = True Then
        TxtNoteSerial1.locked = True
    Else
        TxtNoteSerial1.locked = False
    End If
    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("CityName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next i
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip
    ReloadCombos
    'loadcombo
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub
Public Function ReloadCombos()

    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
  
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
 
    Set Dcombos = New ClsDataCombos
    Dcombos.GetAccountingCodes AccountVat
    Dcombos.GetAccountingCodes AccountVat2
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    'Dcombos.GetIqarUnit 1, DcbUnitNo
    Dcombos.GetBranches Dcbranch
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetUsers Me.DCboUserName
End Function
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Streets Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name"
    Label1(1).Caption = "Neighborhood"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("CityID")) = "Id"
        .TextMatrix(0, .ColIndex("CityName")) = "Name"
        .TextMatrix(0, .ColIndex("GovernmentID")) = "Neighborhood"
    End With

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
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub
Sub Calculte(Optional ByVal IsFromInstall As Boolean = True)
If Me.TxtModFlg.Text <> "R" Then
Dim TotalService  As Double
Dim TotalValueALL As Double
Dim AddTovatValue As Double
Dim Percetage2 As Double
 TotalService = val(TxtWater.Text) + val(TxtElectricity.Text) + val(TxtPhone.Text) 'water,electric,net
 
TotalValueALL = val(TxtTotalContract.Text) - val(TxtDiscountValue.Text) 'rent
TotalValueALL = TotalValueALL + TotalService ' water,electric,net
TotalValueALL = TotalValueALL + val(TxtCommiValue.Text)
TotalValueALL = TotalValueALL + val(TxtInsuranceValue.Text)

AddTovatValue = TotalValueALL
PercentgValueAddedAccount_Transec StrDate.value, 21, 1, , Percetage2

If ComResid(1).value = True Then  'Õ«·Â Œ«÷€
                If WaterElecValueInVAT.value = vbUnchecked Then
                    AddTovatValue = AddTovatValue - TotalService
                End If
                If CommiValueInVAT.value = vbUnchecked Then
                  AddTovatValue = AddTovatValue - val(TxtCommiValue.Text) '
                  End If
                If InsurValueInVAT.value = vbUnchecked Then
                AddTovatValue = AddTovatValue - val(TxtInsuranceValue.Text)
                  End If
Else

 

  End If
 

Subvat = 0
PercentgValueAddedAccount_Transec StrDate.value, 21, 1, , Percetage2
'  AddTovatValue = AddTovatValue - val(TxtTotalContract.Text) - val(TxtDiscountValue.Text)
'AddTovatValue = AddTovatValue - val(TxtTotalContract.Text) - val(TxtDiscountValue.Text)
           If WaterElecValueInVAT.value = Checked Then
                    Subvat = Subvat + val(TotalService) * Percetage2 / 100
                End If
                If CommiValueInVAT.value = vbChecked Then
                Subvat = Subvat + val(TxtCommiValue) * Percetage2 / 100
                   
                  End If
                If InsurValueInVAT.value = vbChecked Then
                Subvat = Subvat + val(TxtInsuranceValue) * Percetage2 / 100
                 
                   End If
                   
    TxtNetValue.Text = val(AddTovatValue)
    'salim here 15 02 2021     If Not IsFromInstall Then
    
   
        If ComResid(1).value = True And val(TxtFATYou.Text) > 0 Then '«·÷—Ì»Â Õ”«» ‰”»…
             TxtFATValue.Text = (val(AddTovatValue) * val(TxtFATYou.Text)) / 100
        Else
 
      
                    TxtFATValue.Text = Subvat
        End If
 

     
 TxtTotalValue.Text = AddTovatValue + val(TxtFATValue.Text) '
    'If CommiValueInVAT.value = vbChecked Then
    '    TxtTotalValue.Text = val(TxtNetValue.Text) + val(TxtFATValue.Text)
    'Else
    '    TxtTotalValue.Text = val(TxtNetValue.Text) + val(TxtFATValue.Text) + val(TxtCommiValue.Text)
    'End If
    
  '  If InsurValueInVAT.value = vbChecked Then
  '      TxtTotalValue.Text = val(TxtNetValue.Text) + val(TxtFATValue.Text)
  '  Else
  '      TxtTotalValue.Text = val(TxtTotalValue.Text) + val(TxtInsuranceValue.Text)
  '
  '  End If
  '
End If
End Sub
Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

    Set cSearch = Nothing
ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Sub SaveInstalPreValue()
Dim StrSQL As String
Dim SumPreValue As Double
Dim RsDetails1 As ADODB.Recordset
Set RsDetails1 = New ADODB.Recordset
SumPreValue = val(TxtOldRent.Text) + val(TxtOldWater.Text) + val(TxtOldElectric.Text) + val(TxtoldCommi.Text)
       If SumPreValue <> 0 Then
      StrSQL = "SELECT     *  from dbo.TblContractInstallments Where (1 = -1)"
     RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     RsDetails1.AddNew
     RsDetails1("ContNo").value = val(TxtContNo.Text)
     RsDetails1("OldValue").value = SumPreValue
     RsDetails1("DES").value = balanceDes.Text
     RsDetails1("InstallNo").value = 0
     RsDetails1("Installdate").value = balanceDate.value
     RsDetails1("InstalldateH").value = balanceDateH.value
     RsDetails1("installValue").value = SumPreValue
     RsDetails1.update
     End If
End Sub
Sub GetSuperVisorOrbion(Optional NoteID As Double = 0)
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
sql = " SELECT     dbo.TblNotesSales.NoteID, dbo.TblNotesSales.ID, dbo.TblNotesSales.rate, dbo.TblNotesSales.valu, dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID,"
sql = sql & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblNotesSales.idd, dbo.TblNotesSales.GroupID,"
sql = sql & "                      dbo.TBLSalesRepGroups.name , dbo.TBLSalesRepGroups.NameE"
sql = sql & " FROM         dbo.TblNotesSales LEFT OUTER JOIN"
sql = sql & "                      dbo.TBLSalesRepGroups ON dbo.TblNotesSales.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblNotesSales.Type = 1) And (dbo.TblNotesSales.NoteID = " & NoteID & ")"

Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   With Me.VSFlexGrid2
       .Rows = 1
        .Clear flexClearScrollable
If Rs7.RecordCount > 0 Then

        If Rs7.RecordCount > 0 Then
           .Rows = Rs7.RecordCount + 1
           Rs7.MoveFirst

            For i = 1 To .Rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
   If SystemOptions.UserInterface = EnglishInterface Then
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(Rs7.Fields("Emp_Namee").value), "", Rs7.Fields("Emp_Namee").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs7.Fields("namee").value), "", Rs7.Fields("namee").value)
      Else
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(Rs7.Fields("Emp_Name").value), "", Rs7.Fields("Emp_Name").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs7.Fields("name").value), "", Rs7.Fields("name").value)
 
    End If
    .TextMatrix(i, .ColIndex("groupid")) = val(IIf(IsNull(Rs7.Fields("GroupID").value), "", Rs7.Fields("GroupID").value))
    .TextMatrix(i, .ColIndex("rate")) = val(IIf(IsNull(Rs7.Fields("rate").value), "", Rs7.Fields("rate").value))
    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(Rs7.Fields("Fullcode").value), "", Rs7.Fields("Fullcode").value)
    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs7.Fields("EmpID").value), "", Rs7.Fields("EmpID").value)
   .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(Rs7.Fields("idd").value), "", Rs7.Fields("idd").value)
        Rs7.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End If
    
    End With


End Sub
Function RtriveInfoOrbon(Optional NotID As Double = 0) As Boolean
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
Dim total As Double
RtriveInfoOrbon = True
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT   *  from  Notes where NoteID =" & NotID & ""
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
   total = (IIf(IsNull(RsDetails1("allowdate").value), Date, RsDetails1("allowdate").value))
   If allowdate.value < ContDate.value And SystemOptions.AllowOrbonDate = False Then
   RtriveInfoOrbon = False
   Exit Function
   End If
   TxtNotSreail1.Text = val(IIf(IsNull(RsDetails1("NoteSerial1").value), "", RsDetails1("NoteSerial1").value))
 TxtNotVal.Text = val(IIf(IsNull(RsDetails1("Note_Value2").value), TxtNotVal.Text, RsDetails1("Note_Value2").value))
 total = val(IIf(IsNull(RsDetails1("Note_Value2").value), "", RsDetails1("Note_Value2").value))
 DcbIqara.BoundText = val(IIf(IsNull(RsDetails1("akarid").value), "", RsDetails1("akarid").value))
 DcbUnitType.BoundText = val(IIf(IsNull(RsDetails1("UnitType").value), "", RsDetails1("UnitType").value))
 DcbUnitNo.BoundText = val(IIf(IsNull(RsDetails1("UnitNo").value), "", RsDetails1("UnitNo").value))
 TxtTotalContract.Text = val(IIf(IsNull(RsDetails1("rent").value) Or RsDetails1("rent").value = 0, TxtTotalContract.Text, RsDetails1("rent").value))
 TxtCommiValue.Text = val(IIf(IsNull(RsDetails1("commission").value) Or RsDetails1("commission").value = 0, TxtCommiValue.Text, RsDetails1("commission").value))
 TxtInsuranceValue.Text = val(IIf(IsNull(RsDetails1("Instrunce").value) Or RsDetails1("Instrunce").value = 0, TxtInsuranceValue.Text, RsDetails1("Instrunce").value))
 TxtWater.Text = val(IIf(IsNull(RsDetails1("Water").value) Or RsDetails1("Water").value = 0, TxtWater.Text, RsDetails1("Water").value))
 TxtElectricity.Text = val(IIf(IsNull(RsDetails1("Electricity").value) Or RsDetails1("Electricity").value = 0, TxtElectricity.Text, RsDetails1("Electricity").value))
 TxtPhone.Text = val(IIf(IsNull(RsDetails1("Servce").value) Or RsDetails1("Servce").value = 0, TxtPhone.Text, RsDetails1("Servce").value))
 TxtFATValue2.Text = val(IIf(IsNull(RsDetails1("VAT").value), 0, RsDetails1("VAT").value))

If val(TxtCommiValue.Text) <= total Then

Me.TxtCommValue2.Text = val(Me.TxtCommiValue.Text)

Else
Me.TxtCommValue2.Text = total
End If
total = total - val(TxtCommValue2.Text)
'''//////////
If val(TxtPhone.Text) <= total Then
Me.TxtServce.Text = Me.TxtPhone.Text
Else
Me.TxtServce.Text = total
End If
total = total - val(TxtServce.Text)

''////////
If val(TxtInsuranceValue.Text) <= total Then
Me.TxtInstrunceValue2.Text = Me.TxtInsuranceValue.Text
ElseIf total > 0 Then
Me.TxtInstrunceValue2.Text = total
Else
Me.TxtInstrunceValue2.Text = 0
End If
total = total - val(TxtInstrunceValue2.Text)
''//
If val(TxtWater.Text) <= total Then
If chkDivWater.value = vbChecked Then
Me.TxtWaterValue2.Text = Round(val(Me.TxtWater.Text) / val(TxtPaymentCount.Text), 2)
Else
Me.TxtWaterValue2.Text = Me.TxtWater.Text
End If
ElseIf total > 0 Then
Me.TxtWaterValue2.Text = total
Else
Me.TxtWaterValue2.Text = 0
End If
total = total - val(TxtWaterValue2.Text)
''//
''//
If val(TxtElectricity.Text) <= total Then
If chkDivElectric.value = vbChecked Then
Me.TxtElectricityValue2.Text = Round(val(Me.TxtElectricity.Text) / val(TxtPaymentCount.Text), 2)
Else
Me.TxtElectricityValue2.Text = Me.TxtElectricity.Text
End If
ElseIf total > 0 Then
Me.TxtElectricityValue2.Text = total
Else
Me.TxtElectricityValue2.Text = 0
End If
''//
total = total - val(TxtElectricityValue2.Text)
If val(TxtTotalContract.Text) <= total Then
Me.TxtRetValue2.Text = Me.TxtTotalContract.Text
ElseIf total > 0 Then
Me.TxtRetValue2.Text = total
Else
Me.TxtRetValue2.Text = 0
End If


   End If
End Function

Private Sub ReLineGrid(Optional ByVal FormRet As Boolean = False)
    Dim i As Integer
    Dim IntCounter As Integer
   ''''///
   If FlagContrNew2 = False Then
   Dim SUM As Double
   Dim RentValue As Double
   If val(TxtMeterValue) <> 0 Then
   RentValue = (val(TxtMeterValue) * val(TxtMeterCount))
   Else
   RentValue = val(TxtTotalContract.Text)
   End If
   Else
   RentValue = val(TxtTotalContract.Text)
   End If
   
   IntCounter = 0
   SUM = 0
     With UnitsGrid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("nameunittype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                  If .TextMatrix(i, .ColIndex("length")) <> "" Then
                  If .TextMatrix(i, .ColIndex("meterPrice")) <> "" Then
                  .TextMatrix(i, .ColIndex("RentValue")) = val(.TextMatrix(i, .ColIndex("meterPrice"))) * val(.TextMatrix(i, .ColIndex("length")))
                  RentValue = RentValue + val(.TextMatrix(i, .ColIndex("RentValue")))
                  End If
                  End If
            End If

        Next i
        
    End With
    'TxtTotalContract.Text = val(TxtTotalContract.Text) + (val(TxtMeterValue) * val(TxtMeterCount))
    If Not FormRet Then
        TxtTotalContract = RentValue
    End If
     IntCounter = 0
    With VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("empname")) <> "" Then
                IntCounter = IntCounter + 1
                
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
                If SUM > 100 Then
                .TextMatrix(i, .ColIndex("rate")) = 0
                MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» «ﬂ»— „‰ 100%"
                Exit Sub
                End If
            End If

        Next i

    End With
    
  With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Des")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If

        Next i

    End With
    
    Dim notpayed As Double
    Dim OldValue As Double
    notpayed = 0

       ' oldvalue = val(TxtOldRent) + val(TxtOldWater) + val(TxtOldElectric) + val(TxtoldCommi)
  With Me.GridInstallments

        For i = .FixedRows To .Rows - 1
    '  .TextMatrix(1, .ColIndex("NpayedValue")) = val(TxtNotVal.text)

'.TextMatrix(i, .ColIndex("OldValue")) = 0
            If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked Then
           ' If i = 1 Then
           '   .TextMatrix(i, .ColIndex("OldValue")) = oldvalue
           '  End If
              .TextMatrix(i, .ColIndex("OldValueDate")) = Format(balanceDate.value, "yyyy/MM/dd")
  .TextMatrix(i, .ColIndex("OldValueDateH")) = Format(balanceDateH.value, "yyyy/MM/dd")
   .TextMatrix(i, .ColIndex("DES")) = balanceDes.Text
   '.TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("OldValue"))) - val(.TextMatrix(i, .ColIndex("NpayedValue")))
   .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) - val(.TextMatrix(i, .ColIndex("NpayedValue")))
                 
   If val(.TextMatrix(i, .ColIndex("Payed"))) = 0 Then
    .TextMatrix(i, .ColIndex("Payed")) = val(.TextMatrix(i, .ColIndex("VATArboon"))) + val(.TextMatrix(i, .ColIndex("RentArbon"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon"))) + val(.TextMatrix(i, .ColIndex("WaterArbon"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon")))
    .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
   End If
               If .Rows > 0 Then
  Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
lblOldValue.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OldValue"), .Rows - 1, .ColIndex("OldValue"))
Else
Me.LblTotalQasts.Caption = 0
End If
             ' Exit Sub
           
                Else
               ' .TextMatrix(i, .ColIndex("OldValue")) = 0
              '                .TextMatrix(i, .ColIndex("OldValueDate")) = ""
'.TextMatrix(i, .ColIndex("OldValueDateH")) = ""
'.TextMatrix(i, .ColIndex("DES")) = ""

                '.TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("OldValue"))) - val(.TextMatrix(i, .ColIndex("NpayedValue")))
                .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) - val(.TextMatrix(i, .ColIndex("NpayedValue")))
            End If
        Next i
  .AutoSize 1, .Cols - 1, False
    End With
    
         With Me.GridInstallments

        For i = .FixedRows To .Rows - 1
 
            If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked Then
 
                    notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
                Else
 
            End If
'.TextMatrix(i, .ColIndex("OldValue")) = 0
'.TextMatrix(i, .ColIndex("OldValueDate")) = ""
'.TextMatrix(i, .ColIndex("OldValueDateH")) = ""
'.TextMatrix(i, .ColIndex("DES")) = ""

        Next i

    End With
    
    LblNotPayed = notpayed
    
             With Me.GridInstallments2

        For i = .FixedRows To .Rows - 1
 
            If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked Then
 
                    notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
                Else
 
            End If
'.TextMatrix(i, .ColIndex("OldValue")) = 0
'.TextMatrix(i, .ColIndex("OldValueDate")) = ""
'.TextMatrix(i, .ColIndex("OldValueDateH")) = ""
'.TextMatrix(i, .ColIndex("DES")) = ""

        Next i

    End With
    LblNotPayed2 = notpayed
    End Sub
Function CheckAcconts() As Boolean
CheckAcconts = False
            Account_Code_dynamic80 = get_account_code_branch(80, my_branch)
            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
            
            Account_Code_dynamic92 = get_account_code_branch(92, my_branch)
            Account_Code_dynamic59 = get_account_code_branch(59, my_branch)
            Account_Code_dynamic123 = get_account_code_branch(123, my_branch)
            Account_Code_dynamic125 = get_account_code_branch(125, my_branch)

Account_Code_dynamic154 = get_account_code_branch(154, my_branch)
Account_Code_dynamic155 = get_account_code_branch(155, my_branch)
Account_Code_dynamic156 = get_account_code_branch(156, my_branch)


            If commisiontype = 1 Then
            If AmolaValues > 0 Then
                     If Account_Code_dynamic125 = "NO account" Then
                                                    If SystemOptions.UserInterface = ArabicInterface Then
                                                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»     ⁄„Ê·«  „” Õﬁ… „‰ «„·«ﬂ «·€Ì—  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                    Else
                                                        MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                    End If
                    
                                        GoTo ErrTrap
                     End If
            End If
                               If SystemOptions.Create2account4Supp = False Then
                                     If Account_Code_dynamic123 = "NO account" Then
                                                        If SystemOptions.UserInterface = ArabicInterface Then
                                                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «ÌÃ«—«  „” Õﬁ… ··€Ì—  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                        Else
                                                            MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                        End If
                        
                                            GoTo ErrTrap
                                 End If
                          
                          End If
            End If
            
            If (val(TxtOldRent) + val(TxtOldWater) + val(TxtOldElectric) + val(TxtoldCommi)) > 0 Then
            
              If Account_Code_dynamic59 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»       Ê”Ìÿ √ ··⁄„·«¡   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
            End If
            
            If Account_Code_dynamic80 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»      «·«ÌÃ«—«  «·„” Õﬁ… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
If Opt(0).value = True Or Opt(1).value = True Then ' ÃœÌœ
                If (val(TxtPayAmini) + val(TxtCommiValue)) > 0 Then
                            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
                            If Account_Code_dynamic81 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»         «·”⁄Ì Ê «·—”Ê„ «·«œ«—Ì… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
              
              
               If (val(TxtInsuranceValue)) > 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «· √„Ì‰ «·„” —œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
                 
                     If (val(txtOldInsurance)) > 0 Then
                            Account_Code_dynamic92 = get_account_code_branch(92, my_branch)
                            If Account_Code_dynamic92 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  Ê”Ìÿ √›  «ÕÌ  ·· √„Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
                 
                 
              
              
                    If (val(TxtWater)) > 0 Then
                            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
                            If Account_Code_dynamic83 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·„Ì«Â «·„ﬁœ„… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
              
              
               If (val(TxtElectricity)) > 0 Then
                            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
                            If Account_Code_dynamic84 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·ﬂÂ—»«¡ «·„ﬁœ„… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
                 
              
                      If (val(TxtPhone) + val(TxtEnternet)) > 0 Then
                            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
                            If Account_Code_dynamic85 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·Œœ„«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                                            Else
                                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
                              
                 End If
              
                
End If



   CheckAcconts = True
   Exit Function
ErrTrap:
      CheckAcconts = False
End Function
Public Sub AddNewRec()
'    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblContract", "ContNo", "")
    RsSavRec.AddNew
TxtContNo.Text = StrRecID
    RsSavRec.Fields("ContNo").value = IIf(StrRecID <> "", StrRecID, Null)

If lblnew.Visible = False And TxtNoteSerial1.Text = "" Then
TxtNoteSerial1 = Voucher_coding(val(my_branch), ContDate.value, 60, 60)
  End If
  
  
  RsSavRec.Fields("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.Text), Null)

RsSavRec.update
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()

    'On Error GoTo ErrTrap
    
    Dim RsDetails1 As ADODB.Recordset
    Dim StrMerg As String
    Dim i As Integer
    Dim StrSQL As String
    Dim TransBegine As Boolean
    StrMerg = ""
    lblnew.Visible = False


    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
        TransBegine = False

    If Me.TxtModFlg.Text = "E" Then
         
            StrSQL = "Delete From TblUnitNoInformation Where ContNo =" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                
            StrSQL = "Delete From TblIqrMerg Where Cont=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                
            StrSQL = "Delete From TblContractDet Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            StrSQL = "Delete From TblCOntractSales Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                
            StrSQL = "Delete From tblContractInsAllocationsDetails1 Where ContractFlag=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                
            'StrSQL = "delete From Notes where NoteID=" & val(Me.TxtNoteID.text) ' Val(rs("Transaction_ID").value)
            'Cn.Execute StrSQL, , adExecuteNoRecords
    
    
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblContractInstallments Where ContNo=" & val(Me.TxtContNo.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 

     End If

    With UnitsGrid
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("nameunittype")) <> "" Then
                StrMerg = StrMerg & .TextMatrix(i, .ColIndex("nameunittype"))
                StrMerg = StrMerg & " " & "—ﬁ„"
                StrMerg = StrMerg & .TextMatrix(i, .ColIndex("unitno"))
                StrMerg = StrMerg & " "
                StrMerg = StrMerg & CHR(13)
            End If
        Next i
    End With
          
    'If TxtNoteSerial1.Text = "" And Opt(0).value = True Then
    'TxtNoteSerial1 = Voucher_coding(val(my_branch), ContDate.value, 60, 60)
    'End If
    If 1 = 1 Then
        If TxtNoteSerial1.Text = "" Then
            TxtNoteSerial1 = Voucher_coding(val(my_branch), ContDate.value, 60, 60)
        End If
    
        RsSavRec.Fields("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", Trim(TxtNoteSerial1.Text), Null)
        RsSavRec.update
        RsSavRec("AccountCodeVat").value = Me.AccountVat.BoundText
        RsSavRec("AccountCodeVat2").value = Me.AccountVat2.BoundText
        RsSavRec("RecorddateH").value = RecorddateH.value
        RsSavRec("FromdateH").value = Me.FromDateH.value
        RsSavRec("FromdateO").value = Me.FromdateO.value
        RsSavRec("FromdateHO").value = Me.FromdateHO.value
        RsSavRec("Remark2").value = Me.TxtRemark2.Text
        RsSavRec("TodateH").value = Me.todateH.value
        If RdRTypeDate(1).value = True Then
        RsSavRec("TypeDate").value = 1
        Else
        RsSavRec("TypeDate").value = 0
        End If
        
        
        RsSavRec.Fields("UserID").value = IIf(DCboUserName.Text <> "", Trim(DCboUserName.BoundText), user_id)
        RsSavRec("FirstInstallDateH").value = Me.FirstInstallDateH.value
        RsSavRec("Branch_NO").value = IIf(val(Dcbranch.BoundText) = 0, Null, Dcbranch.BoundText)
        
        If CommiValueInVAT.value = vbChecked Then
            RsSavRec("CommiValueInVAT").value = 1
        Else
            RsSavRec("CommiValueInVAT").value = 0
        End If
        
                
        If chkIsNotCreateEntry.value = vbChecked Then
            RsSavRec("IsNotCreateEntry").value = 1
        Else
            RsSavRec("IsNotCreateEntry").value = 0
        End If
        
        
        
            If WaterElecValueInVAT.value = vbChecked Then
        RsSavRec("WaterElecValueInVAT").value = 1
        Else
        RsSavRec("WaterElecValueInVAT").value = 0
        End If
        
        If InsurValueInVAT.value = vbChecked Then
            RsSavRec("InsurValueInVAT").value = 1
        Else
            RsSavRec("InsurValueInVAT").value = 0
        End If
        
        If Opt(0).value = True Then
            RsSavRec.Fields("NewOrOpeneing").value = 0
        Else
            RsSavRec.Fields("NewOrOpeneing").value = 1
        End If
        If ChkRenew.value = vbChecked Then
            RsSavRec.Fields("Renew").value = 1
        Else
            RsSavRec.Fields("Renew").value = 0
        End If
        If Opt(4).value = True Then
            RsSavRec("MethodDeci").value = 0
        ElseIf Opt(3).value = True Then
            RsSavRec("MethodDeci").value = 1
        ElseIf Opt(2).value = True Then
            RsSavRec("MethodDeci").value = 2
        End If
        If FlagContrNew2 = True Then
            RsSavRec("FlagContrNew2").value = 1
        Else
            RsSavRec("FlagContrNew2").value = 0
        End If
        'ChKEndContract
        RsSavRec("LegalIssue").value = IIf(ChKLegalIssue.value = vbUnchecked, Null, 1)
        RsSavRec("Employeecontract").value = IIf(ChkEmployeecontract.value = vbUnchecked, Null, 1)
        RsSavRec("Accredit").value = IIf(ChkAccredit.value = vbUnchecked, Null, 1)
        
        
        
        RsSavRec("NewNO").value = (TXTNewNO)
        
        RsSavRec("OutContract").value = IIf(ChKOutContract.value = vbUnchecked, Null, 1)
        RsSavRec("EndContract").value = IIf(ChKEndContract.value = vbUnchecked, Null, 1)
        RsSavRec("DivWater").value = IIf(chkDivWater.value = vbUnchecked, Null, 1)
        RsSavRec("DivElectric").value = IIf(chkDivElectric.value = vbUnchecked, Null, 1)
        RsSavRec("DiscountPercent").value = val(txtDiscountPercent)
        RsSavRec("DiscountvaLUE").value = val(TxtDiscountValue)
        
        RsSavRec("Emp_IDContract").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
        RsSavRec.Fields("UnitElectric").value = IIf(Me.TxtElectAccount.Text <> "", val(TxtElectAccount.Text), Null)
        RsSavRec("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        RsSavRec.Fields("ContDate").value = IIf(ContDate.value <> "", Trim(ContDate.value), Null)
        RsSavRec.Fields("ContType").value = IIf(Me.DcbContType.ListIndex <> -1, val(DcbContType.ListIndex), Null)
        RsSavRec.Fields("Iqar").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
        RsSavRec.Fields("ownerid").value = IIf(val(Me.dcsupplier.BoundText) <> 0, val(dcsupplier.BoundText), Null)
        RsSavRec.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
        'RsSavRec.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
        RsSavRec.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
        RsSavRec.Fields("RentType").value = IIf(Me.DcbRentType.ListIndex <> -1, val(DcbRentType.ListIndex), Null)
        'RsSavRec.Fields("RentType").value = IIf(Me.DcbRentType.ListIndex <> -1, val(DcbRentType.ListIndex), Null)
        RsSavRec.Fields("StrDate").value = IIf(StrDate.value <> "", (StrDate.value), Null)
        RsSavRec.Fields("EndDate").value = IIf(EndDate.value <> "", (EndDate.value), Null)
        RsSavRec.Fields("MeterValue").value = IIf(Me.TxtMeterValue.Text <> "", val(TxtMeterValue.Text), Null)
        RsSavRec.Fields("MeterCount").value = IIf(Me.TxtMeterCount.Text <> "", val(TxtMeterCount.Text), Null)
        RsSavRec.Fields("TotalContract").value = IIf(Me.TxtTotalContract.Text <> "", val(TxtTotalContract.Text), Null)
        RsSavRec.Fields("PayAmini").value = IIf(Me.TxtPayAmini.Text <> "", val(TxtPayAmini.Text), Null)
        RsSavRec.Fields("CommiValue").value = IIf(Me.TxtCommiValue.Text <> "", val(TxtCommiValue.Text), Null)
        RsSavRec.Fields("InsuranceValue").value = IIf(Me.TxtInsuranceValue.Text <> "", val(TxtInsuranceValue.Text), Null)
        RsSavRec.Fields("MiniRentValue").value = IIf(Me.TxtMiniRentValue.Text <> "", val(TxtMiniRentValue.Text), Null)
        RsSavRec.Fields("NotID").value = IIf(Me.TxtNotID.Text <> "", val(TxtNotID.Text), Null)
        RsSavRec.Fields("NotValue").value = IIf(Me.TxtNotVal.Text <> "", val(TxtNotVal.Text), Null)
        RsSavRec.Fields("NoteSrial1").value = IIf(Me.TxtNotSreail1.Text <> "", TxtNotSreail1.Text, Null)
        RsSavRec.Fields("OutOffice").value = IIf(Me.TxtOutOffice.Text <> "", val(TxtOutOffice.Text), Null)
        RsSavRec.Fields("Water").value = IIf(Me.TxtWater.Text <> "", val(TxtWater.Text), Null)
        RsSavRec.Fields("Electricity").value = IIf(Me.TxtElectricity.Text <> "", val(TxtElectricity.Text), Null)
        RsSavRec.Fields("Phone").value = IIf(Me.TxtPhone.Text <> "", val(TxtPhone.Text), Null)
        RsSavRec.Fields("Enternet").value = IIf(Me.TxtEnternet.Text <> "", val(TxtEnternet.Text), Null)
        RsSavRec.Fields("FristPaymentDate").value = IIf(FristPaymentDate.value <> "", (FristPaymentDate.value), Null)
        RsSavRec.Fields("IncresYearValue").value = IIf(Me.TxtIncresYearValue.Text <> "", val(TxtIncresYearValue.Text), Null)
        RsSavRec.Fields("IncresYearRate").value = IIf(Me.TxtIncresYearRate.Text <> "", val(TxtIncresYearRate.Text), Null)
        RsSavRec.Fields("PaymentCount").value = IIf(Me.TxtPaymentCount.Text <> "", val(TxtPaymentCount.Text), Null)
        RsSavRec.Fields("Periods").value = IIf(Me.TxtPeriods.Text <> "", Trim(TxtPeriods.Text), Null)
        RsSavRec.Fields("PeriodsID").value = IIf(Me.DcbPeriodsID.ListIndex <> -1, val(DcbPeriodsID.ListIndex), Null)
        RsSavRec.Fields("CusID").value = IIf(val(Me.dcCustomer.BoundText) <> 0, val(dcCustomer.BoundText), Null)
        RsSavRec.Fields("Furnishing").value = IIf(Me.DcbFurnishing.ListIndex <> -1, val(DcbFurnishing.ListIndex), Null)
        RsSavRec.Fields("Remarks").value = IIf(Me.TxtRemarks.Text <> "", Trim(TxtRemarks.Text), Null)
        RsSavRec.Fields("OthersRules").value = IIf(Me.TxtOthersRules.Text <> "", (TxtOthersRules.Text), Null)
        'RsSavRec.Fields("NoteID").value = IIf(Me.TxtNoteID <> "", Trim(TxtNoteID.text), Null)
        'RsSavRec.Fields("NoteSerial").value = IIf(Me.TxtNoteSerial <> "", Trim(TxtNoteSerial.text), Null)
        RsSavRec.Fields("ContNoOld").value = IIf(Me.TxtContNoOld.Text <> "", val(TxtContNoOld.Text), Null)
        RsSavRec.Fields("RetValue2").value = IIf(Me.TxtRetValue2.Text <> "", val(TxtRetValue2.Text), Null)
        RsSavRec.Fields("FATValue2").value = IIf(Me.TxtFATValue2.Text <> "", val(TxtFATValue2.Text), Null)
        RsSavRec.Fields("WaterValue2").value = IIf(Me.TxtWaterValue2.Text <> "", val(TxtWaterValue2.Text), Null)
        RsSavRec.Fields("CommValue2").value = IIf(Me.TxtCommValue2.Text <> "", val(TxtCommValue2.Text), Null)
        RsSavRec.Fields("InstrunceValue2").value = IIf(Me.TxtInstrunceValue2.Text <> "", val(TxtInstrunceValue2.Text), Null)
        RsSavRec.Fields("StrMerg").value = IIf(StrMerg <> "", StrMerg, Null)
        RsSavRec.Fields("ElectricityValue2").value = IIf(Me.TxtElectricityValue2.Text <> "", val(TxtElectricityValue2.Text), Null)
        RsSavRec.Fields("Servce").value = IIf(Me.TxtServce.Text <> "", val(TxtServce.Text), Null)
        RsSavRec.Fields("OldRent").value = IIf(Me.TxtOldRent.Text <> "", val(TxtOldRent.Text), Null)
        RsSavRec.Fields("OldWater").value = IIf(Me.TxtOldWater.Text <> "", val(TxtOldWater.Text), Null)
        RsSavRec.Fields("OldElectric").value = IIf(Me.TxtOldElectric.Text <> "", val(TxtOldElectric.Text), Null)
        RsSavRec.Fields("oldCommi").value = IIf(Me.TxtoldCommi.Text <> "", val(TxtoldCommi.Text), Null)
        RsSavRec.Fields("OldInsurance").value = IIf(Me.txtOldInsurance.Text <> "", val(txtOldInsurance.Text), Null)
        RsSavRec.Fields("balanceDate").value = IIf(balanceDate.value <> "", (balanceDate.value), Null)
        RsSavRec("balanceDateH").value = balanceDateH.value
        RsSavRec.Fields("balanceDes").value = IIf(Me.balanceDes.Text <> "", (balanceDes.Text), Null)
        If val(TxtContNoOld.Text) <> 0 Then
            RsSavRec("ContNoOld").value = IIf(TxtContNoOld.Text = "", Null, val(TxtContNoOld.Text))
            'RsSavRec("Renew").value = 1
        Else
            RsSavRec("ContNoOld").value = Null
            'RsSavRec("Renew").value = 0
        End If
        If TxtContNoOld.Text <> "" Then
            Cn.Execute "  update TblContract  Set Renew = 1" & "    Where ContNo =" & val(TxtContNoOld.Text)
        End If
        RsSavRec.Fields("NetValue").value = IIf(Me.TxtNetValue.Text <> "", val(TxtNetValue.Text), Null)
        RsSavRec.Fields("FATYou").value = IIf(Me.TxtFATYou.Text <> "", val(TxtFATYou.Text), Null)
        RsSavRec.Fields("FATYou22").value = IIf(Me.TxtFATYou22.Text <> "", val(TxtFATYou22.Text), Null)
        RsSavRec.Fields("FATValue").value = IIf(Me.TxtFATValue.Text <> "", val(TxtFATValue.Text), Null)
        RsSavRec.Fields("TotalValue").value = IIf(Me.TxtTotalValue.Text <> "", val(TxtTotalValue.Text), Null)
        RsSavRec.Fields("FATYou2").value = IIf(Me.TxtFATYou2.Text <> "", val(TxtFATYou2.Text), Null)
        If ComResid(1).value = True Then
        RsSavRec.Fields("ComResid").value = 1
        Else
        RsSavRec.Fields("ComResid").value = 0
        End If
        '*********************
        RsSavRec("Contract_period_no").value = val(Contract_period_no.Text)
        RsSavRec("Contract_period").value = Contract_period.ListIndex

        '**********************
        RsSavRec.update
        RsSavRec.Resync
    
        
        Set RsDetails1 = New ADODB.Recordset
             StrSQL = "SELECT     *  from dbo.TblContractDet Where (1 = -1)"
       RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If VSFlexGrid1.Rows > 1 Then
                    ' fg2.Rows = fg2.Rows - 1
         With VSFlexGrid1
           For i = .FixedRows To .Rows - 1
           
                  If .TextMatrix(i, .ColIndex("Des")) <> "" Then
           
               
           
               RsDetails1.AddNew
              RsDetails1("ContNo").value = val(TxtContNo.Text)
      
               RsDetails1("Des").value = .TextMatrix(i, .ColIndex("Des"))
             '  RsDetails1("Elevatortype").value = .TextMatrix(i, .ColIndex("Elevatortype"))
              RsDetails1("Count").value = val(.TextMatrix(i, .ColIndex("Count")))
                 RsDetails1("Code").value = val(.TextMatrix(i, .ColIndex("Code")))
             '  RsDetails1("MainCo").value = .TextMatrix(i, .ColIndex("MainCo"))
             ' RsDetails1("MaintStrDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("MaintStrDate"))), .TextMatrix(i, .ColIndex("MaintStrDate")), Date)
             ' RsDetails1("MaintEndDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("MaintEndDate"))), .TextMatrix(i, .ColIndex("MaintEndDate")), Date)  '.TextMatrix(i, .ColIndex("MaintEndDate"))
             RsDetails1.update
         
           End If
               Next i
            
        End With
         
        End If
    
        SaveGridPayment True, True
    
   Else
        
         
        Dim s As String
        Dim RsDetails2 As New ADODB.Recordset
        s = "Select * from TblContractInstallmentsHist Where ContNo = " & Trim(TxtContNo.Text)
        RsDetails2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        'If Not RsDetails2.EOF Then
            RsDetails2.AddNew
            RsDetails2!UserID = user_id
            RsDetails2!EditDate = Date
            RsDetails2!ContNo = val(TxtContNo)
            RsDetails2.update
        'End If
       
       SaveGridPayment False
       SaveGridPayment True
       
    
    End If
       '''//
       If 1 = 1 Then
           Set RsDetails1 = New ADODB.Recordset
             StrSQL = "SELECT     *  from dbo.TblIqrMerg Where (1 = -1)"
       RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If UnitsGrid.Rows > 1 Then
                    ' fg2.Rows = fg2.Rows - 1
                     With UnitsGrid
                                            For i = .FixedRows To .Rows - 1
                                            
                                                                     If .TextMatrix(i, .ColIndex("nameunittype")) <> "" Then
                                                              
                                                                  
                                                              
                                                                  RsDetails1.AddNew
                                                                 RsDetails1("Cont").value = val(TxtContNo.Text)
                                                         
                                                                   RsDetails1("Remark").value = .TextMatrix(i, .ColIndex("Remarks"))
                                                                '  RsDetails1("Elevatortype").value = .TextMatrix(i, .ColIndex("Elevatortype"))
                                                                  RsDetails1("Price").value = val(.TextMatrix(i, .ColIndex("meterPrice")))
                                                                  RsDetails1("Area").value = val(.TextMatrix(i, .ColIndex("length")))
                                                                  RsDetails1("TypeID").value = val(.TextMatrix(i, .ColIndex("unittype")))
                                                                  RsDetails1("UntID").value = val(.TextMatrix(i, .ColIndex("id")))
                                                                  RsDetails1("RentType").value = val(.TextMatrix(i, .ColIndex("namerentType")))
                                                                  Cn.Execute "  update TblAqarDetai  Set ContID=" & val(TxtContNo.Text) & ", Status = 1,meterPrice=" & val(.TextMatrix(i, .ColIndex("meterPrice"))) & ",RentValue=" & val(.TextMatrix(i, .ColIndex("RentValue"))) & "    ,customerid=" & val(dcCustomer.BoundText) & "  Where id =" & val(.TextMatrix(i, .ColIndex("id")))
                                                                ' RsDetails1("MaintEndDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("MaintEndDate"))), .TextMatrix(i, .ColIndex("MaintEndDate")), Date)  '.TextMatrix(i, .ColIndex("MaintEndDate"))
                                                                
                                                                RsDetails1.update
                                                            
                                                              End If
                                                Next i
                        
                    End With
         
        End If
    
        
       '''ContractSales//
           Set RsDetails1 = New ADODB.Recordset
             StrSQL = "SELECT     *  from  TblCOntractSales Where (1 = -1)"
       RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If VSFlexGrid2.Rows > 1 Then
                    ' fg2.Rows = fg2.Rows - 1
         With VSFlexGrid2
           For i = .FixedRows To .Rows - 1
           
                  If .TextMatrix(i, .ColIndex("empname")) <> "" Then
               RsDetails1.AddNew
               RsDetails1("ContNo").value = val(TxtContNo.Text)
               RsDetails1("rate").value = val(.TextMatrix(i, .ColIndex("rate")))
               RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
               RsDetails1("idd").value = val(.TextMatrix(i, .ColIndex("idd")))
               RsDetails1("GroupID").value = val(.TextMatrix(i, .ColIndex("groupid")))
             RsDetails1.update
         
           End If
               Next i
            
        End With
         
        End If
   '     If Not mCreateEntryManual Then
            
            If chkIsNotCreateEntry.value = vbUnchecked Then
                createVoucher
            End If
   '     End If
             
        Cn.Execute "  update TblAqarDetai  Set Status = 1,meterPrice=" & val(TxtMeterValue.Text) & ",RentValue=" & val(TxtTotalContract.Text) & ",Services=" & val(TxtPhone.Text) & ",Water=" & val(TxtWater.Text) & ",electric=" & val(TxtElectricity.Text) & "    ,customerid=" & val(dcCustomer.BoundText) & "  Where id =" & val(DcbUnitNo.BoundText)
       If TxtContNoOld.Text = "" Then
        Cn.Execute "  update TblAqarDetai  Set InsuranceValue=" & val(TxtInsuranceValue.Text) & ",Comm=" & val(TxtCommiValue.Text) & "    Where id =" & val(DcbUnitNo.BoundText)
       End If
       Cn.Execute "  update TblAqarDetai  Set ContID=" & val(TxtContNo.Text) & "  Where id =" & val(DcbUnitNo.BoundText)
       
    FillGridWithData
    saveinstdetailforpart2
    
    ReLineGrid
    GetUonitStatus
    SaveUoitInformation
    SaveInstalPreValue
   ' SaveVatNew
       Cn.CommitTrans

    TransBegine = False

    Screen.MousePointer = vbDefault




    
    FiLLTXT
    If val(TxtNotID.Text) > 0 Then
        Cn.Execute "Update Notes set PayedOrBon=1 where NoteID=" & val(TxtNotID.Text) & ""
    End If
End If
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CuurentLogdata
    TxtModFlg = "R"

'CuurentLogdata
    Exit Sub
ErrTrap:






    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault
    Dim Msg As String
    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰  ⁄·Ìﬁ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
            Msg = Msg & CHR(13) & Err.description
            Msg = Msg & CHR(13) & Err.Number
            Msg = Msg & CHR(13) & Err.Source
            Msg = Msg & CHR(13) & Err.LastDllError
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            Msg = "Can't Pending error in Data" & CHR(13)
        End If
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡  ⁄·Ìﬁ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry........Error During Save " & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub


Private Sub SaveVatNew()
        txtDateK.value = "14-05-2020"
           txtDateK2.value = "30-06-2020"
Dim mVatPercent As Double
Dim mVatPercent2 As Double
Dim rsDummyPercent As ADODB.Recordset
Dim s As String
Dim rsDummy As ADODB.Recordset
Dim i As Long
s = " SELECT "
s = s & "        tc.ContDate,TblContractInstallments.NoteID,tc.TotalContract,tc.InsuranceValue,tc.FATYou,tc.CusID,tc.ContType,tc.RentType,"
s = s & "        RetValue2,TblCustemers.CusName,TblAqar.aqartypeid,  tc.ownerid,tc.NotValue,tc.AccountCodeVat,"
s = s & "               InstrunceValue2,"
s = s & "               FATValue2,"
s = s & "               ElectricityValue2,installValue,"
s = s & "               WaterValue2,"
s = s & "               CommValue2,"
s = s & "               "
s = s & "                RetValue2 + InstrunceValue2 + FATValue2 + ElectricityValue2 + WaterValue2 + CommValue2 + Servce AS Arboun,"
s = s & "               tc.StrDate , tc.EndDate, tc.FromDateH, tc.TodateH,tc.Remarks,tc.Branch_NO,tc.TotalContract,tc.FATValue,"
s = s & " tc.NoteSerial1,tc.ContNo,TblContractInstallments.Id,"

s = s & "        TblContractInstallments.InstallNo,TblContractInstallments.Installdate ,TblContractInstallments.Countsofall, TblContractInstallments.RentValue, TblContractInstallments.Commissions, TblContractInstallments.Insurance, TblContractInstallments.Water, TblContractInstallments.Electric, TblContractInstallments.TelandNet, TblContractInstallments.VATValue, TblContractInstallments.allocations"
s = s & " FROM   TblContract AS tc"
s = s & "        INNER JOIN TblContractInstallments"
s = s & "             ON  TblContractInstallments.ContNo = tc.ContNo"
s = s & "             Left Outer join TblCustemers On tc.CusID = TblCustemers.CusID "
s = s & "             "
s = s & "                         Inner Join"
s = s & "                         TblAqar ON tc.Iqar =TblAqar.Aqarid"

s = s & " Where  1 = 1        "
'
'If Not IsNull(txtFromDate10.value) Then
'    s = s & " And TblContractInstallments.Installdate  >= " & SQLDate(txtFromDate10.value, True)
'End If



'If Not IsNull(txtToDate10.value) Then
'    s = s & " AND TblContractInstallments.Installdate <= " & SQLDate(txtToDate10.value, True)
'End If


s = s & "  and TblContractInstallments.ContNo IN (SELECT TblContract.ContNo"
s = s & "                    From TblContract"
s = s & "                    WHERE  ContNo = " & val(TxtContNo)

'Waelcomment s = s & "                  and  ISNULL(ComResid, 0) = 1)"
s = s & " )"
's = s & "                    AND TblContractInstallments.id NOT IN"
's = s & "                    (SELECT ContracttBillInstallmentsDone.istallid FROM ContracttBillInstallmentsDone)"
s = s & " ORDER BY tc.ContDate"
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic

Dim TransBegine  As Boolean
    i = 1
    If rsDummy.EOF Then
        MsgBox "·« ÌÊÃœ ⁄ﬁÊœ ›Ï  ·ﬂ «·› —…"
        TransBegine = False
        Cn.RollbackTrans
        Exit Sub
    End If
    Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String
    On Error GoTo NextRow

    
     Dim mCountDay1 As Integer
Dim mCountDay2 As Integer
Dim mCountDaysTotal As Integer
Dim mCostDay As Double
Dim mVATValue1 As Double
Dim mVATValue2 As Double

Do While Not rsDummy.EOF
    newinstallNo = val(rsDummy!InstallNo & "") + 1
    getnextDate newinstallNo, nextinstalldate, nextinstalldateH ', val(rsDummy!ContNo & "")
    If year(nextinstalldate) < 1900 Then
        nextinstalldate = Time
    End If
    
    
  
  
  mCountDaysTotal = DateDiff("D", rsDummy!installdate, nextinstalldate) '+ 1
If mCountDaysTotal = 0 Then mCountDaysTotal = 1
mCostDay = val(rsDummy!RentValue & "") / mCountDaysTotal
mVATValue2 = 0

If (SQLDate(rsDummy!ContDate, False)) > SQLDate(txtDateK.value, False) And (SQLDate(rsDummy!ContDate, False)) <= SQLDate(txtDateK2.value, False) Then
   
   If DateDiff("d", CDate(rsDummy!installdate & ""), txtDateK2.value) < 0 And DateDiff("d", txtDateK2.value, nextinstalldate) >= 0 Then
        mCountDay1 = mCountDaysTotal
        mVatPercent = 15
        mVatPercent2 = 0
        mVATValue1 = Round(val(val(rsDummy!RentValue & "") * mVatPercent / 100), 4)
    ElseIf DateDiff("d", CDate(rsDummy!installdate & ""), txtDateK2.value) >= 0 And DateDiff("d", txtDateK2.value, nextinstalldate) < 0 Then
        mCountDay1 = mCountDaysTotal
        mVatPercent = 5
        mVatPercent2 = 0
        mVATValue1 = Round(val(val(rsDummy!RentValue & "") * mVatPercent / 100), 4)
    ElseIf DateDiff("d", rsDummy!installdate, txtDateK2.value) >= 0 And DateDiff("d", txtDateK2.value, nextinstalldate) > 0 Then
        mVatPercent = 5
        mVatPercent2 = 15
        mCountDay1 = DateDiff("D", rsDummy!installdate, txtDateK2.value) '+ 1
        mCountDay2 = mCountDaysTotal - mCountDay1
        
        mVATValue1 = Round(val(mCostDay * mCountDay1 * mVatPercent / 100), 4)
        mVATValue2 = Round(val(mCostDay * mCountDay2 * mVatPercent2 / 100), 4)
    End If
    
    mCountDay2 = (mCountDaysTotal - mCountDay1)
  

    
  
   
    
 ElseIf (SQLDate(rsDummy!ContDate, False)) <= SQLDate(txtDateK.value, False) Then
        mVatPercent = 5
        mVatPercent2 = 0
        mVATValue1 = Round(val(rsDummy!RentValue & "") * mVatPercent / 100, 4)
        mCountDay1 = mCountDaysTotal
ElseIf (SQLDate(rsDummy!ContDate, False)) > SQLDate(txtDateK2.value, False) Then
    mVatPercent = 15
    mVatPercent2 = 0
    mVATValue1 = Round(val(rsDummy!RentValue & "") * mVatPercent / 100, 4)
    mCountDay1 = mCountDaysTotal
End If


        s = "Update TblContractInstallments Set "
         s = s & " "
        s = s & "  IsChangVat = 0,"
        s = s & "  CostDay = " & mCostDay & ","
       
        s = s & "  VATYou1 = " & mVatPercent
        s = s & " , VATYou2 =" & mVatPercent2
        s = s & " , CountDay1=" & mCountDay1
        s = s & " , CountDay2=" & mCountDay2
        
       ' If mVatPercent2 <> 0 Then
            s = s & " , VATValue1 = " & mVATValue1
            s = s & " , VATValue2 = " & mVATValue2
            s = s & " , VATValue=  " & mVATValue1 + mVATValue2
            
      '  Else
         '   s = s & " , VATValue = RentValue *  " & mVatPercent / 100
     '   End If

        
        s = s & " Where id = " & val(rsDummy!ID & "")
        Cn.Execute s

        
       
NextRow:
        rsDummy.MoveNext
Loop
        s = " update TblContract"
        s = s & " SET    FATValue = ("
        s = s & "            SELECT SUM(VATValue)"
        s = s & "            From TblContractInstallments"
        s = s & "            Where TblContractInstallments.ContNo = TblContract.ContNo"
        s = s & "                   AND ISNULL(VATValue, 0) <> 0"
        s = s & "        )"
        s = s & " From TblContract"
        s = s & " Where ContNo = " & val(TxtContNo & "")
        Cn.Execute s
        

End Sub
Private Sub SaveGridPayment(ByVal isOld As Boolean, Optional ByVal isNew As Boolean = False)
      
        
       Dim RsDetails1 As ADODB.Recordset
       Set RsDetails1 = New ADODB.Recordset
       Dim mTableName As String
       Dim StrSQL As String
       Dim i As Long
       mTableName = IIf(isOld, "TblContractInstallments", "TblContractInstallmentsOld")
       Dim s As String
       If Not isOld Then
            
            s = "Select * from TblContractInstallmentsOld Where ContNo = " & Trim(TxtContNo.Text)
            RsDetails1.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not RsDetails1.EOF Then
                Exit Sub
            Else
                s = "INSERT INTO TblContractInstallmentsOld"
                s = s & "    ("
                s = s & "         [ID],ContNo,InstallNo,Installdate,installValue,RentValue,Commissions,Insurance,Water,"
                s = s & "                     Electric,TelandNet,payed,Remains,RentValuePayed,CommissionsPayed,InsurancePayed,"
                s = s & "         WaterPayed , ElectricPayed, TelandNetPayed, [Status],lastPayedDate,VATPayed,VATValue,CommissionsArbon,NetCommissions,"
                s = s & " Insurance1,NetInsurance,Water1,WaterArbon,NetWater,Electric1,ElectricArbon,NetElectric,InsuranceArbon"
                s = s & "                   )"
                s = s & "                 SELECT [Id],ContNo,InstallNo,Installdate,installValue,RentValue,Commissions,Insurance,Water,"
                s = s & "                        Electric,TelandNet,payed,Remains,RentValuePayed,CommissionsPayed,InsurancePayed,"
                s = s & "                        WaterPayed , ElectricPayed, TelandNetPayed, [Status],lastPayedDate,VATPayed,VATValue,CommissionsArbon,NetCommissions,"
                s = s & "             Insurance1,NetInsurance,Water1,WaterArbon,NetWater,Electric1,ElectricArbon,NetElectric,InsuranceArbon"
                s = s & "                 From TblContractInstallments      "
                s = s & "                 Where ContNo = " & Trim(TxtContNo.Text)
                Cn.Execute s
                Exit Sub
            End If
            RsDetails1.Close
       Else
            If isNew Then
                StrSQL = "Delete From TblContractInstallments Where ContNo=" & val(Me.TxtContNo.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
'
       End If
       
       
       StrSQL = "SELECT     *  from " & mTableName & " Where   ContNo=" & val(Me.TxtContNo.Text)
       RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If GridInstallments.Rows > 1 Then
                    ' fg2.Rows = fg2.Rows - 1
         With GridInstallments
           For i = .FixedRows To .Rows - 1
           
        If i <> 0 Then
               If isNew Then
                    RsDetails1.AddNew
               End If
               RsDetails1("ContNo").value = val(TxtContNo.Text)
               RsDetails1("hijri").value = val(.TextMatrix(i, .ColIndex("hijri")))
               RsDetails1("TempInstal").value = val(.TextMatrix(i, .ColIndex("TempInstal")))
               RsDetails1("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo")))
               RsDetails1("Installdate").value = .TextMatrix(i, .ColIndex("Due_Date"))
               RsDetails1("InstalldateH").value = .TextMatrix(i, .ColIndex("Due_DateH"))
               RsDetails1("OldValueDate").value = IIf(.TextMatrix(i, .ColIndex("OldValueDate")) = "", Null, .TextMatrix(i, .ColIndex("OldValueDate")))
               RsDetails1("OldValueDateH").value = IIf(.TextMatrix(i, .ColIndex("OldValueDateH")) = "", Null, .TextMatrix(i, .ColIndex("OldValueDateH")))
               If mTableName = "TblContractInstallments" Then
                    RsDetails1("CountDay1").value = val(.TextMatrix(i, .ColIndex("CountDay1")))
                    RsDetails1("CountDay2").value = val(.TextMatrix(i, .ColIndex("CountDay2")))
                    RsDetails1("VATYou1").value = val(.TextMatrix(i, .ColIndex("VATYou1")))
                    RsDetails1("VATYou2").value = val(.TextMatrix(i, .ColIndex("VATYou2")))
                
                    RsDetails1("VATValue1Com").value = val(.TextMatrix(i, .ColIndex("VATValue1Com")))
                    RsDetails1("VATValue2Com").value = val(.TextMatrix(i, .ColIndex("VATValue2Com")))
                    RsDetails1("VATValue1").value = val(.TextMatrix(i, .ColIndex("VATValue1")))
                    RsDetails1("VATValue2").value = val(.TextMatrix(i, .ColIndex("VATValue2")))


  
  
    


               End If
                  'RsDetails1("OldValue").value = val(.TextMatrix(i, .ColIndex("OldValue")))
                  'RsDetails1("DES").value = (.TextMatrix(i, .ColIndex("DES")))
                  
              RsDetails1("installValue").value = val(.TextMatrix(i, .ColIndex("value")))
              
               RsDetails1("RentValue").value = val(.TextMatrix(i, .ColIndex("RentValue")))
               RsDetails1("NpayedValue").value = val(.TextMatrix(i, .ColIndex("NpayedValue")))
               RsDetails1("ServiceArbon").value = val(.TextMatrix(i, .ColIndex("ServiceArbon")))
              
              RsDetails1("Commissions").value = val(.TextMatrix(i, .ColIndex("Commissions")))
              RsDetails1("Insurance").value = val(.TextMatrix(i, .ColIndex("Insurance")))
              RsDetails1("Water").value = val(.TextMatrix(i, .ColIndex("Water")))
              RsDetails1("Electric").value = val(.TextMatrix(i, .ColIndex("Electric")))
            RsDetails1("TelandNet").value = val(.TextMatrix(i, .ColIndex("TelandNet")))
            RsDetails1("payed").value = val(.TextMatrix(i, .ColIndex("payed")))
            RsDetails1("Remains").value = val(.TextMatrix(i, .ColIndex("Remains")))
            RsDetails1("VATPayed").value = val(.TextMatrix(i, .ColIndex("VATPayed")))
            RsDetails1("VATValue").value = val(.TextMatrix(i, .ColIndex("VATValue")))
              RsDetails1("RentValuePayed").value = val(.TextMatrix(i, .ColIndex("RentValuePayed")))
        '      Payed = Payed + val(RsDetails1("RentValuePayed").value)
              RsDetails1("CommissionsPayed").value = val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
        '      Payed = Payed + val(RsDetails1("CommissionsPayed").value)
              RsDetails1("InsurancePayed").value = val(.TextMatrix(i, .ColIndex("InsurancePayed")))
        '      Payed = Payed + val(RsDetails1("InsurancePayed").value)
              RsDetails1("WaterPayed").value = val(.TextMatrix(i, .ColIndex("WaterPayed")))
        '      Payed = Payed + val(RsDetails1("WaterPayed").value)
              RsDetails1("ElectricPayed").value = val(.TextMatrix(i, .ColIndex("ElectricPayed")))
        '      Payed = Payed + val(RsDetails1("ElectricPayed").value)
            RsDetails1("TelandNetPayed").value = val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
        '    Payed = Payed + val(RsDetails1("TelandNetPayed").value)
              RsDetails1("Payed").value = val(.TextMatrix(i, .ColIndex("Payed")))
            '''///19 08 2015
            RsDetails1("Rent1").value = val(.TextMatrix(i, .ColIndex("Rent1")))
            RsDetails1("VATArboon").value = val(.TextMatrix(i, .ColIndex("VATArboon")))
            RsDetails1("RentArbon").value = val(.TextMatrix(i, .ColIndex("RentArbon")))
            RsDetails1("NetRent").value = val(.TextMatrix(i, .ColIndex("NetRent")))
            RsDetails1("Commissions1").value = val(.TextMatrix(i, .ColIndex("Commissions1")))
            RsDetails1("CommissionsArbon").value = val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
            RsDetails1("NetCommissions").value = val(.TextMatrix(i, .ColIndex("NetCommissions")))
            RsDetails1("Insurance1").value = val(.TextMatrix(i, .ColIndex("Insurance1")))
            RsDetails1("InsuranceArbon").value = val(.TextMatrix(i, .ColIndex("InsuranceArbon")))
            RsDetails1("NetInsurance").value = val(.TextMatrix(i, .ColIndex("NetInsurance")))
            RsDetails1("Water1").value = val(.TextMatrix(i, .ColIndex("Water1")))
            RsDetails1("WaterArbon").value = val(.TextMatrix(i, .ColIndex("WaterArbon")))
            RsDetails1("NetWater").value = val(.TextMatrix(i, .ColIndex("NetWater")))
            RsDetails1("Electric1").value = val(.TextMatrix(i, .ColIndex("Electric1")))
            RsDetails1("ElectricArbon").value = val(.TextMatrix(i, .ColIndex("ElectricArbon")))
            RsDetails1("NetElectric").value = val(.TextMatrix(i, .ColIndex("NetElectric")))
    
    'RsDetails1("OldValue").value = val(.TextMatrix(i, .ColIndex("OldValue")))
    
            If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then
            RsDetails1("Status").value = 1
               RsDetails1("allocations").value = 1
            Else
            RsDetails1("Status").value = 0
               RsDetails1("allocations").value = 0
            End If
            
    If Opt(0).value = True Then '⁄ﬁœ ÃœÌœ
           If SystemOptions.WorkWithFirstInstallOnly = True Then ' «À»«  «·«” Õﬁ«ﬁ «Ê· ﬁÌœ ›ﬁÿ
                       
                                If i = 1 Then '«Ê· ﬁ”ÿ
                                   RsDetails1("allocations").value = 1
                                   Else
                                    RsDetails1("allocations").value = 0
                                End If
                                
               Else
               
                                RsDetails1("allocations").value = 1
            End If
        
      End If
            
            
            '  Status
                 RsDetails1("NoteSerial").value = (.TextMatrix(i, .ColIndex("NoteSerial")))
                 RsDetails1("NoteSerial1").value = (.TextMatrix(i, .ColIndex("NoteSerial1")))
                 RsDetails1("NoteId").value = val(.TextMatrix(i, .ColIndex("NoteId")))
                 
                 
                 RsDetails1("OldValueDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("OldValueDate"))), .TextMatrix(i, .ColIndex("OldValueDate")), Null)
               RsDetails1("OldValueDateH").value = .TextMatrix(i, .ColIndex("OldValueDateH"))
               
             RsDetails1("lastPayedDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("lastPayedDate"))), .TextMatrix(i, .ColIndex("lastPayedDate")), Null)
               RsDetails1("lastPayedDateH").value = .TextMatrix(i, .ColIndex("lastPayedDateH"))
              'RsDetails1("allocations").value = val(.TextMatrix(i, .ColIndex("allocations")))
               '
               RsDetails1("Countsofall").value = val(.TextMatrix(i, .ColIndex("Countsofall")))
               RsDetails1("Doneofall").value = val(.TextMatrix(i, .ColIndex("Doneofall")))
                         
                         
               RsDetails1.update
         
           End If
                RsDetails1.MoveNext
               Next i
            RsDetails1.Close
        End With
        
        End If
End Sub

Function SHOWPIC(PICNAME As String)
   
End Function


Function print_report(ID As Double)
    On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

 
'SALIMHERE 05 03 2019 FINGER PRINT
MySQL = " SELECT  TblContract.NewNO, TblContract.Accredit,     TblContract.DiscountvaLUE   ,TblContract.DiscountPercent,        dbo.TblContract.FATYou, dbo.TblContract.FATValue ,      dbo.TblContract.Contract_period, dbo.TblContract.Contract_period_no, dbo.TblContract.StrMerg, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqartypeid, dbo.TblAqar.CountryID, dbo.TblCountriesData.CountryName, "
MySQL = MySQL & "                          dbo.TblAqar.cityid, dbo.TblCountriesGovernments.GovernmentName, dbo.TblAqar.heyid, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblAqar.streetname, dbo.TblAqar.schemeid, dbo.tblSchemes.name AS SchemeName,"
MySQL = MySQL & "                          dbo.tblSchemes.namee AS SchemeNameE, dbo.TblAqar.StatusId, dbo.TblAqar.floorcount, dbo.TblAqar.Location, dbo.TblAqar.aqarname, dbo.TblContract.ContNo, dbo.TblContract.ContType, dbo.TblContract.ContDate,"
MySQL = MySQL & "                          dbo.TblContract.Iqar, dbo.TblContract.UnitType, dbo.TblAkarUnit.namee, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblContract.RentType, dbo.TblContract.StrDate, dbo.TblContract.EndDate,"
MySQL = MySQL & "                          dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini, dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water,"
MySQL = MySQL & "                          dbo.TblContract.Electricity, dbo.TblContract.Phone, dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate,"
MySQL = MySQL & "                          dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.ResponsibleContact,"
MySQL = MySQL & "                          dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.CountryID2, dbo.Nationality.name AS Natinname, dbo.Nationality.namee AS NatinnameE, dbo.TblContract.Furnishing, dbo.TblContract.Remarks,"
MySQL = MySQL & "                          dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.TodateH, dbo.TblContract.FirstInstallDateH, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                          dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1, dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.E_mail,"
MySQL = MySQL & "                          dbo.TblCustemers.JobAddress, dbo.TblCustemers.Address, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.length,"
MySQL = MySQL & "                          dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.Water AS WaterDet, dbo.TblAqarDetai.electric,"
MySQL = MySQL & "                          dbo.TblAqarDetai.ACCountspleat, dbo.TblAqarDetai.UnitElectric, dbo.TblCustemers.CustGIDPlace, dbo.TblAkarUnit.name, dbo.tblAkarType.name AS AqrType, dbo.tblAkarType.namee AS AqrTypeE, dbo.TblCustemers.BrithDateH,"
MySQL = MySQL & "                          dbo.TblCustemers.BrithDate , dbo.TblCustemers.recordno, dbo.tblCustomerFingers.ItemPhoto, dbo.tblCustomerFingers.ItemPhoto1, dbo.tblCustomerFingers.ItemPhoto3"
MySQL = MySQL & "  FROM            dbo.TblAkarUnit RIGHT OUTER JOIN"
MySQL = MySQL & "                          dbo.tblCustomerFingers    RIGHT OUTER JOIN  "
MySQL = MySQL & "                          dbo.TblContract ON dbo.tblCustomerFingers.FCusID = dbo.TblContract.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblCustemers LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.Nationality ON dbo.TblCustemers.CountryID2 = dbo.Nationality.id ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id ON dbo.TblAkarUnit.id = dbo.TblContract.UnitType LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.tblAkarType RIGHT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblAqar ON dbo.tblAkarType.id = dbo.TblAqar.aqartypeid ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.tblSchemes ON dbo.TblAqar.schemeid = dbo.tblSchemes.id LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblCountriesGovernments INNER JOIN"
MySQL = MySQL & "                          dbo.TblCountriesGovernmentsCities ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID INNER JOIN"
MySQL = MySQL & "                          dbo.TblCountriesData ON dbo.TblCountriesGovernments.CountryID = dbo.TblCountriesData.CountryID ON dbo.TblAqar.heyid = dbo.TblCountriesGovernmentsCities.CityID AND"
MySQL = MySQL & "                          dbo.TblAqar.CityID = dbo.TblCountriesGovernments.GovernmentID And dbo.TblAqar.CountryID = dbo.TblCountriesData.CountryID"
MySQL = MySQL & "  Where (dbo.TblContract.ContNo= " & val(TxtContNo.Text) & ")"


   If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Real Etstae\" & "CotractRent.rpt"
    Else
        StrFileName = App.path & "\Reports\Real Etstae\" & "CotractRent.rpt"
    End If


   ' If SystemOptions.UserInterface = ArabicInterface Then
   '     StrFileName = App.path & "\Reports\Real Etstae\" & "Cotract.rpt"
   ' Else
   '     StrFileName = App.path & "\Reports\Real Etstae\" & "Cotract.rpt"
   ' End If

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
    
       
 
        StrReportTitle = ""
 Dim Diff As Double
 Dim diffs As String
    End If
    If val(TxtPaymentCount) > 2 Then
    Diff = Round(val(TxtTotalContract.Text) / val(TxtPaymentCount), 2)
    End If
    diffs = Diff
xReport.ParameterFields(4).AddCurrentValue WriteNo(val(TxtTotalContract.Text), 0, True)
xReport.ParameterFields(7).AddCurrentValue WriteNo(diffs, 0, True)
Dim i As Integer
Dim Units As String
Units = ""
For i = 1 To UnitsGrid.Rows - 1
If val(Me.UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id"))) <> 0 Then
Units = Units & UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno")) & " ,"
End If
Next i
xReport.ParameterFields(5).AddCurrentValue Units
    xReport.ParameterFields(3).AddCurrentValue user_name
   If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(11).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(11).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    
    If CheckFp.value = vbChecked Then
           'SHOWPIC ()
 Dim xLogo As CRAXDRT.OLEObject
    Dim StrFileName1 As String
   
    StrFileName1 = App.path & "\images\FP\RsCustomers\" & val(Me.dcCustomer.BoundText) & ".JPG"

    Set xLogo = xReport.Areas(1).Sections(1).AddPictureObject(StrFileName1, 5000, 13200)
    xLogo.Width = 1300
    xLogo.Height = 1500
    xLogo.backcolor = vbWhite
    xLogo.BorderColor = 255
    xLogo.CloseAtPageBreak = True
 
    End If
    
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL
 
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Function createVoucher()
If SystemOptions.NoCreatJLInRentContract = True Then Exit Function
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "⁄ﬁœ «ÌÃ«— —ﬁ„ " & TxtNoteSerial & " · " & dcCustomer.Text
des = des & "   «·› —… „‰  " & FromDateH.value & " «·Ì " & todateH.value
des = des & " «·„Ê«›ﬁ " & StrDate.value & " «·Ì " & EndDate.value


des = des & " " & TxtRemark2.Text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "TblContract"
Filedname = "ContNo"
ContNo = val(TxtContNo)
Notevalue = 0

If SystemOptions.WorkWithFirstInstallOnly = False Then
Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)
Else

With GridInstallments

If .Rows > 1 Then
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("RentValue")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Commissions")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Insurance")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Water")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Electric")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("TelandNet")))

 
End If

End With


End If
If FlagContrNew2 = True Or lblnew.Visible = True Then
With GridInstallments
If .Rows > 1 Then
DTPicker2.value = CDate(.TextMatrix(1, .ColIndex("Due_Date")))
End If
End With
Else
DTPicker2.value = FristPaymentDate.value
End If

If Opt(0).value = True And Notevalue > 0 Then
                    If Me.TxtModFlg = "N" Then
                    
                          CreateNotes NoteID, (DTPicker2.value), val(Dcbranch.BoundText), 60, Notevalue, NoteSerial, val(TxtNoteSerial1), tablename, Filedname, ContNo, des, ToHijriDate(DTPicker2.value)   'RecorddateH.value
                                  TxtNoteID.Text = NoteID
                                         TxtNoteSerial.Text = NoteSerial
                         Else
                                     If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                CreateNotes NoteID, (DTPicker2.value), val(Dcbranch.BoundText), 60, Notevalue, NoteSerial, val(TxtNoteSerial1), tablename, Filedname, ContNo, des, ToHijriDate(DTPicker2.value)
                                                     TxtNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                                       Else
                                                     sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                    sql = sql & ",NoteSerial1='" & val(TxtNoteSerial1) & "'"
                                                       sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                       Cn.Execute sql
                                                   
                                     End If
                           
                    End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), val(Dcbranch.BoundText), user_id, DTPicker2.value
RsSavRec.Resync adAffectCurrent
      Else

CreateOpeningBalanceRecord

     End If

End Function
Function createVoucher2(Optional Row As Long)
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
With GridInstallments
des = "⁄ﬁœ «ÌÃ«— —ﬁ„ " & TxtNoteSerial & " · " & dcCustomer.Text & " "
des = des & "   «·› —… „‰  " & FromDateH.value & " «·Ì " & todateH.value
des = des & " «·„Ê«›ﬁ " & StrDate.value & " «·Ì " & EndDate.value

des = "«·œ›⁄… —ﬁ„" & .TextMatrix(Row, .ColIndex("InstallNo"))
des = des & " " & TxtRemark2.Text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
Dim installIDCont As Double
tablename = "TblContractInstallments"
Filedname = "id"
ContNo = val(.TextMatrix(Row, .ColIndex("Installid")))
Notevalue = val(.TextMatrix(Row, .ColIndex("VATValue")))
DTPicker2.value = CDate(.TextMatrix(Row, .ColIndex("Due_Date")))

installIDCont = val(.TextMatrix(Row, .ColIndex("Installid")))
Cn.Execute "delete notes where installIDCont=" & installIDCont


        If .TextMatrix(Row, .ColIndex("NoteSerial1")) = "" Then
              .TextMatrix(Row, .ColIndex("NoteSerial1")) = Voucher_coding(val(Me.Dcbranch.BoundText), DTPicker2.value, 75, 75)
              Cn.Execute "Update TblContractInstallments set NoteSerial1='" & .TextMatrix(Row, .ColIndex("NoteSerial1")) & "' where ID =" & ContNo & " "
        End If

                If Notevalue > 0 Then
                    
                   CreateNotes NoteID, (DTPicker2.value), val(Dcbranch.BoundText), 9088, Notevalue, NoteSerial, val(.TextMatrix(Row, .ColIndex("InstallNo"))), tablename, Filedname, ContNo, des, ToHijriDate(DTPicker2.value), , , , CStr(installIDCont)  'RecorddateH.value
                   .TextMatrix(Row, .ColIndex("NoteId")) = NoteID
                   .TextMatrix(Row, .ColIndex("NoteSerial")) = NoteSerial

CREATE_VOUCHER_GE2 val(.TextMatrix(Row, .ColIndex("NoteId"))), val(Dcbranch.BoundText), user_id, DTPicker2.value, Row
FindRec val(Me.TxtContNo.Text)
     End If
End With
End Function
Sub CreateSearial()
Dim Row As Long
Dim ContNo As Double

With GridInstallments
For Row = 1 To .Rows - 1
If .TextMatrix(Row, .ColIndex("Due_Date")) <> "" Then
    DTPicker2.value = CDate(.TextMatrix(Row, .ColIndex("Due_Date")))
    ContNo = 0
    ContNo = val(.TextMatrix(Row, .ColIndex("Installid")))
    If ContNo <> 0 Then
           If .TextMatrix(Row, .ColIndex("NoteSerial1")) = "" Then
                  .TextMatrix(Row, .ColIndex("NoteSerial1")) = Voucher_coding(val(Me.Dcbranch.BoundText), DTPicker2.value, 75, 75)
                  Cn.Execute "Update TblContractInstallments set NoteSerial1='" & .TextMatrix(Row, .ColIndex("NoteSerial1")) & "' where ID =" & ContNo & " "
            End If
     End If
    End If
  Next Row
End With
 FindRec val(Me.TxtContNo.Text)
End Sub
Function CreateOpeningBalanceRecord()
Dim StrDes As String
Dim LngDevID As Long
 Dim LngOpenID As Long
Dim StrTempAccountCode As String
  Dim Notevalue As Single
  Dim FirstPeriodDateInthisYear As Date
  '  Notevalue = val(LblNotPayed.Caption)
     Notevalue = val(TxtOldRent) + val(TxtOldWater) + val(TxtOldElectric) + val(TxtoldCommi) + val(Me.txtOldInsurance)
        If Notevalue = 0 Then Exit Function
        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹··⁄ﬁœ —ﬁ„  " & Trim(Me.TxtNoteSerial1.Text) & "  ··⁄„Ì·  " & dcCustomer.Text
        Else
            StrDes = " Opening Balance For: " & Trim(Me.TxtNoteSerial1.Text) & " " & " customer  " & dcCustomer.Text
        End If
  


      LngOpenID = 1
   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
     StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
  
        getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

'«·Ã“¡ «·„ »ﬁÌ ⁄·Ï «·⁄„Ì·
    Dim lineno As Integer
    lineno = 1
    Notevalue = val(TxtOldRent) + val(TxtOldWater) + val(TxtOldElectric) + val(TxtoldCommi)
  If Notevalue > 0 Then
    If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, Notevalue, 0, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , val(Dcbranch.BoundText), , , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
End If
  If val(TxtOldRent) > 0 Then

     lineno = lineno + 1
                    If ModAccounts.AddNewDev(LngDevID, lineno, Account_Code_dynamic59, val(TxtOldRent), 1, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , , val(Dcbranch.BoundText), , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
   End If
   
  If val(TxtOldWater) > 0 Then
    lineno = lineno + 1
                    If ModAccounts.AddNewDev(LngDevID, lineno, Account_Code_dynamic59, val(TxtOldWater), 1, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , , val(Dcbranch.BoundText), , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
  End If
                    
   If val(TxtOldElectric) > 0 Then
  lineno = lineno + 1
                    If ModAccounts.AddNewDev(LngDevID, lineno, Account_Code_dynamic59, val(TxtOldElectric), 1, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , , val(Dcbranch.BoundText), , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
                    
   End If
     If val(TxtoldCommi) > 0 Then
  lineno = lineno + 1
                    If ModAccounts.AddNewDev(LngDevID, lineno, Account_Code_dynamic59, val(TxtoldCommi), 1, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , , val(Dcbranch.BoundText), , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
   End If
   
        If val(txtOldInsurance) > 0 Then
    Notevalue = val(txtOldInsurance)
    
         If ModAccounts.AddNewDev(LngDevID, lineno, Account_Code_dynamic92, val(txtOldInsurance), 0, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , , val(Dcbranch.BoundText), , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
        
        If SystemOptions.CreateInsuranceAccountForCustomers Then
    StrTempAccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText), "InsuranceAccount")
  End If
  
        
    If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, Notevalue, 1, StrDes, LngOpenID, , , , FirstPeriodDateInthisYear, , , , , , , , , , , , , , True, -1, , , val(Dcbranch.BoundText), , , , , , , , , , val(TxtContNo)) = False Then
                        GoTo ErrTrap
                    End If
                
     lineno = lineno + 1
            
     End If
                    
'    —’Ìœ „” Õﬁ ··⁄„Ì·
                    
ErrTrap:
End Function
Public Sub FiLLTXT()
Label400.Visible = False
    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
     Me.LblTotalQasts.Caption = 0
     TXTNewNO.Text = IIf(IsNull(RsSavRec("NewNO").value), "", RsSavRec("NewNO").value)
     AccountVat.BoundText = IIf(IsNull(RsSavRec("AccountCodeVat").value), "", RsSavRec("AccountCodeVat").value)
     TxtRemark2.Text = IIf(IsNull(RsSavRec("Remark2").value), "", RsSavRec("Remark2").value)
     AccountVat2.BoundText = IIf(IsNull(RsSavRec("AccountCodeVat2").value), "", RsSavRec("AccountCodeVat2").value)
     RecorddateH.value = IIf(IsNull(RsSavRec("RecorddateH").value), ToHijriDate(Date), RsSavRec("RecorddateH").value)
     FromDateH.value = IIf(IsNull(RsSavRec("FromdateH").value), ToHijriDate(Date), RsSavRec("FromdateH").value)
     todateH.value = IIf(IsNull(RsSavRec("TodateH").value), ToHijriDate(Date), RsSavRec("TodateH").value)
     FirstInstallDateH.value = IIf(IsNull(RsSavRec("FirstInstallDateH").value), ToHijriDate(Date), RsSavRec("FirstInstallDateH").value)
     Dcbranch.BoundText = IIf(IsNull(RsSavRec("Branch_NO").value), 0, (RsSavRec("Branch_NO").value))
     Me.TxtMiniRentValue.Text = IIf(IsNull(RsSavRec.Fields("MiniRentValue").value), "", RsSavRec.Fields("MiniRentValue").value)
     Me.DcboEmp.BoundText = IIf(IsNull(RsSavRec("Emp_ID").value), "", RsSavRec("Emp_ID").value)
     Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
     txtDiscountPercent.Text = IIf(IsNull(RsSavRec.Fields("DiscountPercent").value), 0, RsSavRec.Fields("DiscountPercent").value)
     TxtDiscountValue.Text = IIf(IsNull(RsSavRec.Fields("DiscountvaLUE").value), 0, RsSavRec.Fields("DiscountvaLUE").value)
     
      
    If Not IsNull(RsSavRec("TypeDate").value) Then
                If RsSavRec("TypeDate").value = 1 Then
                         RdRTypeDate(1).value = True
                Else
                           RdRTypeDate(0).value = True
                End If
    Else
                          RdRTypeDate(0).value = True
    End If
    
    If Not IsNull(RsSavRec("FlagContrNew2").value) Then
    If RsSavRec("FlagContrNew2").value = True Then
    FlagContrNew2 = True
    Else
    FlagContrNew2 = False
    End If
    Else
    FlagContrNew2 = False
    End If
    
    If Not (IsNull(RsSavRec("CommiValueInVAT").value)) Then
        If RsSavRec("CommiValueInVAT").value = 1 Then
            CommiValueInVAT.value = vbChecked
        Else
            CommiValueInVAT.value = vbUnchecked
        End If
    Else
        CommiValueInVAT.value = vbUnchecked
    End If
    
    
    If Not (IsNull(RsSavRec("IsNotCreateEntry").value)) Then
        If RsSavRec("IsNotCreateEntry").value = 1 Then
            chkIsNotCreateEntry.value = vbChecked
        Else
            chkIsNotCreateEntry.value = vbUnchecked
        End If
    Else
        chkIsNotCreateEntry.value = vbUnchecked
    End If
    
    
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec("UserID").value), user_id, RsSavRec("UserID").value)

        
    If Not (IsNull(RsSavRec("WaterElecValueInVAT").value)) Then
    If RsSavRec("WaterElecValueInVAT").value = 1 Then
    WaterElecValueInVAT.value = vbChecked
    Else
    WaterElecValueInVAT.value = vbUnchecked
    End If
    Else
    WaterElecValueInVAT.value = vbUnchecked
    End If
    
    
        If Not (IsNull(RsSavRec("InsurValueInVAT").value)) Then
    If RsSavRec("InsurValueInVAT").value = 1 Then
    InsurValueInVAT.value = vbChecked
    Else
    InsurValueInVAT.value = vbUnchecked
    End If
    Else
    InsurValueInVAT.value = vbUnchecked
    End If
    
    
 If Not (IsNull(RsSavRec("MethodDeci").value)) Then
 If RsSavRec("MethodDeci").value = 0 Then
 Opt(4).value = True
 ElseIf RsSavRec("MethodDeci").value = 1 Then
 Opt(3).value = True
 ElseIf RsSavRec("MethodDeci").value = 2 Then
 Opt(2).value = True
 End If
End If

If IsNull(RsSavRec.Fields("NewOrOpeneing").value) Then
Opt(0).value = True
Else

                If RsSavRec.Fields("NewOrOpeneing").value = 0 Then
                
                        Opt(0).value = True
                Else
                          Opt(1).value = True
                Me.TxtNoteSerial.Text = ""
                End If
 
End If
 
 If IsNull(RsSavRec.Fields("OutContract").value) Then
 ChKOutContract.value = vbUnchecked
 Else
 ChKOutContract.value = vbChecked
   
 End If
   
   'ChKEndContract
   
   
 If IsNull(RsSavRec.Fields("Employeecontract").value) Then
 ChkEmployeecontract.value = vbUnchecked
 Else
 ChkEmployeecontract.value = vbChecked
   
 End If
    
    
 If IsNull(RsSavRec.Fields("Accredit").value) Then
 ChkAccredit.value = vbUnchecked
 Else
 ChkAccredit.value = vbChecked
   
 End If
        
        
 If IsNull(RsSavRec.Fields("EndContract").value) Then
 ChKEndContract.value = vbUnchecked
 Else
 ChKEndContract.value = vbChecked
   
 End If
 
 
 
  If IsNull(RsSavRec.Fields("LegalIssue").value) Then
 ChKLegalIssue.value = vbUnchecked
 Else
 ChKLegalIssue.value = vbChecked
   
 End If
 
 
 
   
 If IsNull(RsSavRec.Fields("DivWater").value) Then
 chkDivWater.value = vbUnchecked
 Else
 chkDivWater.value = vbChecked
  
 
 End If
 
 
 
If IsNull(RsSavRec.Fields("Renew").value) Then
lblnew.Visible = False
ChkRenew.value = vbUnchecked
FrmContractOldData.Visible = False

Else

                If RsSavRec.Fields("Renew").value = 0 Then
                lblnew.Visible = False
                      ChkRenew.value = vbUnchecked
                      FrmContractOldData.Visible = False

                Else
                lblnew.Visible = True
                     ChkRenew.value = vbChecked
                     FrmContractOldData.Visible = True
                     lblnew.Caption = " „ «· ÃœÌœ"

                End If
 
End If

  If IsNull(RsSavRec.Fields("DivElectric").value) Then
 chkDivElectric.value = vbUnchecked
 Else
 chkDivElectric.value = vbChecked
  
 
 End If
 Me.TxtElectAccount.Text = IIf(IsNull(RsSavRec.Fields("UnitElectric").value), "", RsSavRec.Fields("UnitElectric").value)
 
Me.TxtOthersRules.Text = IIf(IsNull(RsSavRec.Fields("OthersRules").value), "", RsSavRec.Fields("OthersRules").value)


Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial1").value), "", RsSavRec.Fields("NoteSerial1").value)
Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
     
     
         Me.TxtContNoOld.Text = IIf(IsNull(RsSavRec.Fields("ContNoOld").value), "", RsSavRec.Fields("ContNoOld").value)

   DcboEmpName.BoundText = IIf(IsNull(RsSavRec("Emp_IDContract").value), "", RsSavRec("Emp_IDContract").value)

    Me.TxtContNo.Text = IIf(IsNull(RsSavRec.Fields("ContNo").value), "", RsSavRec.Fields("ContNo").value)
    Me.ContDate.value = IIf(IsNull(RsSavRec.Fields("ContDate").value), Date, RsSavRec.Fields("ContDate").value)
   Me.DcbContType.ListIndex = val(IIf(IsNull(RsSavRec.Fields("ContType").value), -1, RsSavRec.Fields("ContType").value))
   Me.DcbIqara.BoundText = val(IIf(IsNull(RsSavRec.Fields("Iqar").value), 0, RsSavRec.Fields("Iqar").value))
   Me.dcCustomer.BoundText = val(IIf(IsNull(RsSavRec.Fields("CusID").value), 0, RsSavRec.Fields("CusID").value))
     Me.dcsupplier.BoundText = val(IIf(IsNull(RsSavRec.Fields("ownerid").value), 0, RsSavRec.Fields("ownerid").value))
     Me.DcbUnitType.BoundText = val(IIf(IsNull(RsSavRec.Fields("UnitType").value), -1, RsSavRec.Fields("UnitType").value))
  ReloadUonit
   
     Me.DcbUnitNo.BoundText = val(IIf(IsNull(RsSavRec.Fields("UnitNo").value), -1, RsSavRec.Fields("UnitNo").value))
     Me.DcbRentType.ListIndex = val(IIf(IsNull(RsSavRec.Fields("RentType").value), -1, RsSavRec.Fields("RentType").value))
     Me.StrDate.value = IIf(IsNull(RsSavRec.Fields("StrDate").value), Date, RsSavRec.Fields("StrDate").value)
     Me.EndDate.value = IIf(IsNull(RsSavRec.Fields("EndDate").value), Date, RsSavRec.Fields("EndDate").value)
   Me.TxtMeterValue.Text = IIf(IsNull(RsSavRec.Fields("MeterValue").value), "", RsSavRec.Fields("MeterValue").value)
   Me.TxtMeterCount.Text = IIf(IsNull(RsSavRec.Fields("MeterCount").value), "", RsSavRec.Fields("MeterCount").value)
   Me.TxtTotalContract.Text = IIf(IsNull(RsSavRec.Fields("TotalContract").value), "", RsSavRec.Fields("TotalContract").value)
   Me.TxtPayAmini.Text = IIf(IsNull(RsSavRec.Fields("PayAmini").value), "", RsSavRec.Fields("PayAmini").value)
   Me.TxtCommiValue.Text = IIf(IsNull(RsSavRec.Fields("CommiValue").value), "", RsSavRec.Fields("CommiValue").value)
   Me.TxtInsuranceValue.Text = IIf(IsNull(RsSavRec.Fields("InsuranceValue").value), "", RsSavRec.Fields("InsuranceValue").value)
   ''//
    Me.TxtNotID.Text = IIf(IsNull(RsSavRec.Fields("NotID").value), "", RsSavRec.Fields("NotID").value)
   Me.TxtNotSreail1.Text = IIf(IsNull(RsSavRec.Fields("NoteSrial1").value), "", RsSavRec.Fields("NoteSrial1").value)
   Me.TxtNotVal.Text = IIf(IsNull(RsSavRec.Fields("NotValue").value), "", RsSavRec.Fields("NotValue").value)
   ''//
   ''//
   
      Me.TxtRetValue2.Text = IIf(IsNull(RsSavRec.Fields("RetValue2").value), "", RsSavRec.Fields("RetValue2").value)
      Me.TxtFATValue2.Text = IIf(IsNull(RsSavRec.Fields("FATValue2").value), "", RsSavRec.Fields("FATValue2").value)
      Me.TxtWaterValue2.Text = IIf(IsNull(RsSavRec.Fields("WaterValue2").value), "", RsSavRec.Fields("WaterValue2").value)
      Me.TxtCommValue2.Text = IIf(IsNull(RsSavRec.Fields("CommValue2").value), "", RsSavRec.Fields("CommValue2").value)
      Me.TxtInstrunceValue2.Text = IIf(IsNull(RsSavRec.Fields("InstrunceValue2").value), "", RsSavRec.Fields("InstrunceValue2").value)
      Me.TxtElectricityValue2.Text = IIf(IsNull(RsSavRec.Fields("ElectricityValue2").value), "", RsSavRec.Fields("ElectricityValue2").value)
      Me.TxtServce.Text = IIf(IsNull(RsSavRec.Fields("Servce").value), 0, RsSavRec.Fields("Servce").value)
   ''//
   
   Me.TxtOutOffice.Text = IIf(IsNull(RsSavRec.Fields("OutOffice").value), "", RsSavRec.Fields("OutOffice").value)
   

    Me.TxtWater.Text = IIf(IsNull(RsSavRec.Fields("Water").value), "", RsSavRec.Fields("Water").value)
   Me.TxtElectricity.Text = IIf(IsNull(RsSavRec.Fields("Electricity").value), "", RsSavRec.Fields("Electricity").value)
   Me.TxtPhone.Text = IIf(IsNull(RsSavRec.Fields("Phone").value), "", RsSavRec.Fields("Phone").value)
   Me.TxtEnternet.Text = IIf(IsNull(RsSavRec.Fields("Enternet").value), "", RsSavRec.Fields("Enternet").value)
   Me.TxtIncresYearValue.Text = IIf(IsNull(RsSavRec.Fields("IncresYearValue").value), "", RsSavRec.Fields("IncresYearValue").value)
     Me.TxtIncresYearRate.Text = IIf(IsNull(RsSavRec.Fields("IncresYearRate").value), "", RsSavRec.Fields("IncresYearRate").value)
   Me.TxtPaymentCount.Text = IIf(IsNull(RsSavRec.Fields("PaymentCount").value), "", RsSavRec.Fields("PaymentCount").value)
   Me.TxtPeriods.Text = IIf(IsNull(RsSavRec.Fields("Periods").value), "", RsSavRec.Fields("Periods").value)
   Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
   Me.FristPaymentDate.value = IIf(IsNull(RsSavRec.Fields("FristPaymentDate").value), Date, RsSavRec.Fields("FristPaymentDate").value)
   Me.DcbPeriodsID.ListIndex = val(IIf(IsNull(RsSavRec.Fields("PeriodsID").value), -1, RsSavRec.Fields("PeriodsID").value))
   Me.dcsupplier.BoundText = val(IIf(IsNull(RsSavRec.Fields("ownerid").value), 0, RsSavRec.Fields("ownerid").value))
   Me.DcbFurnishing.ListIndex = val(IIf(IsNull(RsSavRec.Fields("Furnishing").value), -1, RsSavRec.Fields("Furnishing").value))
   Me.TxtOldRent.Text = IIf(IsNull(RsSavRec.Fields("OldRent").value), "", (RsSavRec.Fields("OldRent").value))
   Me.TxtOldWater.Text = IIf(IsNull(RsSavRec.Fields("OldWater").value), "", (RsSavRec.Fields("OldWater").value))
   Me.TxtOldElectric.Text = IIf(IsNull(RsSavRec.Fields("OldElectric").value), "", (RsSavRec.Fields("OldElectric").value))
   Me.TxtoldCommi.Text = IIf(IsNull(RsSavRec.Fields("oldCommi").value), "", (RsSavRec.Fields("oldCommi").value))
   Me.txtOldInsurance.Text = IIf(IsNull(RsSavRec.Fields("OldInsurance").value), "", (RsSavRec.Fields("OldInsurance").value))
   Me.balanceDate.value = IIf(IsNull(RsSavRec.Fields("balanceDate").value), Date, RsSavRec.Fields("balanceDate").value)
   balanceDateH.value = IIf(IsNull(RsSavRec("balanceDateH").value), ToHijriDate(Date), RsSavRec("balanceDateH").value)
   Me.balanceDes.Text = IIf(IsNull(RsSavRec.Fields("balanceDes").value), "", RsSavRec.Fields("balanceDes").value)
   Me.FromdateO.value = IIf(IsNull(RsSavRec.Fields("FromdateO").value), Date, RsSavRec.Fields("FromdateO").value)
   FromdateHO.value = IIf(IsNull(RsSavRec("FromdateHO").value), ToHijriDate(Date), RsSavRec("FromdateHO").value)
   Me.TxtNetValue.Text = IIf(IsNull(RsSavRec.Fields("NetValue").value), 0, (RsSavRec.Fields("NetValue").value))
   Me.TxtFATYou.Text = IIf(IsNull(RsSavRec.Fields("FATYou").value), 0, (RsSavRec.Fields("FATYou").value))
   Me.TxtFATYou22.Text = IIf(IsNull(RsSavRec.Fields("FATYou22").value), 0, (RsSavRec.Fields("FATYou22").value))
   Me.TxtFATYou2.Text = IIf(IsNull(RsSavRec.Fields("FATYou2").value), 0, (RsSavRec.Fields("FATYou2").value))
   Me.TxtFATValue.Text = IIf(IsNull(RsSavRec.Fields("FATValue").value), 0, (RsSavRec.Fields("FATValue").value))
   Me.TxtTotalValue.Text = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, (RsSavRec.Fields("TotalValue").value))
   If Not IsNull(RsSavRec.Fields("ComResid").value) Then
   If RsSavRec.Fields("ComResid").value = 1 Then
   ComResid(1).value = True
   Else
   ComResid(0).value = True
   End If
   Else
   ComResid(0).value = True
   End If
'*********************************
    Contract_period_no.Text = IIf(IsNull(RsSavRec("Contract_period_no").value), 0, RsSavRec("Contract_period_no").value)
 
    If IsNull(RsSavRec("Contract_period").value) Then
        Me.Contract_period.ListIndex = 0
    Else
        Me.Contract_period.ListIndex = RsSavRec("Contract_period").value
    End If

'*********************************

  FillGridWithData
  RetriveOldPayment
  ReLineGrid True

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

  '  With Grid

  '      For i = 1 To .Rows - 1
'
'            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("CityID")) Then
'                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
'                .Row = i
'                Exit Sub
'            End If

'        Next

'    End With
Dim Percetage2 As Double
Dim TotalService  As Double
Subvat = 0
TotalService = val(TxtWater.Text) + val(TxtElectricity.Text) + val(TxtPhone.Text)

PercentgValueAddedAccount_Transec StrDate.value, 21, 1, , Percetage2
'  AddTovatValue = AddTovatValue - val(TxtTotalContract.Text) - val(TxtDiscountValue.Text)
'AddTovatValue = AddTovatValue - val(TxtTotalContract.Text) - val(TxtDiscountValue.Text)
           If WaterElecValueInVAT.value = Checked Then
                    Subvat = Subvat + val(TotalService) * Percetage2 / 100
                End If
                If CommiValueInVAT.value = vbChecked Then
                Subvat = Subvat + val(TxtCommiValue) * Percetage2 / 100
                   
                  End If
                If InsurValueInVAT.value = vbChecked Then
                Subvat = Subvat + val(TxtInsuranceValue) * Percetage2 / 100
                 
                   End If
ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

'Private Sub Grid_EnterCell()
'    On Error GoTo ErrTrap
'    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("CityID")))
'ErrTrap:
'End Sub



Public Function FindRec(ByVal RecId As Long, Optional ByVal iSFromSearch As Boolean = False, Optional NoteSerial1 As String)
    On Error GoTo ErrTrap
    If RecId = 0 Then Exit Function
    Dim My_SQL As String
    If iSFromSearch Then
           My_SQL = " select * from TblContract "
        If SystemOptions.usertype = UserAdminAll Then
            My_SQL = My_SQL & " where   1<>-1"
        Else
       '     My_SQL = My_SQL & " where   Branch_NO=" & Current_branch
        End If

'        If RereivID <> 0 Then
'            My_SQL = My_SQL & "  and ContNo=" & RereivID & ""
'        End If
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    End If
    If NoteSerial1 = "" Then
     RsSavRec.find "ContNo=" & RecId, , adSearchForward, 1
    Else
    RsSavRec.find "noteserial1='" & NoteSerial1 & "'", , adSearchForward, 1
    
    End If
    If Not (RsSavRec.EOF) Then
        FiLLTXT
    Else
    RsSavRec.MoveFirst
      FiLLTXT
    End If
  
    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

Private Sub RemoveGridRow2()

    With Me.UnitsGrid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow()

    With Me.VSFlexGrid2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub FristPaymentDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         FirstInstallDateH.value = ToHijriDate(FristPaymentDate.value)
       
End If
End Sub

Private Sub FristPaymentDate_GotFocus()
hijriorJerojian = 1
End Sub

Private Sub Fromdateh_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    StrDate.value = ToGregorianDate(FromDateH.value)
       FirstInstallDateH.value = FromDateH.value
          FristPaymentDate.value = ToGregorianDate(FirstInstallDateH.value)
          ClculteVAT
       hijriorJerojian = 0
       CalcContractIntervalAuto
End If
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "NetWater"
.TextMatrix(Row, .ColIndex("Water")) = .TextMatrix(Row, .ColIndex("NetWater"))
Case "NetElectric"
.TextMatrix(Row, .ColIndex("Electric")) = .TextMatrix(Row, .ColIndex("NetElectric"))
End Select
End With

ReLineGrid
End Sub
Function ChecStopeCustomer(CusID As Double) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     CusID"
sql = sql & " From dbo.TblCustemers"
sql = sql & " WHERE     (locked = 1) AND (CusID = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic
If rs2.RecordCount > 0 Then
ChecStopeCustomer = True
Else
ChecStopeCustomer = False
End If
End Function
Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

'   If (ChkRenew Or checkContractTransactions(val(TxtContNo.Text))) And mchkAllowEditPaymentCont Then
'        mCanEdit = True
'
'    Else
'        mCanEdit = False
'    End If
'
'    If ChkRenew.value = vbChecked And Not mchkAllowEditPaymentCont Then
'        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
'        Exit Sub
'    End If
'
'    If checkContractTransactions(val(TxtContNo.Text)) = True And Not mchkAllowEditPaymentCont Then
'        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰  ⁄œÌ·…", vbCritical
'        Exit Sub
'
'    End If

    If (Me.TxtModFlg.Text = "R" And GridInstallments.ColKey(Col) <> "PrintJE" And GridInstallments.ColKey(Col) <> "Print" And GridInstallments.ColKey(Col) <> "RecalcVAt") Then
        If Not mchkAllowEditPaymentCont Then
            Cancel = True
        End If
    Else

    
    
    With GridInstallments
 If (Opt(4).value = True Or Opt(3).value = True) And .ColKey(Col) <> "Print" Then
 Cancel = True
 ElseIf GridInstallments.ColKey(Col) = "RecalcVAt" And val(TxtFATValue.Text) <> 0 Then
 Cancel = True
 Else
 Cancel = False
 End If
 
    '     If .ColKey(Col) <> "Status" And .ColKey(Col) <> "TelandNet" And .ColKey(Col) <> "Insurance" Then
   
   '      If Opt(0).value = True Then Cancel = True: Exit Sub
   '     Cancel = True
   '
   '     End If
 
        
    End With
  End If
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempCustomerCodeInsuranceAccount  As String
    Dim StrTempAccountCode2 As String
    Dim Msg2 As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim VAtComm As Double
    Dim StrSQL As String
    
    
        
Dim mVATValue1Com  As Double
Dim mVATValue2Com  As Double
Dim Commission2 As Double

    Iqar = val(DcbIqara.BoundText)
        commisiontype = AqarCommisionType(Iqar, AmolaValues, ownerid)
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

 
'        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
'GoTo ll
            
   
            StrTempDes = "⁄ﬁœ «ÌÃ«— —ﬁ„   " & TxtNoteSerial1 & CHR(13) & "  ··„” √Ã—   " & dcCustomer.Text & CHR(13) & " Ê«·„«·ﬂ " & dcsupplier.Text & CHR(13)
            StrTempDes = StrTempDes & "   «·⁄ﬁ«— " & DcbIqara.Text & CHR(13)
             StrTempDes = StrTempDes & "   «·ÊÕœ…  " & DcbUnitType.Text & " —ﬁ„ " & DcbUnitNo.Text & CHR(13)
             
      StrTempDes = StrTempDes & "     »œ«Ì… «·⁄ﬁœ „‰  " & FromDateH.value & CHR(13) & " «·Ì " & todateH.value & CHR(13)
        StrTempDes = StrTempDes & " «·„Ê«›ﬁ " & StrDate.value & CHR(13) & " «·Ì " & EndDate.value & CHR(13)

      StrTempDes = StrTempDes & CHR(13) & balanceDes.Text
       
  
            LngDevNO = LngDevNO + 1
'Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)
Notevalue = 0
If SystemOptions.WorkWithFirstInstallOnly = False Then
Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)

Else

With GridInstallments


Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String
Dim mCountDay1  As Double
Dim mCountDay2  As Double
Dim mVATYou1 As Double
Dim mVATYou2 As Double
newinstallNo = (.TextMatrix(1, .ColIndex("InstallNo")))  ' val(.TextMatrix(Row + 1, .ColIndex("InstallNo")))
newinstallNo = newinstallNo + 1
getnextDate newinstallNo, nextinstalldate, nextinstalldateH
If year(nextinstalldate) < 1900 Then
nextinstalldate = Time
End If

StrTempDes = StrTempDes & " «À»«  «” Õﬁ«ﬁ «·œ›⁄Â «·«Ê·Ï  «· Ì  »œ√ » «—ÌŒ " & CHR(13) & (.TextMatrix(1, .ColIndex("Due_DateH"))) & CHR(13) & "  «·Ì    " & nextinstalldateH
StrTempDes = StrTempDes & "«·„Ê«›ﬁ „Ì·«œÌ „‰ " & CHR(13) & (.TextMatrix(1, .ColIndex("Due_Date"))) & CHR(13) & "  «·Ì    " & nextinstalldate & CHR(13)
If .Rows > 1 Then

Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("RentValue")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Commissions")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Insurance")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Water")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("Electric")))
Notevalue = Notevalue + val(.TextMatrix(1, .ColIndex("TelandNet")))
mVATValue1Com = val(.TextMatrix(1, .ColIndex("VATValue1Com")))
mVATValue2Com = val(.TextMatrix(1, .ColIndex("VATValue2Com")))
'Commission2 = val(.TextMatrix(1, .ColIndex("Commissions2")))
mCountDay1 = val(.TextMatrix(1, .ColIndex("CountDay1")))
mCountDay2 = val(.TextMatrix(1, .ColIndex("CountDay2")))

mVATYou1 = val(.TextMatrix(1, .ColIndex("VATYou1")))
mVATYou2 = val(.TextMatrix(1, .ColIndex("VATYou2")))
 
End If

End With
End If
            
'll:
   LngDevNO = 0
 '           If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchId) = False Then
 '               GoTo ErrTrap
 '           End If
 
 
 If val(TxtTotalContract.Text) > 0 Then
       
       
       If SystemOptions.WorkWithFirstInstallOnly = False Then
          Notevalue = val(TxtTotalContract.Text)
     Else
       Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("RentValue")))
     End If
     Dim val2 As Double
        
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
        If SystemOptions.WorkWithFirstInstallOnly = False Then
        val2 = val(TxtFATValue.Text)
        Else
        val2 = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("VATValue")))
        End If
        
        


If SystemOptions.OpenVATAccountOwner = True And commisiontype = 1 Then
AccountVat.BoundText = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_VAT")

End If



          If val(TxtFATValue.Text) > 0 Then
            LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val2, 0, "      «·ﬁÌ„… «·„÷«›…  " & StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
              LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVat.BoundText, val2, 1, "      «·ﬁÌ„… «·„÷«›…  " & StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
        End If
        If SystemOptions.InsuranceOnOwner = True And commisiontype = 0 Then
        '  Notevalue = Notevalue + val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("TelandNet")))
          Else
          Notevalue = Notevalue '+ val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("TelandNet")))
          End If
          
              LngDevNO = LngDevNO + 1
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             


  If commisiontype = 0 Then
   StrTempAccountCode = Account_Code_dynamic80
   Else
   StrTempAccountCode = Account_Code_dynamic123 '··€Ì—

   End If
   
   If AmolaValues = 0 Then
      
      LngDevNO = LngDevNO + 1
  If SystemOptions.Create2account4Supp = True And commisiontype <> 0 Then
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
                If StrTempAccountCode = "" Then
                            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code")
                End If
                
                 If StrTempAccountCode = "" Then
                            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "accountaccountaqar")
                End If
                
                
     End If
     
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «· ⁄«ﬁœ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
   Else 'commissions
            Dim mCountDaysTotal As Double
            Dim mCostDay As Double
            Dim commission As Double
            Dim mVATValue2 As Double
'            Dim mCountDay1 As Double
            'Dim mCountDay2 As Double
            Dim mVatPercent  As Double
            Dim mVatPercent2 As Double
            Dim mVATValue1 As Double
            Dim mPecr1 As Double
            Dim mPecr2 As Double
             txtDateK.value = CDate("2020-05-14")
            txtDateK2.value = CDate("2020-06-30")
            If commisiontype = 1 Then
                commission = Notevalue * AmolaValues / 100
                LngDevNO = LngDevNO + 1
                If SystemOptions.CommissionDue = True Then
                    Notevalue = Notevalue - commission
                    If ComResid(1).value = True Then
                        VAtComm = commission * val(TxtFATYou2.Text) / 100
                    Else
                        VAtComm = 0
                    End If
                End If
            Else
                     ' commission = Notevalue * AmolaValues / 100
                LngDevNO = LngDevNO + 1
                Notevalue = Notevalue
            End If
            

   
   
  'ownerid
 If SystemOptions.Create2account4Supp = True Then
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
     End If

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
            LngDevNO = LngDevNO + 1
            
   StrTempAccountCode = Account_Code_dynamic125 '⁄„Ê·«  „” Õﬁ…
     If commisiontype = 1 And SystemOptions.CommissionDue = True Then
'StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, commission, 1, StrTempDes & "       ﬁÌ„… «·⁄„Ê·Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
         If mVATValue1Com > 0 Then
            LngDevNO = LngDevNO + 1
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, mVATValue1Com, 0, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ ⁄·Ï ⁄œœ «Ì«„ " & mCountDay1 & " »‰”»… " & mVatPercent, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
            LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVat2.BoundText, mVATValue1Com, 1, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ ⁄·Ï ⁄œœ «Ì«„ " & mCountDay1 & " »‰”»… " & mVatPercent, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
        End If
            
            
       If mVATValue2Com > 0 Then
            LngDevNO = LngDevNO + 1
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, mVATValue2Com, 0, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ ⁄·Ï ⁄œœ «Ì«„ " & mCountDay2 & " »‰”»… " & mVatPercent2, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
            LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVat2.BoundText, mVATValue2Com, 1, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ ⁄·Ï ⁄œœ «Ì«„ " & mCountDay2 & " »‰”»… " & mVatPercent2, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
             
            End If
        End If
         End If
    End If
            
  End If
  
  
 If (val(TxtCommiValue.Text)) > 0 Then

       
              If SystemOptions.WorkWithFirstInstallOnly = False Then
             Notevalue = (val(TxtCommiValue.Text))
     Else
     Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Commissions")))
     End If
     

   
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
       If SystemOptions.DueComm = False Then
          Msg2 = "⁄„Ê·«  Ê—”Ê„ «œ«—Ì…"
          StrTempAccountCode2 = Account_Code_dynamic81
      Else
          Msg2 = "«” Õﬁ«ﬁ «·”⁄Ì"
          StrTempAccountCode2 = get_account_code_branch(153, val(Dcbranch.BoundText))
        
  End If
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & Msg2, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If

   
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode2, Notevalue, 1, StrTempDes & Msg2, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
   If val(TxtInsuranceValue.Text) > 0 Then
       
               StrTempCustomerCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   StrTempCustomerCodeInsuranceAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText), "InsuranceAccount")
'StrTempAccountCode = Account_Code_dynamic82

              If SystemOptions.WorkWithFirstInstallOnly = False Then
      Notevalue = val(TxtInsuranceValue.Text)
     Else
     Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Insurance")))
     End If
            
      LngDevNO = LngDevNO + 1
           
                 If SystemOptions.CreateInsuranceAccountForCustomers Then
    
 StrTempAccountCode = StrTempCustomerCodeInsuranceAccount
  
 Else
 StrTempAccountCode = Account_Code_dynamic82
  End If
       
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempCustomerCode, Notevalue, 0, StrTempDes & "     √„Ì‰ „” —œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  
If commisiontype = 1 And SystemOptions.InsuranceOnOwner = True Then
StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")

End If
    If StrTempAccountCode = "" Then
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code")
    End If
    
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "     √„Ì‰ „” —œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
            
  End If
  
  
     If val(TxtWater.Text) > 0 Then
       
'
   
              If SystemOptions.WorkWithFirstInstallOnly = False Then
    Notevalue = val(TxtWater.Text)
     Else
     Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Water")))
     End If
     
           LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "    „Ì«Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

StrTempAccountCode = Account_Code_dynamic83 ''

 If SystemOptions.DueWater = True Then '«” Õﬁ«ﬁ «·„Ì«Â
 StrTempAccountCode = Account_Code_dynamic154
 End If
If commisiontype = 1 And SystemOptions.InsuranceOnOwner = True Then
StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
End If

    If StrTempAccountCode = "" Then
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code")
    End If


   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "    „Ì«Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
       If val(TxtElectricity.Text) > 0 Then
       
     '  Notevalue = val(TxtElectricity.text)
   
                If SystemOptions.WorkWithFirstInstallOnly = False Then
    Notevalue = val(TxtElectricity.Text)
     Else
     Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("Electric")))
     End If
     
     
             LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "      ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  
StrTempAccountCode = Account_Code_dynamic84


 If SystemOptions.DueElectr = True Then  '«” Õﬁ«ﬁ «·ﬂÂ—»«¡
 StrTempAccountCode = Account_Code_dynamic155 ''
 End If
 
 
If commisiontype = 1 And SystemOptions.InsuranceOnOwner = True Then
StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
End If

    If StrTempAccountCode = "" Then
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code")
    End If


   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "      ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
  
  
       If (val(TxtPhone.Text)) > 0 Then
       
'       Notevalue = (val(TxtPhone.text) + val(TxtEnternet.text))
   
                If SystemOptions.WorkWithFirstInstallOnly = False Then
    Notevalue = val(TxtPhone.Text)
     Else
     Notevalue = val(GridInstallments.TextMatrix(1, GridInstallments.ColIndex("TelandNet")))
     End If
     
             LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   Dim StrTempDes1 As String
With GridInstallments
         StrTempDes1 = "⁄ﬁœ «ÌÃ«— —ﬁ„    " & TxtNoteSerial1 & "  ··„” √Ã—   " & dcCustomer.Text & " Ê«·„«·ﬂ " & dcsupplier.Text
          StrTempDes = StrTempDes & "   «·› —… „‰  " & FromDateH.value & " «·Ì " & todateH.value
StrTempDes = StrTempDes & " «·„Ê«›ﬁ " & StrDate.value & " «·Ì " & EndDate.value


StrTempDes1 = StrTempDes & " «À»«  «” Õﬁ«ﬁ «·œ›⁄Â «·«Ê·Ï  «· Ì  »œ√ » «—ÌŒ " & (.TextMatrix(1, .ColIndex("Due_DateH")))
   End With
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes1 & "    Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

StrTempAccountCode = Account_Code_dynamic85


 If SystemOptions.DueService = True Then   '«” Õﬁ«ﬁ «·Œœ„« 
 StrTempAccountCode = Account_Code_dynamic156 ''
 End If
 
 
 
 
If commisiontype = 1 And SystemOptions.InsuranceOnOwner = True Then
StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
End If

    If StrTempAccountCode = "" Then
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code")
    End If

   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "    Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
 '''///////////////////////////
  If (val(TxtNotVal.Text)) > 0 Then
  Notevalue = (val(TxtNotVal.Text))
        LngDevNO = LngDevNO + 1
        
               StrTempAccountCode = get_account_code_branch(95, my_branch)
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "      ⁄—»Ê‰ «ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
     
     '  StrTempAccountCode = Account_Code_dynamic81
     StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "      ⁄—»Ê‰ «ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  End If
    updateNotesValueAndNobytext CDbl(general_noteid)
ErrTrap:
End Function
Function GetMaxLin(Optional general_noteid As Long) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select max(DEV_ID_Line_No) as DEV_ID_Line_No From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxLin = IIf(IsNull(rs2("DEV_ID_Line_No").value), 0, rs2("DEV_ID_Line_No").value)
Else
GetMaxLin = 0
End If
End Function
Sub DeleteJE()
Dim i As Integer
Dim StrSQL As String
With GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("NoteId"))) <> 0 Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(i, .ColIndex("NoteId"))) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "delete From Notes where NoteID =" & val(.TextMatrix(i, .ColIndex("NoteId"))) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute "Update TblContractInstallments set NoteIdDiff=Null,NoteSerialDiff=null, NoteID=null ,NoteSerial=null where id=" & val(.TextMatrix(i, .ColIndex("Installid"))) & " "

        StrSQL = "delete From Notes where NoteID =" & val(.TextMatrix(i, .ColIndex("NoteIdDiff"))) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords

FindRec val(Me.TxtContNo.Text)
End If
Next i
End With

End Sub
Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional Row As Long)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Variant
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim StrSQL As String
    Dim account As String
     PercentgValueAddedAccount_Transec NoteDate, 8, 1, account
            AccountVat.BoundText = account
       If AccountVat.BoundText = "" Then
       MsgBox "Ì—ÃÏ  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
       Exit Function
       End If
       
       
      With GridInstallments
      Notevalue = val(.TextMatrix(Row, .ColIndex("VATValue")))
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
         If Notevalue > 0 Then
         LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

          LngDevNO = GetMaxLin(general_noteid) + 1
    '«·ÿ—› «·„Ì‰
         my_branch = BranchID


         StrTempDes = "—ﬁ„ «·œ›⁄…" & "  " & .TextMatrix(Row, .ColIndex("InstallNo"))
         StrTempDes = StrTempDes & "⁄ﬁœ «ÌÃ«— —ﬁ„    " & TxtNoteSerial1 & "  ··„” √Ã—   " & dcCustomer.Text & " Ê«·„«·ﬂ " & dcsupplier.Text
          StrTempDes = StrTempDes & "   «·› —… „‰  " & FromDateH.value & " «·Ì " & todateH.value
StrTempDes = StrTempDes & " «·„Ê«›ﬁ " & StrDate.value & " «·Ì " & EndDate.value

            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
            
        
        
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "      Õ”«» «·„” «Ã— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
            
             If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVat.BoundText, Notevalue, 1, StrTempDes & "     Õ”«» «·ﬁÌ„… «·„÷«›… ·⁄ﬁÊœ «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         End If

    
    End With
ErrTrap:
End Function

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Dim newinstallNo  As Double
Dim nextinstalldate As Date
Dim nextinstalldateH As String

newinstallNo = 0
With GridInstallments
Select Case .ColKey(Col)
Case "Print"
newinstallNo = val(.TextMatrix(Row + 1, .ColIndex("InstallNo")))
getnextDate newinstallNo, nextinstalldate, nextinstalldateH
PeintInstalMent val(.TextMatrix(Row, .ColIndex("InstallNo"))), nextinstalldate, nextinstalldateH

Case "PrintJE"
ShowGL_cc .TextMatrix(Row, .ColIndex("NoteSerial")), , 200
Case "RecalcVAt"
RecalcVAt Row
createVoucher2 (Row)
MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
End Select
End With
End Sub
Private Sub Contract_period_no_Change()
CalcContractIntervalAuto

End Sub
Function CalcContractIntervalAuto()
If Me.TxtModFlg = "R" Or val(Contract_period_no.Text) = 0 Then Exit Function
If RdRTypeDate(0).value = True Then 'ÂÃ—Ì
  VBA.Calendar = vbCalHijri
 todateH.value = calcenaddate(FromDateH.value, val(Contract_period_no.Text), val(Contract_period.ListIndex))
 
       VBA.Calendar = vbCalGreg
    EndDate.value = ToGregorianDate(todateH.value)
       hijriorJerojian = 0
  Else '„Ì·«œÌ
  
  EndDate.value = calcenaddate(StrDate.value, val(Contract_period_no.Text), val(Contract_period.ListIndex))
  
         todateH.value = ToHijriDate(EndDate.value)
       hijriorJerojian = 1
End If

End Function
Private Sub Contract_period_no_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Contract_period_no.Text, 0)
End Sub

Function RecalcVAt(rowno As Long)
 'GridInstallments VATValue

Dim Percetage As Double
Dim account As String
Dim StrDate As Date
Dim i As Integer
i = rowno

If ComResid(1).value = True Then
With GridInstallments
StrDate = .TextMatrix(i, .ColIndex("Due_Date"))
End With
PercentgValueAddedAccount_Transec StrDate, 8, 1, account, Percetage
'TxtFATYou.Text = Percetage
'AccountVat.BoundText = account
Else
Exit Function
 End If
Dim strasstring As String
  With GridInstallments
  '+ val(.TextMatrix(i, .ColIndex("Insurance")))
  





If InsurValueInVAT.value = vbChecked Then
     .TextMatrix(i, .ColIndex("VATValue")) = Percetage / 100 * (val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Insurance"))))
     .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))) + val(.TextMatrix(i, .ColIndex("Insurance")))
   Else
   .TextMatrix(i, .ColIndex("VATValue")) = Percetage / 100 * (val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet"))))
     .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet")))
   End If
   
   .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
   
     strasstring = "update  TblContractInstallments set  VATValue=" & val(.TextMatrix(i, .ColIndex("VATValue"))) & ",installValue= " & val(.TextMatrix(i, .ColIndex("Value")))
    strasstring = strasstring & ",Remains=" & val(.TextMatrix(i, .ColIndex("Remains")))
     
     strasstring = strasstring & " where id=" & val(.TextMatrix(i, .ColIndex("Installid")))
     Cn.Execute strasstring
RsSavRec.Resync adAffectCurrent

   End With
   

End Function
Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "Print"
.ColComboList(.ColIndex("Print")) = "..."

Case "RecalcVAt"
.ColComboList(.ColIndex("RecalcVAt")) = "..."
Case "PrintJE"
.ColComboList(.ColIndex("PrintJE")) = "..."

End Select
End With
End Sub
Function getnextDate(Optional newinstallNo As Double, Optional ByRef installdate, Optional ByRef installdateh)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    
    MySQL = " SELECT    installdate,installdateH "
     
    MySQL = MySQL & "      FROM         dbo.TblContract LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
    MySQL = MySQL & "        Where (dbo.TblContract.ContNo = " & val(TxtContNo.Text) & ") And (dbo.TblContractInstallments.InstallNo =" & newinstallNo & ")"
   Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
    Else
    installdate = IIf(IsNull(RsData("installdate").value), Null, RsData("installdate").value)
    installdateh = IIf(IsNull(RsData("installdateH").value), Null, RsData("installdateH").value)
    
    
    End If
    
End Function
 
Function PeintInstalMent(Optional InstallNo As Double, Optional nextinstalldate As Date, Optional nextinstalldateH As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    
    MySQL = " SELECT  '" & nextinstalldate & "' as installdateNext ,'" & nextinstalldateH & "' as installdateHNext,   dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.ownerid, dbo.TblContract.UnitType, "
    MySQL = MySQL & "                   dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.unitno, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    MySQL = MySQL & "                   dbo.TblCustemers.Fullcode, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.TodateH,"
    MySQL = MySQL & "                   dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblContract.NoteSerial1 AS HNoteSerial1,"
    MySQL = MySQL & "                   dbo.TblContract.NetValue, dbo.TblContract.FATYou, dbo.TblContract.FATValue, dbo.TblContract.TotalValue, dbo.TblContractInstallments.*,"
    MySQL = MySQL & "                   dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Address, dbo.TblCustemers.E_mail,"
    MySQL = MySQL & "                   dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Remark2, dbo.TblCustemers.CustGID, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTitle,"
    MySQL = MySQL & "                   dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2,"
    MySQL = MySQL & "                   dbo.TblCustemers.BoxMil , dbo.TblCustemers.VATNO"
    MySQL = MySQL & "      FROM         dbo.TblContract LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
    MySQL = MySQL & "        Where (dbo.TblContract.ContNo = " & val(TxtContNo.Text) & ") And (dbo.TblContractInstallments.InstallNo =" & InstallNo & ")"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBiilRent.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBiilRent.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        StrReportTitle = ""
    End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub Image1_Click()
Label400.Visible = False
saveinstdetailforpart2
Label400.Visible = True

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton3_Click()
  Load FrmNotesSearch
           FrmNotesSearch.SearchType = 7
            FrmNotesSearch.show vbModal
End Sub

Private Sub Optx_Click(Index As Integer)
'On Error Resume Next
Dim My_SQL As String
RsSavRec.Close

Select Case Index

Case 0
 If SystemOptions.usertype = UserAdminAll Then
 My_SQL = " select * from TblContract "
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   Else
   
    My_SQL = " select * from TblContract where Branch_NO=" & Current_branch
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 End If

Case 1

  My_SQL = " select * from TblContract where Iqar=" & val(DcbIqara.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
Case 2

  My_SQL = " select * from TblContract where ownerid=" & val(dcsupplier.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
Case 3

  My_SQL = " select * from TblContract where CusID=" & val(dcCustomer.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
Case 4

  My_SQL = " select * from TblContract where Emp_ID=" & val(DcboEmp.BoundText)
      RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      
End Select
BtnFirst_Click
End Sub

Private Sub RdRTypeDate_Click(Index As Integer)
datetype
CalcContractIntervalAuto
End Sub

Private Sub RecordDateH_LostFocus()
     If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            ContDate.value = ToGregorianDate(RecorddateH.value)
            datetype
     If ChekSanNumber(Current_branch, 60) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial.Text = ""
     End If
End Sub
Sub datetype()
If Me.TxtModFlg = "R" Then Exit Sub
If RdRTypeDate(0).value = True Then
StrDate.value = ContDate.value
FristPaymentDate.value = ContDate.value
FirstInstallDateH.value = (RecorddateH.value)
FromDateH.value = RecorddateH.value
 hijriorJerojian = 0
Else

StrDate.value = (ContDate.value)
hijriorJerojian = 1
FristPaymentDate.value = (ContDate.value)
FirstInstallDateH.value = RecorddateH.value
FromDateH.value = RecorddateH.value
End If
End Sub
Private Sub StrDate_Change()
'If Me.TxtModFlg.Text <> "R" Then
         FromDateH.value = ToHijriDate(StrDate.value)
       
       ClculteVAT
       CalcContractIntervalAuto
'End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text1.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 1215
        FrmCustemerSearch.show vbModal

    End If
 

If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub ToDateH_LostFocus()
'If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    EndDate.value = ToGregorianDate(todateH.value)
       hijriorJerojian = 0
'End If
End Sub

Private Sub TxtCommiValue_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
Calculte
End If
End Sub

Private Sub TxtContNo_Change()
  Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Private Sub TxtElectricity_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
Calculte
End If

End Sub

Private Sub TxtEmployeeID_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.Text, True)
End If

End Sub

Private Sub TxtInsuranceValue_Change()
Calculte
End Sub

Private Sub TxtMeterCount_Change()
    If Me.TxtModFlg <> "R" Then
    ReLineGrid
'    TxtTotalContract.Text = val(TxtTotalContract.Text) + (val(TxtMeterValue) * val(TxtMeterCount))
    
    End If

End Sub

Private Sub TxtMeterValue_Change()
    If Me.TxtModFlg <> "R" Then
    ReLineGrid
   ' TxtTotalContract.Text = val(TxtTotalContract.Text) + (val(TxtMeterValue) * val(TxtMeterCount))
    
    End If

End Sub

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
    Ele(15).Enabled = True
    Ele(18).Enabled = True
    Ele(0).Enabled = True
    
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
       ' Command12.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Ele(15).Enabled = False
    Ele(18).Enabled = False
       Ele(0).Enabled = True
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = True
        btnDelete.Enabled = False

        If TxtContNo.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
'        Command12.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.Text = "E" Then
        Ele(15).Enabled = True
    Ele(18).Enabled = True
    Ele(0).Enabled = True
    
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
'        Command12.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

  '  On Error GoTo ErrTrap
Dim ActulaPyaed As Double
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    ActulaPyaed = 0
 My_SQL = " SELECT     dbo.TblIqrMerg.ID, dbo.TblIqrMerg.Cont, dbo.TblIqrMerg.RentType, dbo.TblIqrMerg.Price, dbo.TblIqrMerg.Area, dbo.TblIqrMerg.Remark, dbo.TblIqrMerg.UntID,"
 My_SQL = My_SQL & "                      dbo.TblAqarDetai.unitno , dbo.TblIqrMerg.typeid, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee"
My_SQL = My_SQL & " FROM         dbo.TblIqrMerg LEFT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.TblAkarUnit ON dbo.TblIqrMerg.TypeID = dbo.TblAkarUnit.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqarDetai ON dbo.TblIqrMerg.UntID = dbo.TblAqarDetai.Id"
My_SQL = My_SQL & " Where (dbo.TblIqrMerg.cont = " & val(Me.TxtContNo.Text) & ")"
rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
''/
  With Me.UnitsGrid
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst
 
            For i = 1 To .Rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
                 If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
            End If

              .TextMatrix(i, .ColIndex("unittype")) = val(IIf(IsNull(rs.Fields("TypeID").value), "", rs.Fields("TypeID").value))
               .TextMatrix(i, .ColIndex("unitno")) = IIf(IsNull(rs.Fields("unitno").value), "", rs.Fields("unitno").value)
              .TextMatrix(i, .ColIndex("id")) = val(IIf(IsNull(rs.Fields("UntID").value), "", rs.Fields("UntID").value))
              
                  .TextMatrix(i, .ColIndex("length")) = val(IIf(IsNull(rs.Fields("Area").value), "", rs.Fields("Area").value))
              .TextMatrix(i, .ColIndex("namerentType")) = val(IIf(IsNull(rs.Fields("RentType").value), "", rs.Fields("RentType").value))
              
 .TextMatrix(i, .ColIndex("meterPrice")) = val(IIf(IsNull(rs.Fields("Price").value), "", rs.Fields("Price").value))
  .TextMatrix(i, .ColIndex("RentValue")) = val(.TextMatrix(i, .ColIndex("meterPrice"))) * val(.TextMatrix(i, .ColIndex("length")))
    .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs.Fields("Remark").value), "", rs.Fields("Remark").value)
        rs.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With

   rs.Close
   ''////contractsales

    Set rs = New ADODB.Recordset
My_SQL = "SELECT     dbo.TblCOntractSales.ContNo, dbo.TblCOntractSales.ID, dbo.TblCOntractSales.rate, dbo.TblCOntractSales.EmpID, dbo.TblEmployee.Emp_Name, "
My_SQL = My_SQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblCOntractSales.idd, dbo.TblCOntractSales.GroupID, dbo.TBLSalesRepGroups.name,"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups.NameE"
My_SQL = My_SQL & " FROM         dbo.TblCOntractSales LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups ON dbo.TblCOntractSales.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.TblCOntractSales.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & " Where (dbo.TblCOntractSales.ContNo =" & val(Me.TxtContNo.Text) & ")"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.VSFlexGrid2
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
   If SystemOptions.UserInterface = EnglishInterface Then
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
      Else
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
 
    End If
    .TextMatrix(i, .ColIndex("groupid")) = val(IIf(IsNull(rs.Fields("GroupID").value), "", rs.Fields("GroupID").value))
    .TextMatrix(i, .ColIndex("rate")) = val(IIf(IsNull(rs.Fields("rate").value), "", rs.Fields("rate").value))
    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
   .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(rs.Fields("idd").value), "", rs.Fields("idd").value)
        rs.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With
''///
''//

    Set rs = New ADODB.Recordset
My_SQL = " SELECT  * from TblContractDet"
My_SQL = My_SQL & " WHERE     (ContNo =" & val(Me.TxtContNo.Text) & ")"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.VSFlexGrid1
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i

                .TextMatrix(i, .ColIndex("Count")) = val(IIf(IsNull(rs.Fields("Count").value), "", rs.Fields("Count").value))
 .TextMatrix(i, .ColIndex("Code")) = val(IIf(IsNull(rs.Fields("Code").value), "", rs.Fields("Code").value))
  .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs.Fields("Des").value), "", rs.Fields("Des").value)
        rs.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With
    
    

   rs.Close



Dim notpayed As Double
notpayed = 0
 
My_SQL = " SELECT  * from TblContractInstallments"
My_SQL = My_SQL & " WHERE     (ContNo =" & val(Me.TxtContNo.Text) & ")  order by InstallNo"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
           .TextMatrix(i, .ColIndex("DevID")) = (IIf(IsNull(rs.Fields("DevID").value), 0, rs.Fields("DevID").value))
          .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
          .TextMatrix(i, .ColIndex("TempInstal")) = (IIf(IsNull(rs.Fields("TempInstal").value), 0, rs.Fields("TempInstal").value))
          .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
          .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 1, rs.Fields("hijri").value))
          .TextMatrix(i, .ColIndex("DES")) = (IIf(IsNull(rs.Fields("DES").value), "", rs.Fields("DES").value))
          .TextMatrix(i, .ColIndex("Due_DateH")) = Format((IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value)), "yyyy/MM/dd")
           .TextMatrix(i, .ColIndex("Due_Date")) = Format(IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value), "yyyy/MM/dd")
           .TextMatrix(i, .ColIndex("NoteSerialDiff")) = (IIf(IsNull(rs.Fields("NoteSerialDiff").value), "", rs.Fields("NoteSerialDiff").value))
           
           .TextMatrix(i, .ColIndex("NoteIdDiff")) = (IIf(IsNull(rs.Fields("NoteIdDiff").value), "", rs.Fields("NoteIdDiff").value))
           
        .TextMatrix(i, .ColIndex("CountDay1")) = rs!CountDay1 & ""
        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
        
'        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
'        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
'        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
        
        .TextMatrix(i, .ColIndex("VATYou1")) = rs!VATYou1 & ""
        .TextMatrix(i, .ColIndex("VATYou2")) = rs!VATYou2 & ""
        
       
  
  
    
            .TextMatrix(i, .ColIndex("VATValue1")) = rs!VATValue1 & ""
           .TextMatrix(i, .ColIndex("VATValue2")) = rs!VATValue2 & ""
      
'                   .TextMatrix(i, .ColIndex("CountDay1Com")) = rs!CountDay1Com & ""
'        .TextMatrix(i, .ColIndex("CountDay2Com")) = rs!CountDay2Com & ""
'
''        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
''        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
''        .TextMatrix(i, .ColIndex("CountDay2")) = rs!CountDay2 & ""
'
'        .TextMatrix(i, .ColIndex("VATYou1Com")) = rs!VATYou1Com & ""
'        .TextMatrix(i, .ColIndex("VATYou2Com")) = rs!VATYou2Com & ""
'
'
            .TextMatrix(i, .ColIndex("VATValue1Com")) = rs!VATValue1Com & ""
           .TextMatrix(i, .ColIndex("VATValue2Com")) = rs!VATValue2Com & ""
'
'
'
        'yyyy/MM/dd
       .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
       .TextMatrix(i, .ColIndex("ServiceArbon")) = (IIf(IsNull(rs.Fields("ServiceArbon").value), 0, rs.Fields("ServiceArbon").value))
       .TextMatrix(i, .ColIndex("NoteSerial")) = (IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value))
       .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
       .TextMatrix(i, .ColIndex("NoteId")) = (IIf(IsNull(rs.Fields("NoteId").value), "", rs.Fields("NoteId").value))
If Not IsNull(rs.Fields("Status").value) Then
             If rs.Fields("Status").value = 0 Then
                    .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked
            Else
                     .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked
                       notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
            End If

End If

    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("VATPayed")) = (IIf(IsNull(rs.Fields("VATPayed").value), 0, rs.Fields("VATPayed").value))
    .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs.Fields("VATValue").value), 0, rs.Fields("VATValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
 .TextMatrix(i, .ColIndex("NpayedValue")) = (IIf(IsNull(rs.Fields("NpayedValue").value), 0, rs.Fields("NpayedValue").value))
        
    .TextMatrix(i, .ColIndex("OldValue")) = (IIf(IsNull(rs.Fields("OldValue").value), 0, rs.Fields("OldValue").value))
'    .TextMatrix(i, .ColIndex("Remains")) = (IIf(IsNull(rs.Fields("Remains").value), 0, rs.Fields("Remains").value))
    
    
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value))
'   .TextMatrix(i, .ColIndex("Payed")) = (IIf(IsNull(rs.Fields("Payed").value), 0, rs.Fields("Payed").value))
  '.TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
  ''// 19 08 2015
  .TextMatrix(i, .ColIndex("Rent1")) = (IIf(IsNull(rs.Fields("Rent1").value), 0, rs.Fields("Rent1").value))
  .TextMatrix(i, .ColIndex("VATArboon")) = (IIf(IsNull(rs.Fields("VATArboon").value), 0, rs.Fields("VATArboon").value))
  .TextMatrix(i, .ColIndex("RentArbon")) = (IIf(IsNull(rs.Fields("RentArbon").value), 0, rs.Fields("RentArbon").value))
  .TextMatrix(i, .ColIndex("NetRent")) = (IIf(IsNull(rs.Fields("NetRent").value), 0, rs.Fields("NetRent").value))
  .TextMatrix(i, .ColIndex("Commissions1")) = (IIf(IsNull(rs.Fields("Commissions1").value), 0, rs.Fields("Commissions1").value))
  .TextMatrix(i, .ColIndex("CommissionsArbon")) = (IIf(IsNull(rs.Fields("CommissionsArbon").value), 0, rs.Fields("CommissionsArbon").value))
  .TextMatrix(i, .ColIndex("NetCommissions")) = (IIf(IsNull(rs.Fields("NetCommissions").value), 0, rs.Fields("NetCommissions").value))
  .TextMatrix(i, .ColIndex("Insurance1")) = (IIf(IsNull(rs.Fields("Insurance1").value), 0, rs.Fields("Insurance1").value))
  .TextMatrix(i, .ColIndex("InsuranceArbon")) = (IIf(IsNull(rs.Fields("InsuranceArbon").value), 0, rs.Fields("InsuranceArbon").value))
  .TextMatrix(i, .ColIndex("NetInsurance")) = (IIf(IsNull(rs.Fields("NetInsurance").value), 0, rs.Fields("NetInsurance").value))
  .TextMatrix(i, .ColIndex("Water1")) = (IIf(IsNull(rs.Fields("Water1").value), 0, rs.Fields("Water1").value))
  .TextMatrix(i, .ColIndex("WaterArbon")) = (IIf(IsNull(rs.Fields("WaterArbon").value), 0, rs.Fields("WaterArbon").value))
  
  .TextMatrix(i, .ColIndex("Electric1")) = (IIf(IsNull(rs.Fields("Electric1").value), 0, rs.Fields("Electric1").value))
  .TextMatrix(i, .ColIndex("ElectricArbon")) = (IIf(IsNull(rs.Fields("ElectricArbon").value), 0, rs.Fields("ElectricArbon").value))
  .TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
  .TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
  '.TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("NetElectric").value), 0, rs.Fields("NetElectric").value))
  '.TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("NetWater").value), 0, rs.Fields("NetWater").value))
  
  ''//
  Dim X As String
  Dim RentValuePayed   As Double
  Dim CommissionsPayed  As Double
  Dim InsurancePayed    As Double
  Dim WaterPayed   As Double
  Dim ElectricPayed   As Double
  Dim TelandNetPayed  As Double
  Dim payed As Double
  Dim VATPayed As Double
'   getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed)
            payed = getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, , , , VATPayed)

.TextMatrix(i, .ColIndex("RentValuePayed")) = RentValuePayed
.TextMatrix(i, .ColIndex("CommissionsPayed")) = CommissionsPayed
.TextMatrix(i, .ColIndex("InsurancePayed")) = InsurancePayed
.TextMatrix(i, .ColIndex("WaterPayed")) = WaterPayed
.TextMatrix(i, .ColIndex("ElectricPayed")) = ElectricPayed
.TextMatrix(i, .ColIndex("TelandNetPayed")) = TelandNetPayed
.TextMatrix(i, .ColIndex("VATPayed")) = VATPayed
     
    '      payed = payed + (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value)) 'val(rs("RentValuePayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))  ' val(rs("CommissionsPayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))  '   val(rs("InsurancePayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))  ' val(rs("WaterPayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))     'val(rs("ElectricPayed").value)
  '
  '      payed = payed + (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value)) ' val(rs("TelandNetPayed").value)
  '
        .TextMatrix(i, .ColIndex("Payed")) = payed
                    ActulaPyaed = ActulaPyaed + val(.TextMatrix(i, .ColIndex("Payed")))
  .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
   ' .TextMatrix(i, .ColIndex("payedPayed")) = (IIf(IsNull(rs.Fields("payedPayed").value), 0, rs.Fields("payedPayed").value))
   ' .TextMatrix(i, .ColIndex("RemainsPayed")) = (IIf(IsNull(rs.Fields("RemainsPayed").value), 0, rs.Fields("RemainsPayed").value))
    
       .TextMatrix(i, .ColIndex("lastPayedDate")) = Format((IIf(IsNull(rs.Fields("lastPayedDate").value), Format(Date, "yyyy/MM/dd"), rs.Fields("lastPayedDate").value)), "yyyy/MM/dd")
 .TextMatrix(i, .ColIndex("lastPayedDateH")) = Format((IIf(IsNull(rs.Fields("lastPayedDateH").value), Format(ToHijriDate(Date), "yyyy/MM/dd"), rs.Fields("lastPayedDateH").value)), "yyyy/MM/dd")
     .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
      
If rs.RecordCount > 0 Then
  Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
Else
Me.LblTotalQasts.Caption = 0
End If
      LblActulaPyaed.Caption = ActulaPyaed
            lblremain.Caption = val(LblTotalQasts.Caption) - val(LblActulaPyaed)
            
            
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False
        'Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        .RowHeight(-1) = 300
    End With

End Sub

Function checkistallment() As Boolean
Dim installtotals As Double
Dim contracttotals As Double

Dim NpayedValue As Double

With GridInstallments

        If .Rows > 1 Then
            installtotals = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            NpayedValue = .Aggregate(flexSTSum, .FixedRows, .ColIndex("NpayedValue"), .Rows - 1, .ColIndex("NpayedValue"))
            
        Else
           installtotals = 0
           NpayedValue = 0
        End If

          '  contracttotals = val(TxtTotalContract) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtOldRent) + val(TxtOldWater) - NpayedValue
            contracttotals = val(TxtTotalContract) - val(TxtDiscountValue) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) - NpayedValue
            '- val(TxtNotVal.text)
            
            If Round(contracttotals + val(TxtFATValue.Text), 0) <> Round(installtotals, 0) Then
            
              MsgBox " «Ã„«·Ì «·œ›⁄«  ·« Ì ”«ÊÏ  „⁄ «Ã„«·Ì «·⁄ﬁœ ", vbCritical
             checkistallment = False
             Else
             checkistallment = True
             
            
            End If

    End With
End Function
'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If

    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If

    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If

    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If

    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If

    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If

    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If

    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If

    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If

    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If

    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNotID_Change()
If Me.TxtModFlg.Text <> "R" Then

If RtriveInfoOrbon(val(TxtNotID.Text)) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ⁄„· ⁄ﬁœ ·Â–««·⁄—»Ê‰  ·ﬁœ  Ã«Ê“ › —… «·”„«Õ Ê·Ì” ·œÌﬂ ’·«ÕÌ…"
Else
MsgBox "You can not create contract for this earnest money has exceeded the grace period do not have the authority"
End If
Exit Sub
End If
GetSuperVisorOrbion val(TxtNotID.Text)
End If
End Sub

Private Sub TxtNotSreail1_Change()
If TxtNotSreail1.Text <> "" Then
DcbIqara.Enabled = False
TxtSearch.Enabled = False
DcbUnitType.Enabled = False
DcbUnitNo.Enabled = False
Else
DcbIqara.Enabled = True
TxtSearch.Enabled = True
DcbUnitType.Enabled = True
DcbUnitNo.Enabled = True
End If
End Sub

Private Sub TxtPhone_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
Calculte
End If

End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmpCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub

 

Private Sub TxtServce_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtServce.Text)
End Sub

Private Sub TxtTotalContract_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtDiscountValue.Text = val(TxtTotalContract) * val(txtDiscountPercent) * 0.01
    Calculte
End If

 

End Sub

Private Sub txtWater_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
Calculte
End If

End Sub

Private Sub UnitsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid
               
    

        Select Case .ColKey(Col)
 Case "nameunittype"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unittype"), False, True)
                .TextMatrix(Row, .ColIndex("unittype")) = StrAccountCode


 'Case "nameunittype"
 'StrAccountCode = .ComboData
 '               LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unittype"), False, True)
 '               .TextMatrix(Row, .ColIndex("unittype")) = StrAccountCode
                
Case "unitno"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
         Dim X As String
         Dim meterPrice As Double
         Dim lengh As Double
         Dim rentType As Integer
      X = GetIqarUnitData(val(StrAccountCode), , meterPrice, lengh, , rentType)
  .TextMatrix(Row, .ColIndex("meterPrice")) = meterPrice
  .TextMatrix(Row, .ColIndex("length")) = lengh
   ' .TextMatrix(Row, .ColIndex("rentType")) = rentType
   ' If rentType = 0 Then
   '  .TextMatrix(Row, .ColIndex("namerentType")) = "«·ﬁÌ„… «·«ÌÃ«—Ì…"
   ' Else
   ' .TextMatrix(Row, .ColIndex("namerentType")) = "»«·„ — "
   ' End If
    
           If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

  End Select
End With
ReLineGrid
End Sub

Private Sub UnitsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.Text = "R" Then
Cancel = True
Else
With UnitsGrid

      
        Select Case .ColKey(Col)
      
               Case "meterPrice"
    .ComboList = ""
             Case "length"
             .ComboList = ""
             
               Case "RentValue"
             .ComboList = ""
                    Case "Remarks"
             .ComboList = ""
        End Select

    End With
 End If
End Sub

Private Sub UnitsGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid

        Select Case .ColKey(Col)
 
            Case "nameunittype"
             .TextMatrix(Row, .ColIndex("unitno")) = ""
                StrSQL = "select * from TblAkarUnit"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
             Case "unitno"
                StrSQL = "select * from dbo.TblAqarDetai  where id<>" & val(DcbUnitNo.BoundText) & " and (Status IS NULL or Status=0 or Status=2 ) and Aqarid=" & val(DcbIqara.BoundText) & " and unittype=" & val(.TextMatrix(Row, .ColIndex("unittype")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "unitno", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "unitno", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 


 Case "namerentType"
                StrSQL = "select * from TblRentType"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
    


          
   


        End Select

    End With
    ReLineGrid

End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Row = VSFlexGrid1.Rows - 1 Then
    
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        End If
ReLineGrid
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
  
    Dim StrAccountType As String
    Dim StrComboList As String
 

Dim StrAccountCode1 As String
Dim i As Integer

    With VSFlexGrid2
               
    

        Select Case .ColKey(Col)
         Case "group"
        StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("groupid"), False, True)
                .TextMatrix(Row, .ColIndex("groupid")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("empname")) = ""
                .TextMatrix(Row, .ColIndex("id")) = ""
                .TextMatrix(Row, .ColIndex("code")) = ""
                
 Case "empname"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
                '''//
                         
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID, dbo.TBLSalesRepData.id, dbo.TblEmployee.Fullcode, dbo.TBLSalesRepData.GroupID, "
    StrSQL = StrSQL & "                 dbo.TBLSalesRepGroups.name ,dbo.TBLSalesRepGroups.NameE "
   
    StrSQL = StrSQL & " FROM         dbo.TBLSalesRepGroups RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TBLSalesRepData ON dbo.TBLSalesRepGroups.id = dbo.TBLSalesRepData.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"

    StrSQL = StrSQL & " where dbo.TBLSalesRepData.EmpID  = " & val(StrAccountCode) & ""
                ''//
                
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 If rs.RecordCount > 0 Then
                  .TextMatrix(Row, .ColIndex("groupid")) = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                  Else
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                  End If
                  
                  .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                  .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                  Else
                   .TextMatrix(Row, .ColIndex("code")) = ""
                   End If
                               
For i = 1 To .Rows - 1
If Row <> i Then
If (.TextMatrix(i, .ColIndex("id")) = .TextMatrix(Row, .ColIndex("id"))) And (.TextMatrix(i, .ColIndex("groupid")) = .TextMatrix(Row, .ColIndex("groupid"))) Then
MsgBox "·«Ì„ﬂ‰  ﬂ—«— «·„‰œÊ» "
.TextMatrix(Row, .ColIndex("id")) = 0
.TextMatrix(Row, .ColIndex("empname")) = ""
Exit Sub
End If
End If
Next i
                '''//
                      If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.id , dbo.TblEmployee.Fullcode "
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.id , dbo.TblEmployee.Fullcode "
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    StrSQL = StrSQL & " where dbo.TBLSalesRepData.EmpID = " & val(StrAccountCode) & ""
                ''//
                Set rs = New ADODB.Recordset
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 If rs.RecordCount > 0 Then

                  .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                  .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                  Else
                   .TextMatrix(Row, .ColIndex("code")) = ""
                   End If
               Case "code"
     
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID, dbo.TblEmployee.Emp_Name, dbo.TBLSalesRepData.id, dbo.TBLSalesRepData.GroupID, dbo.TBLSalesRepGroups.name, "
    StrSQL = StrSQL & "                   dbo.TBLSalesRepGroups.NameE"
   StrSQL = StrSQL & "  FROM         dbo.TBLSalesRepGroups RIGHT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TBLSalesRepData ON dbo.TBLSalesRepGroups.id = dbo.TBLSalesRepData.GroupID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"
    StrSQL = StrSQL & " where dbo.TblEmployee.Fullcode ='" & .TextMatrix(Row, .ColIndex("code")) & "'"
    
                   rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
                     .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
                      .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                Else
                .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs("emp_nameE").value), "", rs("emp_nameE").value)
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
                     .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                End If
                End If
          
                For i = 1 To .Rows - 1
If Row <> i Then
If .TextMatrix(i, .ColIndex("id")) = .TextMatrix(Row, .ColIndex("id")) And (.TextMatrix(i, .ColIndex("groupid")) = .TextMatrix(Row, .ColIndex("groupid"))) Then
MsgBox "·«Ì„ﬂ‰  ﬂ—«— «·„‰œÊ» "
.TextMatrix(Row, .ColIndex("id")) = 0
.TextMatrix(Row, .ColIndex("empname")) = ""
.TextMatrix(Row, .ColIndex("idd")) = 0
Exit Sub
End If
End If
Next i

               ' StrSQL = " select Fullcode from TblEmployee where Emp_ID= " & val(StrAccountCode) & ""
               '  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               '  If rs.RecordCount > 0 Then
               '   .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
               '   Else
               '    .TextMatrix(Row, .ColIndex("code")) = ""
               '    End If
                
             
       
    
  End Select
      If Row = .Rows - 1 Then
            .Rows = .Rows + 1
             End If
End With
ReLineGrid
                
                
                
                
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.Text = "R" Then
Cancel = True
Else
With VSFlexGrid2

      
        Select Case .ColKey(Col)
      
 
           
               Case "code"
             .ComboList = ""
                    Case "rate"
             .ComboList = ""
        End Select

    End With
 End If
End Sub

'Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
'End Function


Private Sub VSFlexGrid2_Click()

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid2

        Select Case .ColKey(Col)
                   Case "group"
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     id ,  name "
    Else
    StrSQL = "SELECT     id , namee"
    End If
    StrSQL = StrSQL & " FROM  TBLSalesRepGroups "
    
    
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
 
            Case "empname"
              If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.GroupID"
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.GroupID"
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    If val(.TextMatrix(Row, .ColIndex("groupid"))) <> 0 Then
    StrSQL = StrSQL & " where dbo.TBLSalesRepData.GroupID=" & val(.TextMatrix(Row, .ColIndex("groupid"))) & ""
    End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "emp_name", "EmpID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "emp_nameE", "EmpID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 End Select
                 End With
End Sub



Private Sub RetriveOldPayment()
    Dim Num As Integer
    On Error GoTo ErrTrap
    GridInstallments2.Clear flexClearScrollable, flexClearEverything
        Dim rs As New ADODB.Recordset
    Dim My_SQL  As String
    Dim i As Long

Dim notpayed As Double
notpayed = 0
 
My_SQL = " SELECT  * from TblContractInstallmentsOld"
My_SQL = My_SQL & " WHERE     (ContNo =" & val(TxtContNo) & ")  order by InstallNo"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments2
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
         '  .TextMatrix(i, .ColIndex("DevID")) = (IIf(IsNull(rs.Fields("DevID").value), 0, rs.Fields("DevID").value))
          .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
      '    .TextMatrix(i, .ColIndex("TempInstal")) = (IIf(IsNull(rs.Fields("TempInstal").value), 0, rs.Fields("TempInstal").value))
          .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
     '     .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 1, rs.Fields("hijri").value))
   '       .TextMatrix(i, .ColIndex("DES")) = (IIf(IsNull(rs.Fields("DES").value), "", rs.Fields("DES").value))
          .TextMatrix(i, .ColIndex("Due_DateH")) = Format((IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value)), "yyyy/MM/dd")
           .TextMatrix(i, .ColIndex("Due_Date")) = Format(IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value), "yyyy/MM/dd")
        'yyyy/MM/dd
       .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
   '    .TextMatrix(i, .ColIndex("ServiceArbon")) = (IIf(IsNull(rs.Fields("ServiceArbon").value), 0, rs.Fields("ServiceArbon").value))
       .TextMatrix(i, .ColIndex("NoteSerial")) = (IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value))
       .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
       .TextMatrix(i, .ColIndex("NoteId")) = (IIf(IsNull(rs.Fields("NoteId").value), "", rs.Fields("NoteId").value))
If Not IsNull(rs.Fields("Status").value) Then
             If rs.Fields("Status").value = 0 Then
                    .Cell(flexcpChecked, i, .ColIndex("Status")) = flexUnchecked
            Else
                     .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked
                       notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
            End If

End If

    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("VATPayed")) = (IIf(IsNull(rs.Fields("VATPayed").value), 0, rs.Fields("VATPayed").value))
    .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs.Fields("VATValue").value), 0, rs.Fields("VATValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
 '.TextMatrix(i, .ColIndex("NpayedValue")) = (IIf(IsNull(rs.Fields("NpayedValue").value), 0, rs.Fields("NpayedValue").value))
        
  '  .TextMatrix(i, .ColIndex("OldValue")) = (IIf(IsNull(rs.Fields("OldValue").value), 0, rs.Fields("OldValue").value))
'    .TextMatrix(i, .ColIndex("Remains")) = (IIf(IsNull(rs.Fields("Remains").value), 0, rs.Fields("Remains").value))
    
    
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value))
   .TextMatrix(i, .ColIndex("Payed")) = (IIf(IsNull(rs.Fields("Payed").value), 0, rs.Fields("Payed").value))
  '.TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
  ''// 19 08 2015
'  .TextMatrix(i, .ColIndex("Rent1")) = (IIf(IsNull(rs.Fields("Rent1").value), 0, rs.Fields("Rent1").value))
'  .TextMatrix(i, .ColIndex("RentArbon")) = (IIf(IsNull(rs.Fields("RentArbon").value), 0, rs.Fields("RentArbon").value))
'  .TextMatrix(i, .ColIndex("NetRent")) = (IIf(IsNull(rs.Fields("NetRent").value), 0, rs.Fields("NetRent").value))
'  .TextMatrix(i, .ColIndex("Commissions1")) = (IIf(IsNull(rs.Fields("Commissions1").value), 0, rs.Fields("Commissions1").value))
  .TextMatrix(i, .ColIndex("CommissionsArbon")) = (IIf(IsNull(rs.Fields("CommissionsArbon").value), 0, rs.Fields("CommissionsArbon").value))
  .TextMatrix(i, .ColIndex("NetCommissions")) = (IIf(IsNull(rs.Fields("NetCommissions").value), 0, rs.Fields("NetCommissions").value))
  .TextMatrix(i, .ColIndex("Insurance1")) = (IIf(IsNull(rs.Fields("Insurance1").value), 0, rs.Fields("Insurance1").value))
  .TextMatrix(i, .ColIndex("InsuranceArbon")) = (IIf(IsNull(rs.Fields("InsuranceArbon").value), 0, rs.Fields("InsuranceArbon").value))
  .TextMatrix(i, .ColIndex("NetInsurance")) = (IIf(IsNull(rs.Fields("NetInsurance").value), 0, rs.Fields("NetInsurance").value))
  .TextMatrix(i, .ColIndex("Water1")) = (IIf(IsNull(rs.Fields("Water1").value), 0, rs.Fields("Water1").value))
  .TextMatrix(i, .ColIndex("WaterArbon")) = (IIf(IsNull(rs.Fields("WaterArbon").value), 0, rs.Fields("WaterArbon").value))
  
  .TextMatrix(i, .ColIndex("Electric1")) = (IIf(IsNull(rs.Fields("Electric1").value), 0, rs.Fields("Electric1").value))
  .TextMatrix(i, .ColIndex("ElectricArbon")) = (IIf(IsNull(rs.Fields("ElectricArbon").value), 0, rs.Fields("ElectricArbon").value))
  .TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
  .TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
  .TextMatrix(i, .ColIndex("NetElectric")) = (IIf(IsNull(rs.Fields("NetElectric").value), 0, rs.Fields("NetElectric").value))
  .TextMatrix(i, .ColIndex("NetWater")) = (IIf(IsNull(rs.Fields("NetWater").value), 0, rs.Fields("NetWater").value))
  
  ''//
  Dim X As String
  Dim RentValuePayed   As Double
  Dim CommissionsPayed  As Double
  Dim InsurancePayed    As Double
  Dim WaterPayed   As Double
  Dim ElectricPayed   As Double
  Dim TelandNetPayed  As Double
  Dim payed As Double
  Dim VATPayed As Double
'   getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed)
            payed = getinsttPayedTocontract(val(rs.Fields("id").value), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, , , , VATPayed)

.TextMatrix(i, .ColIndex("RentValuePayed")) = RentValuePayed
.TextMatrix(i, .ColIndex("CommissionsPayed")) = CommissionsPayed
.TextMatrix(i, .ColIndex("InsurancePayed")) = InsurancePayed
.TextMatrix(i, .ColIndex("WaterPayed")) = WaterPayed
.TextMatrix(i, .ColIndex("ElectricPayed")) = ElectricPayed
.TextMatrix(i, .ColIndex("TelandNetPayed")) = TelandNetPayed
.TextMatrix(i, .ColIndex("VATPayed")) = VATPayed
     
    '      payed = payed + (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value)) 'val(rs("RentValuePayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))  ' val(rs("CommissionsPayed").value)
     
  '        payed = payed + (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))  '   val(rs("InsurancePayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))  ' val(rs("WaterPayed").value)
  '
  '        payed = payed + (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))     'val(rs("ElectricPayed").value)
  '
  '      payed = payed + (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value)) ' val(rs("TelandNetPayed").value)
  '
        .TextMatrix(i, .ColIndex("Payed")) = payed
              
  .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Payed")))
   ' .TextMatrix(i, .ColIndex("payedPayed")) = (IIf(IsNull(rs.Fields("payedPayed").value), 0, rs.Fields("payedPayed").value))
   ' .TextMatrix(i, .ColIndex("RemainsPayed")) = (IIf(IsNull(rs.Fields("RemainsPayed").value), 0, rs.Fields("RemainsPayed").value))
    
       .TextMatrix(i, .ColIndex("lastPayedDate")) = Format((IIf(IsNull(rs.Fields("lastPayedDate").value), Format(Date, "yyyy/MM/dd"), rs.Fields("lastPayedDate").value)), "yyyy/MM/dd")
 .TextMatrix(i, .ColIndex("lastPayedDateH")) = Format((IIf(IsNull(rs.Fields("lastPayedDateH").value), Format(ToHijriDate(Date), "yyyy/MM/dd"), rs.Fields("lastPayedDateH").value)), "yyyy/MM/dd")
     .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
If rs.RecordCount > 0 Then
  Me.LblTotalQasts2.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
Else
Me.LblTotalQasts2.Caption = 0
End If
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False
        'Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        .RowHeight(-1) = 300
    End With


My_SQL = "Select ContNo,UserId,id,EditDate,UserName = (Select UserName From TblUsers Where UserId =TblContractInstallmentsHist.UserID ) from TblContractInstallmentsHist Where ContNo = " & Trim(TxtContNo.Text)
Set rs = New ADODB.Recordset
rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
  With Me.grdHistory
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst
      
            For i = 1 To .Rows - 1
        
            .TextMatrix(i, .ColIndex("UserID")) = rs!UserID & ""
            .TextMatrix(i, .ColIndex("EditDate")) = rs!EditDate & ""
            .TextMatrix(i, .ColIndex("UserName")) = rs!UserName & ""
            '.TextMatrix(i, .ColIndex("ContNo")) = rs!ContNo & ""
rs.MoveNext
        Next
        rs.Close
        
        .AutoSize 1, .Cols - 1, False
        'Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        .RowHeight(-1) = 300
        End If
        End With
        
    Exit Sub
ErrTrap:
End Sub



Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "«”„ «·⁄ﬁ«— " & CHR(13) & DcbIqara.Text & CHR(13) & " —ﬁ„ «·⁄ﬁœ   " & TxtNoteSerial1.Text & CHR(13) & "«·„” √Ã—" & dcCustomer.Text & CHR(13) & " «· «—ÌŒ " & RecorddateH.value & CHR(13) & ContDate.value & CHR(13) & "  «·„«·ﬂ   " & dcsupplier.Text & CHR(13)
    LogTextA = LogTextA & "⁄œœ «·œ›⁄« " & TxtPaymentCount & CHR(13)
     LogTextA = LogTextA & "› —… «·œ›⁄« " & TxtPeriods & "  " & DcbPeriodsID.Text & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·«ÌÃ«—" & TxtTotalContract & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·”⁄Ì" & TxtCommiValue & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·‹ √„Ì‰" & TxtInsuranceValue & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·ﬂÂ—»«¡" & TxtElectricity & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·Œœ„« /«·”⁄Ì" & TxtPhone & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·„Ì«Â " & TxtWater & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·”⁄Ì «·Œ«—ÃÌ " & TxtOutOffice & CHR(13)
     LogTextA = LogTextA & "  «·’«›Ì     " & TxtNetValue & CHR(13)
     LogTextA = LogTextA & "  ‰”»Â «·›«      " & TxtFATYou & CHR(13)
     LogTextA = LogTextA & "  ﬁÌ„Â «·›«      " & TxtFATValue & CHR(13)
     LogTextA = LogTextA & "    «·«Ã„«·Ì     " & TxtTotalValue & CHR(13)
     LogTextA = LogTextA & "    —ﬁ„ ”‰œ «·⁄—»Ê‰     " & TxtNotID & CHR(13)
     LogTextA = LogTextA & "    ﬁÌ„Â ”‰œ «·⁄—»Ê‰     " & TxtNotVal & CHR(13)
     LogTextA = LogTextA & "   »œ«ÌÂ «·⁄›œ ÂÃ—Ì     " & FromDateH.value & CHR(13)
     LogTextA = LogTextA & "   »œ«ÌÂ «·⁄›œ „Ì·«œÌ     " & StrDate.value & CHR(13)
     LogTextA = LogTextA & "   ‰Â«Ì… «·⁄›œ ÂÃ—Ì     " & todateH.value & CHR(13)
     LogTextA = LogTextA & "   ‰Â«Ì… «·⁄›œ „Ì·«œÌ     " & ToDate.value & CHR(13)
     LogTextA = LogTextA & "   „œÂ «·⁄›œ       " & Contract_period_no & "  " & Contract_period.Text & CHR(13)
     LogTextA = LogTextA & " —’Ìœ „ »ﬁÌ ⁄ﬁœ " & TxtOldRent.Text & CHR(13)
     LogTextA = LogTextA & " —’Ìœ „ »ﬁÌ  „Ì«Â" & TxtOldWater.Text & CHR(13)
     LogTextA = LogTextA & " —’Ìœ „ »ﬁÌ  ﬂÂ—»«¡" & TxtOldElectric.Text & CHR(13)
     LogTextA = LogTextA & " —’Ìœ „ »ﬁÌ  Œœ„« " & TxtoldCommi.Text & CHR(13)
     
     LogTextA = LogTextA & "    „‰  «—ÌŒ " & balanceDateH.value & CHR(13) & balanceDate.value & CHR(13)
     LogTextA = LogTextA & "  „·«ÕŸ« " & balanceDes.Text & CHR(13)
     LogTextA = LogTextA & "  «ﬁ· ﬁÌ„Â «ÌÃ«—Ì…" & TxtMiniRentValue.Text & CHR(13)
     LogTextA = LogTextA & "   √„Ì‰ „”œœ ”«»ﬁ" & txtOldInsurance.Text & CHR(13)
     If Opt(0).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ ÃœÌœ" & CHR(13)
     End If
     
     If Opt(1).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ «›  «ÕÌ" & CHR(13)
     End If
     
   If ComResid(0).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ€Ì— Œ«÷⁄" & CHR(13)
     End If
     
     If ComResid(1).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ  Œ«÷⁄" & CHR(13)
     End If
       
       
       
   If RdRTypeDate(0).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ ÂÃ—Ì    " & CHR(13)
     End If
     
     If RdRTypeDate(1).value = True Then
     LogTextA = LogTextA & " ‰Ê⁄ «·⁄ﬁœ „Ì·«œÌ " & CHR(13)
     End If
       
       
     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Real Estate Name " & CHR(13) & DcbIqara.Text & " Contract No. " & TxtNoteSerial1.Text & CHR(13) & " Date " & Date & CHR(13) & " Owner" & dcsupplier.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , val(TxtNoteSerial.Text), TxtNoteSerial1.Text
    Else
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtNoteSerial.Text), TxtNoteSerial1.Text
    End If
End Function

