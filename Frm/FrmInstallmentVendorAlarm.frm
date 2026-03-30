VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInstallmentVendorAlarm 
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17880
   Icon            =   "FrmInstallmentVendorAlarm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkPrintDirect 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÿ»«⁄… „»«‘—…"
      Height          =   525
      Left            =   11730
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   60
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   6960
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic12 
      Height          =   10950
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17880
      _cx             =   31538
      _cy             =   19315
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
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
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
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   10845
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   17835
         _cx             =   31459
         _cy             =   19129
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
         Caption         =   " ‰»ÌÂ«  «·« ›«ﬁÌ« | ‰»ÌÂ«  «ﬁ”«ÿ «·«’Ê·  |«Ê«„— «·«‰ «Ã"
         Align           =   0
         CurrTab         =   2
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
            Height          =   10470
            Index           =   1
            Left            =   -18690
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   17745
            _cx             =   31300
            _cy             =   18468
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
            Begin VB.CommandButton cmdPrint 
               Caption         =   "ÿ»«⁄…"
               Height          =   375
               Left            =   1680
               TabIndex        =   56
               Top             =   9270
               Width           =   2355
            End
            Begin VB.Frame Frame6 
               Caption         =   " "
               Height          =   1125
               Left            =   1650
               TabIndex        =   49
               Top             =   8010
               Width           =   15915
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»ÌÂ »‰”» «· ‰›Ì– „⁄ ‰”» «·„ﬁ»Ê÷« "
                  Height          =   345
                  Index           =   5
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   690
                  Width           =   6195
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·« ›«ﬁÌ«  «· Ï ﬁ«—»  ⁄·Ï «·«‰ Â«¡ „‰ «⁄„«· «· —ﬂÌ» »«·‰”»… «·„⁄—›… ›Ï «·« ›«ﬁÌ…"
                  Height          =   345
                  Index           =   4
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   270
                  Width           =   6195
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·« ›«ﬁÌ«  «· Ï ·„ Ì „ ⁄·ÌÂ« Õ—ﬂ… —›⁄ „ﬁÌ«”"
                  Height          =   345
                  Index           =   3
                  Left            =   7740
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   720
                  Width           =   4005
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·« ›«ﬁÌ«  «· Ï ·„ Ì „ ⁄·ÌÂ« Õ—ﬂ… ÿ·»« "
                  Height          =   345
                  Index           =   2
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   270
                  Width           =   3225
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·« ›«ﬁÌ«  «· Ï ·„ Ì „  ”ÃÌ· «·«⁄„«· «·ÌÊ„Ì… ·Â«"
                  Height          =   345
                  Index           =   1
                  Left            =   12030
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   720
                  Width           =   3795
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·« ›«ﬁÌ«  «· Ï ·„ Ì „  ÕœÌœ „œ… “„‰Ì… ·Â«"
                  Height          =   345
                  Index           =   0
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   270
                  Width           =   3225
               End
            End
            Begin VB.Timer Timer2 
               Left            =   0
               Top             =   0
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   1
               Left            =   8445
               TabIndex        =   5
               Text            =   "5"
               Top             =   9930
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.TextBox txtPercent 
               Alignment       =   1  'Right Justify
               Height          =   435
               Index           =   0
               Left            =   11910
               TabIndex        =   4
               Text            =   "75"
               Top             =   9330
               Width           =   1950
            End
            Begin VB.TextBox txtPercent 
               Alignment       =   1  'Right Justify
               Height          =   435
               Index           =   1
               Left            =   11910
               TabIndex        =   3
               Text            =   "100"
               Top             =   9855
               Width           =   1950
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   990
               Index           =   0
               Left            =   9780
               TabIndex        =   6
               Top             =   9330
               Visible         =   0   'False
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   1746
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «· ‰»ÌÂ"
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Times New Roman"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmInstallmentVendorAlarm.frx":6852
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   12632064
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   12632064
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   7680
               Left            =   0
               TabIndex        =   7
               Top             =   165
               Width           =   17520
               _cx             =   30903
               _cy             =   13547
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
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInstallmentVendorAlarm.frx":D0B4
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
               Begin MSComctlLib.ProgressBar ProgressBar1 
                  Height          =   615
                  Left            =   5040
                  TabIndex        =   55
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   8415
                  _ExtentX        =   14843
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ÕœÌÀ ﬂ·"
               Height          =   570
               Index           =   8
               Left            =   8445
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   9420
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "‰”»… «· ‰›Ì– «ﬂ»— „‰ «Ê Ì”«ÊÏ"
               Height          =   345
               Index           =   0
               Left            =   14160
               TabIndex        =   9
               Top             =   9420
               Width           =   3375
            End
            Begin VB.Label Label1 
               Caption         =   "‰”»… «·„»·€ «·„ﬁ»Ê÷ «ﬁ· „‰ «Ê Ì”«ÊÌ"
               Height          =   330
               Index           =   1
               Left            =   14160
               TabIndex        =   8
               Top             =   9960
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   10470
            Index           =   0
            Left            =   -18390
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   45
            Width           =   17745
            _cx             =   31300
            _cy             =   18468
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   1170
               Left            =   0
               TabIndex        =   37
               Top             =   9090
               Width           =   17925
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   495
                  Index           =   6
                  Left            =   480
                  TabIndex        =   38
                  Top             =   240
                  Width           =   3045
                  _ExtentX        =   5371
                  _ExtentY        =   873
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
                  ButtonImage     =   "FrmInstallmentVendorAlarm.frx":D3AC
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
                  Height          =   495
                  Left            =   5040
                  TabIndex        =   39
                  Top             =   240
                  Width           =   2835
                  _ExtentX        =   5001
                  _ExtentY        =   873
                  ButtonPositionImage=   1
                  Caption         =   "„”Õ"
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
                  ButtonImage     =   "FrmInstallmentVendorAlarm.frx":36FCE
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
                  Height          =   495
                  Index           =   9
                  Left            =   9240
                  TabIndex        =   40
                  Top             =   240
                  Width           =   3045
                  _ExtentX        =   5371
                  _ExtentY        =   873
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
                  ButtonImage     =   "FrmInstallmentVendorAlarm.frx":3D830
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   6285
               Left            =   0
               TabIndex        =   34
               Top             =   2805
               Width           =   17925
               Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                  Height          =   5055
                  Left            =   120
                  TabIndex        =   35
                  Top             =   120
                  Width           =   18705
                  _cx             =   32994
                  _cy             =   8916
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
                  BackColorAlternate=   16777152
                  GridColor       =   0
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmInstallmentVendorAlarm.frx":44092
                  ScrollTrack     =   -1  'True
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
                  Begin VB.Label Label2 
                     Caption         =   "%"
                     Height          =   375
                     Index           =   0
                     Left            =   10440
                     TabIndex        =   36
                     Top             =   -600
                     Width           =   375
                  End
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Height          =   1455
               Left            =   0
               TabIndex        =   21
               Top             =   -60
               Width           =   17925
               Begin VB.TextBox Text5 
                  Height          =   288
                  Index           =   0
                  Left            =   8040
                  TabIndex        =   25
                  Top             =   600
                  Width           =   1692
               End
               Begin VB.TextBox Text4 
                  Height          =   288
                  Left            =   8040
                  TabIndex        =   24
                  Top             =   240
                  Width           =   1692
               End
               Begin VB.TextBox Text3 
                  Height          =   288
                  Left            =   11160
                  TabIndex        =   23
                  Top             =   600
                  Width           =   1692
               End
               Begin VB.TextBox Text2 
                  Height          =   288
                  Left            =   11160
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1692
               End
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   288
                  Left            =   4680
                  TabIndex        =   26
                  Top             =   600
                  Width           =   1692
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   336
                  Left            =   4680
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1692
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   146997249
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·›« Ê—…"
                  Height          =   288
                  Index           =   2
                  Left            =   6360
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1248
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
                  Height          =   288
                  Index           =   7
                  Left            =   9480
                  TabIndex        =   32
                  Top             =   600
                  Width           =   1488
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁ”ÿ"
                  Height          =   288
                  Index           =   6
                  Left            =   9600
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1368
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ ›« Ê—… «·„Ê—œ"
                  Height          =   288
                  Index           =   5
                  Left            =   12960
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1368
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·›« Ê—…"
                  Height          =   288
                  Index           =   1
                  Left            =   13320
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1008
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ê—œ"
                  Height          =   288
                  Index           =   3
                  Left            =   6600
                  TabIndex        =   28
                  Top             =   600
                  Width           =   888
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   1440
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   1380
               Width           =   17925
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   6960
                  TabIndex        =   18
                  Text            =   "5"
                  Top             =   600
                  Width           =   810
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "› —… «·«” Õﬁ«ﬁ"
                  Height          =   735
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   240
                  Width           =   5895
                  Begin MSComCtl2.DTPicker todate 
                     Height          =   330
                     Left            =   360
                     TabIndex        =   14
                     Top             =   240
                     Width           =   1695
                     _ExtentX        =   2990
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   146997249
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker Fromdate 
                     Height          =   336
                     Left            =   3120
                     TabIndex        =   15
                     Top             =   240
                     Width           =   1692
                     _ExtentX        =   2990
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   146997249
                     CurrentDate     =   41640
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰"
                     Height          =   315
                     Index           =   0
                     Left            =   4680
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   240
                     Width           =   585
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "≈«·Ï"
                     Height          =   435
                     Index           =   14
                     Left            =   2100
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   240
                     Width           =   540
                  End
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   735
                  Index           =   5
                  Left            =   240
                  TabIndex        =   19
                  Top             =   240
                  Width           =   6285
                  _ExtentX        =   11086
                  _ExtentY        =   1296
                  ButtonPositionImage=   1
                  Caption         =   " ÕœÌÀ"
                  BackColor       =   14871017
                  FontSize        =   14.25
                  FontName        =   "Times New Roman"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   14.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmInstallmentVendorAlarm.frx":441F7
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   12632064
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   12632064
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " ÕœÌÀ ﬂ·"
                  Height          =   435
                  Index           =   4
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   240
                  Width           =   780
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   10320
               Index           =   1
               Left            =   24705
               TabIndex        =   41
               Top             =   900
               Width           =   17520
               _cx             =   30903
               _cy             =   18203
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
               FormatString    =   $"FrmInstallmentVendorAlarm.frx":4AA59
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   10470
            Index           =   2
            Left            =   45
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   45
            Width           =   17745
            _cx             =   31300
            _cy             =   18468
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
            Begin VB.CommandButton Command2 
               Caption         =   "⁄—÷ «·Œ«„« "
               Height          =   405
               Left            =   11280
               TabIndex        =   65
               Top             =   9420
               Width           =   2055
            End
            Begin VB.CommandButton cmdAdminPer 
               Caption         =   "⁄—÷ »’·«ÕÌ«  «·«œ„‰"
               Height          =   405
               Left            =   11280
               TabIndex        =   63
               Top             =   9900
               Width           =   2055
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   2
               Left            =   8565
               TabIndex        =   43
               Text            =   "2"
               Top             =   9930
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Timer Timer3 
               Interval        =   20000
               Left            =   780
               Top             =   0
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   540
               Index           =   1
               Left            =   525
               TabIndex        =   44
               Top             =   9840
               Width           =   7680
               _ExtentX        =   13547
               _ExtentY        =   953
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Times New Roman"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmInstallmentVendorAlarm.frx":4AB19
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   12632064
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   12632064
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin VSFlex8UCtl.VSFlexGrid grd 
               Height          =   8550
               Left            =   60
               TabIndex        =   45
               Top             =   630
               Width           =   17475
               _cx             =   30824
               _cy             =   15081
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
               GridLines       =   13
               GridLinesFixed  =   2
               GridLineWidth   =   40
               Rows            =   50
               Cols            =   51
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   800
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInstallmentVendorAlarm.frx":5137B
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
               Ellipsis        =   1
               ExplorerBar     =   7
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   -1  'True
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   -1  'True
               TextStyle       =   3
               TextStyleFixed  =   4
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
               Begin VB.Frame FramRows 
                  Caption         =   "«·„Ê«œ «·Œ«„"
                  Height          =   6615
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   1650
                  Visible         =   0   'False
                  Width           =   14055
                  Begin VB.TextBox txtMySQL 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   930
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   61
                     Top             =   90
                     Visible         =   0   'False
                     Width           =   10035
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FGRows 
                     Height          =   5685
                     Left            =   120
                     TabIndex        =   60
                     Top             =   690
                     Width           =   13830
                     _cx             =   24395
                     _cy             =   10028
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
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   54
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmInstallmentVendorAlarm.frx":51B77
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
                     WallPaperAlignment=   0
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
                  Begin VB.Label Label20 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "X"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   495
                     Left            =   12810
                     TabIndex        =   59
                     Top             =   240
                     Width           =   1335
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic EleHeader 
               Height          =   585
               Left            =   0
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   0
               Width           =   17715
               _cx             =   31247
               _cy             =   1032
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   21.75
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
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "    ‰»ÌÂ«  «ﬁ”«ÿ «·«’Ê·   "
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
               Begin VB.CommandButton Command1 
                  Caption         =   "Command1"
                  Height          =   405
                  Left            =   5460
                  TabIndex        =   62
                  Top             =   1530
                  Width           =   1785
               End
               Begin VB.Image Image1 
                  Height          =   555
                  Index           =   0
                  Left            =   9840
                  Picture         =   "FrmInstallmentVendorAlarm.frx":5238B
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  ForeColor       =   &H000000FF&
                  Height          =   555
                  Index           =   27
                  Left            =   2520
                  TabIndex        =   48
                  Top             =   0
                  Width           =   2205
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ÕœÌÀ ﬂ·"
               Height          =   585
               Index           =   9
               Left            =   9330
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   9990
               Visible         =   0   'False
               Width           =   1005
            End
         End
      End
   End
End
Attribute VB_Name = "FrmInstallmentVendorAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LngRow As Long
Dim mQtyTotal As Double
Dim mOld As Boolean
Dim mTotalSecond As Double
Dim CostTOTAL As Double
Dim mProdId  As Long
Dim mIsByAdmin As Boolean
Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
ProgressBar1.Visible = True
: ProgressBar1.value = 10
FillGrid
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
print_report
Case 0
    mOld = True
    FillGrid2 , , True
Case 1
    FillGrid3
    
End Select

End Sub
Private Sub FillGridRows(ByVal mRow As Long)
  FGRows.Clear flexClearScrollable, flexClearEverything
    FGRows.Rows = 2
    FGRows.Clear flexClearScrollable, flexClearEverything
    FGRows.Refresh
    Dim mItemId As Long
    Dim mLineNo As Long
    Dim mIID As Long
    mItemId = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
    mLineNo = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
    mIID = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
    Dim RsDetails As New ADODB.Recordset

    StrSQL = "SELECT    DISTINCT T2.ItemName ItemName2,T2.ItemNamee ItemNamee2, ItemCode2,ItemID2, TblDefComItemDet.OldPrice,   TblDefComItemDet.lowering,TblDefComItemDet.increase,dbo.TblDefComItemDet.ID,TblDefComItemDet.IsDeleted, dbo.TblDefComItemDet.IDDefCIT,dbo.TblDefComItemDet.IsAdd,dbo.TblDefComItemDet.Price,dbo.TblDefComItemDet.Total, dbo.TblDefComItemDet.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
    StrSQL = StrSQL & "                   dbo.TblItems.Fullcode ,dbo.TblItems.ItemNamee, dbo.TblDefComItemDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblDefComItemDet.SpecID1,"
    StrSQL = StrSQL & "                  dbo.TblSpecification.Name AS Name1, dbo.TblSpecification.Namee AS Namee1, dbo.TblDefComItemDet.SpecID2, TblSpecification_1.Name AS Name2,"
    StrSQL = StrSQL & "                  TblSpecification_1.Namee AS Namee2, dbo.TblDefComItemDet.SpecID3, TblSpecification_2.Name AS Name3, TblSpecification_2.Namee AS Namee3,"
    StrSQL = StrSQL & "                  dbo.TblDefComItemDet.SpecID4, TblSpecification_3.Name AS Name4, TblSpecification_3.Namee AS Namee4, dbo.TblDefComItemDet.Amout1,"
    StrSQL = StrSQL & "                 dbo.TblDefComItemDet.Amout2 ,TblDefComItemDet.LineID ,dbo.TblDefComItemDet.Amout3, dbo.TblDefComItemDet.Amout4, dbo.TblDefComItemDet.Qty, dbo.TblDefComItemDet.cost ,dbo.TblDefComItemDet.FlgX ,dbo.TblDefComItemDet.TepQty"
    StrSQL = StrSQL & " FROM         dbo.TblDefComItemDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblSpecification TblSpecification_3 ON dbo.TblDefComItemDet.SpecID4 = TblSpecification_3.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                 dbo.TblSpecification TblSpecification_2 ON dbo.TblDefComItemDet.SpecID3 = TblSpecification_2.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblSpecification TblSpecification_1 ON dbo.TblDefComItemDet.SpecID2 = TblSpecification_1.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                 dbo.TblSpecification ON dbo.TblDefComItemDet.SpecID1 = dbo.TblSpecification.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.TblDefComItemDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.TblDefComItemDet.ItemID = dbo.TblItems.ItemID"
    
    StrSQL = StrSQL & "                  Left OUTER JOIN "
  

    


    StrSQL = StrSQL & "      dbo.TblItems T2 ON dbo.TblDefComItemDet.ItemID2 = T2.ItemID"
    StrSQL = StrSQL & " Where dbo.TblDefComItemDet.IDDefCIT =" & mIID
    StrSQL = StrSQL & " And TblDefComItemDet.ItemID2 = " & val(mItemId)
    StrSQL = StrSQL & " And TblDefComItemDet.LineID = " & val(mLineNo)
    StrSQL = StrSQL & " Order By TblDefComItemDet.ID"

 

 
    

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim Num As Long
Dim mm  As Long
Dim mTableID As String
mTableID = "(0,0"
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FGRows.Rows = RsDetails.RecordCount + 1
        
                  Dim rsDummy2 As New ADODB.Recordset
          Dim PartItemQty As Double, ForUnit As Double, lowering As Double, increase As Double, MethodCalc As Double
        For Num = 1 To RsDetails.RecordCount
        
            
            StrSQL = "select    dbo.TblItemsParts.PartItemQty ,TableID, ForUnit,    TblItemsParts.lowering,TblItemsParts.increase,MethodCalc from TblItemsParts"
            StrSQL = StrSQL & " Where dbo.TblItemsParts.ItemID = " & val(RsDetails!ItemID2 & "")
            StrSQL = StrSQL & " and PartItemID = " & val(RsDetails!ItemID & "")
            StrSQL = StrSQL & " and UnitID = " & val(RsDetails!UnitID & "")
         '    StrSQL = StrSQL & " and TableID Not In  " & mTableID & ")"
            
            rsDummy2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
       
       
            If Not rsDummy2.EOF Then
              ' If val(RsDetails!PartItemQty & "") = 0 Then
                    PartItemQty = val(rsDummy2!PartItemQty & "")
             '   Else
             '   PartItemQty = val(RsDetails!PartItemQty & "")
             '   End If
            '    If val(RsDetails!ForUnit & "") = 0 Then
                    ForUnit = val(rsDummy2!ForUnit & "")
              '  Else
              '      ForUnit = val(RsDetails!ForUnit & "")
              '  End If
                FGRows.TextMatrix(Num, FGRows.ColIndex("TableID")) = val(rsDummy2!TableID & "")
                    If mTableID = "" Then
               mTableID = "(" & FGRows.TextMatrix(Num, FGRows.ColIndex("TableID"))
            Else
                mTableID = mTableID & "," & FGRows.TextMatrix(Num, FGRows.ColIndex("TableID"))
            End If
                If val(RsDetails!lowering & "") = 0 Then
                    lowering = val(rsDummy2!lowering & "")
                Else
                    lowering = val(RsDetails!lowering & "")
                End If

                If val(RsDetails!increase & "") = 0 Then
                    increase = val(rsDummy2!increase & "")
                Else
                    increase = val(RsDetails!increase & "")
                End If

               ' If val(RsDetails!MethodCalc & "") = 0 Then
                    MethodCalc = val(rsDummy2!MethodCalc & "")
              '  Else
               ' MethodCalc = val(RsDetails!MethodCalc & "")
               ' End If

               
                

            End If
            If val(FGRows.TextMatrix(Num, FGRows.ColIndex("TableID"))) = 0 Then
                Num = Num
            End If
              rsDummy2.Close
            FGRows.TextMatrix(Num, FGRows.ColIndex("Ser")) = Num
            FGRows.TextMatrix(Num, FGRows.ColIndex("FlgX")) = IIf(IsNull(RsDetails("FlgX").value), "", Trim(RsDetails("FlgX").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("SpecID4")) = IIf(IsNull(RsDetails("SpecID4").value), "", Trim(RsDetails("SpecID4").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("SpecID3")) = IIf(IsNull(RsDetails("SpecID3").value), "", (RsDetails("SpecID3").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("SpecID2")) = IIf(IsNull(RsDetails("SpecID2").value), "", (RsDetails("SpecID2").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("Fullcode")) = IIf(IsNull(RsDetails("Fullcode").value), "", (RsDetails("Fullcode").value))
        
            FGRows.TextMatrix(Num, FGRows.ColIndex("SpecID1")) = IIf(IsNull(RsDetails("SpecID1").value), "", (RsDetails("SpecID1").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("ItemID")) = IIf(IsNull(RsDetails("ItemID").value), "", (RsDetails("ItemID").value))
            
            FGRows.TextMatrix(Num, FGRows.ColIndex("ItemID2")) = IIf(IsNull(RsDetails("ItemID2").value), "", (RsDetails("ItemID2").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("ItemCode2")) = IIf(IsNull(RsDetails("ItemCode2").value), "", (RsDetails("ItemCode2").value))
           If SystemOptions.UserInterface = ArabicInterface Then
                FGRows.TextMatrix(Num, FGRows.ColIndex("ItemName2")) = IIf(IsNull(RsDetails("ItemName2").value), "", (RsDetails("ItemName2").value))
            Else
                FGRows.TextMatrix(Num, FGRows.ColIndex("ItemName2")) = IIf(IsNull(RsDetails("ItemNamee2").value), "", (RsDetails("ItemNamee2").value))
            End If
            FGRows.TextMatrix(Num, FGRows.ColIndex("LineID")) = IIf(IsNull(RsDetails("LineID").value), "", (RsDetails("LineID").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID").value), "", (RsDetails("UnitID").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("PartItemQty")) = PartItemQty
            FGRows.TextMatrix(Num, FGRows.ColIndex("ForUnit")) = ForUnit
            FGRows.TextMatrix(Num, FGRows.ColIndex("MethodCalc")) = MethodCalc
            FGRows.TextMatrix(Num, FGRows.ColIndex("lowering")) = lowering
            FGRows.TextMatrix(Num, FGRows.ColIndex("Increase")) = increase

            
            FGRows.TextMatrix(Num, FGRows.ColIndex("itemcode")) = IIf(IsNull(RsDetails("ItemCode").value), "", (RsDetails("ItemCode").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("cost")) = IIf(IsNull(RsDetails("cost").value), "", (RsDetails("cost").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("Qty")) = IIf(IsNull(RsDetails("Qty").value), "", (RsDetails("Qty").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("TepQty")) = IIf(IsNull(RsDetails("TepQty").value), val(FGRows.TextMatrix(Num, FGRows.ColIndex("Qty"))), Trim(RsDetails("TepQty").value))
           If SystemOptions.UserInterface = EnglishInterface Then
               ' FgRows.Cell(flexcpData, Num, FgRows.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemNamee").value), "", (RsDetails("ItemNamee").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("unitname")) = IIf(IsNull(RsDetails("UnitNamee").value), "", (RsDetails("UnitNamee").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name1")) = IIf(IsNull(RsDetails("Namee1").value), "", (RsDetails("Namee1").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name2")) = IIf(IsNull(RsDetails("Namee2").value), "", (RsDetails("Namee2").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name3")) = IIf(IsNull(RsDetails("Namee3").value), "", (RsDetails("Namee3").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name4")) = IIf(IsNull(RsDetails("Namee4").value), "", (RsDetails("Namee4").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemNamee").value), "", (RsDetails("ItemNamee").value))

       Else
            'FgRows.Cell(flexcpData, Num, FgRows.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemName").value), "", (RsDetails("ItemName").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("unitname")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name1")) = IIf(IsNull(RsDetails("name1").value), "", (RsDetails("name1").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name2")) = IIf(IsNull(RsDetails("name2").value), "", (RsDetails("name2").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name3")) = IIf(IsNull(RsDetails("name3").value), "", (RsDetails("name3").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("name4")) = IIf(IsNull(RsDetails("name4").value), "", (RsDetails("name4").value))
         FGRows.TextMatrix(Num, FGRows.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemName").value), "", (RsDetails("ItemName").value))

       
    End If
            FGRows.TextMatrix(Num, FGRows.ColIndex("IsDeleted")) = IIf(IsNull(RsDetails("IsDeleted").value), 0, IIf((RsDetails("IsDeleted").value), -1, 0))
            FGRows.TextMatrix(Num, FGRows.ColIndex("IsAdd")) = IIf(IsNull(RsDetails("IsAdd").value), 0, (RsDetails("IsAdd").value))
           
           If val(RsDetails!Price & "") = 0 Then
'                CalcTotal Num
            Else
               FGRows.TextMatrix(Num, FGRows.ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", (RsDetails("Price").value))
                FGRows.TextMatrix(Num, FGRows.ColIndex("Total")) = IIf(IsNull(RsDetails("Total").value), "", (RsDetails("Total").value))
            End If
            FGRows.TextMatrix(Num, FGRows.ColIndex("OldPrice")) = IIf(IsNull(RsDetails("OldPrice").value), "", (RsDetails("OldPrice").value))
            
           
            
            FGRows.TextMatrix(Num, FGRows.ColIndex("Amout1")) = IIf(IsNull(RsDetails("Amout1").value), "", (RsDetails("Amout1").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("Amout2")) = IIf(IsNull(RsDetails("Amout2").value), "", (RsDetails("Amout2").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("Amout3")) = IIf(IsNull(RsDetails("Amout3").value), "", (RsDetails("Amout3").value))
            FGRows.TextMatrix(Num, FGRows.ColIndex("Amout4")) = IIf(IsNull(RsDetails("Amout4").value), "", (RsDetails("Amout4").value))

            
            If IIf(IsNull(RsDetails("IsDeleted").value), False, (RsDetails("IsDeleted").value)) Then
                FGRows.RowHidden(Num) = True
                mmmm = (RsDetails("ItemID").value)
                
            Else
                FGRows.RowHidden(Num) = False
            End If
                    
            
            RsDetails.MoveNext
           

        Next Num
      
        FGRows.AutoSize 0, FGRows.Cols - 1, False
    End If
    

End Sub
Function print_report3(Optional NoteSerial As String, Optional ByVal mOld As Boolean = False)
     
         Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

Dim cont1 As Double
Dim cont As Double
Dim Typ As Integer
Dim mPercent  As Double
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
 MySQL = ""


MySQL = MySQL & "  SELECT ProjectTotal  RentValuePayed,BDet_Qun Note_Value,BdQun ,BDet_Qun -BdQun     WaterPayed,TContract_CustID,"



MySQL = MySQL & "         Rent = ROUND((CASE ISNULL(TotalInstallPrice, 0)"
MySQL = MySQL & "                            WHEN 0 THEN 0"
MySQL = MySQL & "                            ELSE PaymentTotal /"
MySQL = MySQL & "                                     TotalInstallPrice"
                                 
                                 
MySQL = MySQL & "                                     End"
MySQL = MySQL & "         * 100),2),"
MySQL = MySQL & "  Servce  =( CASE ISNULL(BDet_Qun, 0)"
MySQL = MySQL & "                        WHEN 0 THEN 0"
MySQL = MySQL & "                        ELSE (BDet_Qun -BdQun) /"
MySQL = MySQL & "                                 BDet_Qun"
                                    
MySQL = MySQL & "                                 End"
MySQL = MySQL & "                              * 100),"
MySQL = MySQL & "         CusName  , TradingContractID NoteSerial1, ExpensesTotal, ReciptTotal RemaiValue,PaymentTotal TelandNetPayed,TotalInstallPrice Instrunce,"

MySQL = MySQL & "         ReciptTotal /"
               
MySQL = MySQL & "            CASE"
MySQL = MySQL & "                 WHEN IsNull(ProjectTotal,0) = 0 THEN 1"
MySQL = MySQL & "                 ELSE IsNull(ProjectTotal,0)"
MySQL = MySQL & "            End"
           
MySQL = MySQL & "            * 100 as InsurancePayed"
MySQL = MySQL & "  FROM   ("
MySQL = MySQL & "             SELECT Tbl_TradingContract.ProjectTotal,"
MySQL = MySQL & "                    SUM(ISNULL(Tbl_TradingContractDet.TContractDet_Qun, 0)) AS BDet_Qun,"
MySQL = MySQL & "                    BdQun = ISNULL("
MySQL = MySQL & "                        ("
MySQL = MySQL & "                            SELECT SUM(ISNULL(BDet_Qun, 0))"
MySQL = MySQL & "                            From Tbl_BusinessDialyDet"
MySQL = MySQL & "                                   RIGHT OUTER JOIN Tbl_BusinessDialy"
MySQL = MySQL & "                                        ON  Tbl_BusinessDialyDet.BDet_BD_ID = Tbl_BusinessDialy.ID"
MySQL = MySQL & "                            Where Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                                   AND Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                        ),"
MySQL = MySQL & "  0"
MySQL = MySQL & "                    ),"


MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,"
MySQL = MySQL & "                    TblCustemers.CusName,"
MySQL = MySQL & "                    Tbl_TradingContract.ID  AS TradingContractID,"
MySQL = MySQL & "                     TotalInstallPrice = (SELECT Sum(tcd.TContractDet_TotalInstallPrice)"
MySQL = MySQL & "                    FROM Tbl_TradingContractDet AS tcd Where TContractDet_TContractID = Tbl_TradingContract.ID) ,"

MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 5 OR NoteType = 3)"
MySQL = MySQL & "                    )                       AS ExpensesTotal,"
MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 5 )"
MySQL = MySQL & "                    )                       AS PaymentTotal,"


MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 4)"
MySQL = MySQL & "                    )                       AS ReciptTotal"
MySQL = MySQL & "             From Tbl_TradingContractDet"
MySQL = MySQL & "                    RIGHT OUTER JOIN Tbl_TradingContract"
MySQL = MySQL & "                         ON  Tbl_TradingContractDet.TContractDet_TContractID = Tbl_TradingContract.ID"
MySQL = MySQL & "                    LEFT OUTER JOIN TblCustemers"
MySQL = MySQL & "                         ON  Tbl_TradingContract.TContract_CustID = TblCustemers.CusID"
MySQL = MySQL & "             Where (1 = 1)"
If StrWhere = "" Then
     MySQL = MySQL & StrWhere
End If
If Option1(3) And Not mOld Then
         MySQL = MySQL & "         and IsNull(Tbl_TradingContract.NewMeasureNo,0) = 0"
End If
MySQL = MySQL & "             Group By"
MySQL = MySQL & "                    Tbl_TradingContract.ProjectTotal,"
MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,"
MySQL = MySQL & "                    TblCustemers.CusName,"
MySQL = MySQL & "                    Tbl_TradingContract.ID"
MySQL = MySQL & "                                         "
MySQL = MySQL & "         )                   T "


    

  If Option1(0) And Not mOld Then
        MySQL = MySQL & "          Where  T.TradingContractID In (Select Tbl_TradingContract.ID FROM Tbl_TradingContract Where IsNull(Tbl_TradingContract.Period,0) = 0 )"
    ElseIf Option1(1) And Not mOld Then
         MySQL = MySQL & "         Where T.TradingContractID  NOt In (Select IsNull(Tbl_BusinessDialy.TradingContractID,0) from Tbl_BusinessDialy)"
    ElseIf Option1(2) And Not mOld Then
         MySQL = MySQL & "         Where T.TradingContractID  NOt In (Select IsNull(Tbl_TransOrder.TradingContractID,0) from Tbl_TransOrder)"
    ElseIf Option1(3) And Not mOld Then
    
    ElseIf Option1(4) And Not mOld Then
    
        MySQL = MySQL & "         Where CASE IsNull(T.TotalInstallPrice,0) WHEN 0 THEN 0"
        MySQL = MySQL & "          Else"
        MySQL = MySQL & "            T.PaymentTotal / (CASE IsNull(T.TotalInstallPrice,0) WHEN 0  THEN 1 ELSE  T.TotalInstallPrice  End ) *100"
        MySQL = MySQL & "              END > 80"
    

ElseIf Option1(4) Then

    
    
    MySQL = MySQL & "         WHERE IsNull(T.ProjectTotal,0)- IsNull(T.ReciptTotal,0) > 0"
    
    
    If val(txtPercent(0)) > 0 Then
'        MySQL = MySQL & "  AND CONVERT(MONEY, BdQun) /"
'
'        MySQL = MySQL & "            CASE"
'        MySQL = MySQL & "                 WHEN CONVERT(MONEY, BDet_Qun) = 0 THEN 1"
'        MySQL = MySQL & "                 ELSE CONVERT(MONEY, BDet_Qun)"
'        MySQL = MySQL & "            End"
'
'        MySQL = MySQL & "            * 100"
'        MySQL = MySQL & "            >= " & val(txtPercent(0))
        
            MySQL = MySQL & " and Case IsNull(BDet_Qun, 0)"
            MySQL = MySQL & "       WHEN 0 THEN 0"
            MySQL = MySQL & "                  ELSE (BDet_Qun -BdQun) / BDet_Qun"
            MySQL = MySQL & "  END * 100"
            MySQL = MySQL & "            >= " & val(txtPercent(0))

    End If
    If val(txtPercent(1)) > 0 Then
    
        MySQL = MySQL & "  AND ReciptTotal /"
                   
        MySQL = MySQL & "            CASE"
        MySQL = MySQL & "                 WHEN IsNull(ProjectTotal,0) = 0 THEN 1"
        MySQL = MySQL & "                 ELSE IsNull(ProjectTotal,0)"
        MySQL = MySQL & "            End"
                   
        MySQL = MySQL & "            * 100"
        MySQL = MySQL & "            <= " & val(txtPercent(1))
    End If

End If

   
  


 
 


   If Option1(0) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant.rpt"
        End If
    ElseIf Option1(1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant1.rpt"
        End If
 
  ElseIf Option1(2) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant2.rpt"
        End If
 
   ElseIf Option1(3) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant3.rpt"
        End If
 
   ElseIf Option1(4) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant4.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant4.rpt"
        End If
 
   ElseIf Option1(5) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant5.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "InstallmentVendorAlarmContrant5.rpt"
        End If
 

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
'111
    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        Dim mTitle As String
        If Option1(0) Then
            mTitle = Option1(0).Caption
        ElseIf Option1(1) Then
            mTitle = Option1(1).Caption
        ElseIf Option1(2) Then
            mTitle = Option1(2).Caption
        ElseIf Option1(3) Then
            mTitle = Option1(3).Caption
        ElseIf Option1(4) Then
            mTitle = Option1(4).Caption
        ElseIf Option1(5) Then
            mTitle = Option1(5).Caption
        End If
        xReport.ParameterFields(4).AddCurrentValue mTitle
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'       If Not (IsNull(Me.Fromdate.value)) Then
'        xReport.ParameterFields(6).AddCurrentValue Fromdate.value
'       End If
'      If Not (IsNull(Me.todate.value)) Then
'        xReport.ParameterFields(7).AddCurrentValue todate.value
'      End If
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'   xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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



Public Sub FillGrid2(Optional StrWhere As String = "", Optional StrWhere2 As String = "", Optional ByVal mOld As Boolean = False)
Dim cont1 As Double
Dim cont As Double
Dim Typ As Integer
Dim mPercent  As Double
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset


 MySQL = ""


MySQL = MySQL & "  SELECT ProjectTotal ,PercentAlarm,BDet_Qun,BdQun,BDet_Qun -BdQun     Diff, "
MySQL = MySQL & "  ISNULL(ReciptTotal, 0) -(ISNULL(PercentAlarm, 0) * ISNULL(ReciptTotal, 0) / 100) as RepaymentRate ,"


MySQL = MySQL & "         PercentInstall = ROUND((CASE ISNULL(TotalInstallPrice, 0)"
MySQL = MySQL & "                            WHEN 0 THEN 0"
MySQL = MySQL & "                            ELSE PaymentTotal /"
MySQL = MySQL & "                                     TotalInstallPrice"

MySQL = MySQL & "                                     End"
MySQL = MySQL & "         * 100),2),"

MySQL = MySQL & "                                    ("
MySQL = MySQL & "                                       Case IsNull(BDet_Qun, 0)"
MySQL = MySQL & "                                            WHEN 0 THEN 0"
MySQL = MySQL & "                                            ELSE (BdQun) / BDet_Qun"
MySQL = MySQL & "                                       END * 100"
MySQL = MySQL & "                                    ) * T.ProjectTotal / 100 AS RealTotal,"


MySQL = MySQL & "  [PERCENT] =( CASE ISNULL(BDet_Qun, 0)"
MySQL = MySQL & "                        WHEN 0 THEN 0"
MySQL = MySQL & "                        ELSE (BdQun) /"
MySQL = MySQL & "                                 BDet_Qun"
                                    
MySQL = MySQL & "                                 End"
MySQL = MySQL & "                              * 100),"
 
        
MySQL = MySQL & "  TContract_CustID , "
MySQL = MySQL & "         CusName , TradingContractID, ExpensesTotal, ReciptTotal,PaymentTotal,TotalInstallPrice,"

MySQL = MySQL & "         ReciptTotal /"
               
MySQL = MySQL & "            CASE"
MySQL = MySQL & "                 WHEN IsNull(ProjectTotal,0) = 0 THEN 1"
MySQL = MySQL & "                 ELSE IsNull(ProjectTotal,0)"
MySQL = MySQL & "            End"
           
MySQL = MySQL & "            * 100 as Percent2,"

MySQL = MySQL & "         ReciptTotal /"
               
MySQL = MySQL & "            CASE"
MySQL = MySQL & "                 WHEN IsNull(ProjectTotal,0) = 0 THEN 1"
MySQL = MySQL & "                 ELSE IsNull(ProjectTotal,0)"
MySQL = MySQL & "            End"
           
MySQL = MySQL & "            * 100 as Percent3"
MySQL = MySQL & "  FROM   ("
MySQL = MySQL & "             SELECT Tbl_TradingContract.ProjectTotal+ IsNull(Tbl_TradingContract.Vat2,0)  as ProjectTotal,"
MySQL = MySQL & "                    SUM(ISNULL(Tbl_TradingContractDet.TContractDet_Qun, 0)) AS BDet_Qun,"
MySQL = MySQL & "                    BdQun = ISNULL("
MySQL = MySQL & "                        ("
MySQL = MySQL & "                            SELECT SUM(ISNULL(BDet_Qun, 0))"
MySQL = MySQL & "                            From Tbl_BusinessDialyDet"
MySQL = MySQL & "                                   RIGHT OUTER JOIN Tbl_BusinessDialy"
MySQL = MySQL & "                                        ON  Tbl_BusinessDialyDet.BDet_BD_ID = Tbl_BusinessDialy.ID"
MySQL = MySQL & "                            Where Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                                   AND Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                        ),"
MySQL = MySQL & "  0"
MySQL = MySQL & "                    ),"


MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,Tbl_TradingContract.PercentAlarm,"
MySQL = MySQL & "                    TblCustemers.CusName,"
MySQL = MySQL & "                    Tbl_TradingContract.ID  AS TradingContractID,"
MySQL = MySQL & "                     TotalInstallPrice = (SELECT Sum(tcd.TContractDet_TotalInstallPrice)"
MySQL = MySQL & "                    FROM Tbl_TradingContractDet AS tcd Where TContractDet_TContractID = Tbl_TradingContract.ID) ,"

MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 5 OR NoteType = 3)"
MySQL = MySQL & "                    )                       AS ExpensesTotal,"
MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 5 )"
MySQL = MySQL & "                    )                       AS PaymentTotal,"


MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 4)"
MySQL = MySQL & "                    )                       AS ReciptTotal"
MySQL = MySQL & "             From Tbl_TradingContractDet"
MySQL = MySQL & "                    RIGHT OUTER JOIN Tbl_TradingContract"
MySQL = MySQL & "                         ON  Tbl_TradingContractDet.TContractDet_TContractID = Tbl_TradingContract.ID"
MySQL = MySQL & "                    LEFT OUTER JOIN TblCustemers"
MySQL = MySQL & "                         ON  Tbl_TradingContract.TContract_CustID = TblCustemers.CusID"
MySQL = MySQL & "             Where (1 = 1)"
If StrWhere = "" Then
     MySQL = MySQL & StrWhere
End If

If Option1(3) And Not mOld Then
         MySQL = MySQL & "         and IsNull(Tbl_TradingContract.NewMeasureNo,0) = 0"
End If

MySQL = MySQL & "             Group By"
MySQL = MySQL & "                    Tbl_TradingContract.ProjectTotal,Tbl_TradingContract.Vat2,"
MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,"
MySQL = MySQL & "                    TblCustemers.CusName,"
MySQL = MySQL & "                    Tbl_TradingContract.ID,PercentAlarm"
MySQL = MySQL & "                                         "
MySQL = MySQL & "         )                   T "



  If Option1(0) And Not mOld Then
        MySQL = MySQL & "          Where  T.TradingContractID In (Select Tbl_TradingContract.ID FROM Tbl_TradingContract Where IsNull(Tbl_TradingContract.Period,0) = 0 )"
    ElseIf Option1(1) And Not mOld Then
         MySQL = MySQL & "         Where T.TradingContractID  NOt In (Select IsNull(Tbl_BusinessDialy.TradingContractID,0) from Tbl_BusinessDialy)"
    ElseIf Option1(2) And Not mOld Then
         MySQL = MySQL & "         Where T.TradingContractID  NOt In (Select IsNull(Tbl_TransOrder.TradingContractID,0) from Tbl_TransOrder)"
    ElseIf Option1(3) And Not mOld Then
    
    ElseIf Option1(4) And Not mOld Then
'
'        MySQL = MySQL & "         Where CASE IsNull(T.TotalInstallPrice,0) WHEN 0 THEN 0"
'        MySQL = MySQL & "          Else"
'        MySQL = MySQL & "            T.PaymentTotal / (CASE IsNull(T.TotalInstallPrice,0) WHEN 0  THEN 1 ELSE  T.TotalInstallPrice  End ) *100"
'        MySQL = MySQL & "              END > 80"
'2

   ' MySQL = MySQL & "         WHERE IsNull(T.ProjectTotal,0)- IsNull(T.ReciptTotal,0) > 0"
   
   
    
    MySQL = MySQL & "WHERE  ("
    MySQL = MySQL & "       ("
    MySQL = MySQL & "               Case IsNull(BDet_Qun, 0)"
    MySQL = MySQL & "                    WHEN 0 THEN 0"
    MySQL = MySQL & "                    ELSE (BdQun) / BDet_Qun"
    MySQL = MySQL & "               END * 100"
    MySQL = MySQL & "           ) * T.ProjectTotal / 100"
    MySQL = MySQL & ") >=     ISNULL(ReciptTotal, 0) -(ISNULL(PercentAlarm, 0) * ISNULL(ReciptTotal, 0) / 100)"
    
       
ElseIf Option1(5) Then
    MySQL = MySQL & "         WHERE IsNull(T.ProjectTotal,0)- IsNull(T.ReciptTotal,0) > 0"
    
    
    If val(txtPercent(0)) > 0 Then
        
        
         MySQL = MySQL & " and Case IsNull(BDet_Qun, 0)"
            MySQL = MySQL & "       WHEN 0 THEN 0"
            MySQL = MySQL & "                  ELSE (BDet_Qun -BdQun) / BDet_Qun"
            MySQL = MySQL & "  END * 100"
            MySQL = MySQL & "            >= " & val(txtPercent(0))

    End If
    If val(txtPercent(1)) > 0 Then
    
        MySQL = MySQL & "  AND ReciptTotal /"
                   
        MySQL = MySQL & "            CASE"
        MySQL = MySQL & "                 WHEN IsNull(ProjectTotal,0) = 0 THEN 1"
        MySQL = MySQL & "                 ELSE IsNull(ProjectTotal,0)"
        MySQL = MySQL & "            End"
                   
        MySQL = MySQL & "            * 100"
        MySQL = MySQL & "            <= " & val(txtPercent(1))
    End If

End If

   
  

Dim ActualTotal As Double
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      
      
      With Me.FG
       .Rows = 1
        .Clear flexClearScrollable
Dim j As Integer
Dim Notstr As String
j = 0
Notstr = ""
        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
               .TextMatrix(i, .ColIndex("ProjectTotal")) = rs!ProjectTotal & ""
              .TextMatrix(i, .ColIndex("BDet_Qun")) = rs!BDet_Qun & ""
              .TextMatrix(i, .ColIndex("BdQun")) = rs!BdQun & ""
              .TextMatrix(i, .ColIndex("TContract_CustID")) = rs!TContract_CustID & ""
              .TextMatrix(i, .ColIndex("CusName")) = rs!CusName & ""
              .TextMatrix(i, .ColIndex("TradingContractID")) = rs!TradingContractID & ""
              .TextMatrix(i, .ColIndex("ExpensesTotal")) = rs!ExpensesTotal & ""
              .TextMatrix(i, .ColIndex("ReciptTotal")) = rs!ReciptTotal & ""
              .TextMatrix(i, .ColIndex("Percent2")) = Round(val(rs!percent2 & ""), 2)
              .TextMatrix(i, .ColIndex("PercentAlarm")) = Round(val(rs!PercentAlarm & ""), 2)
              .TextMatrix(i, .ColIndex("RepaymentRate")) = Round(val(rs!RepaymentRate & ""), 2)
              
              
              .TextMatrix(i, .ColIndex("PaymentTotal")) = Round(val(rs!PaymentTotal & ""), 2)
              .TextMatrix(i, .ColIndex("TotalInstallPrice")) = Round(val(rs!TotalInstallPrice & ""), 2)
              .TextMatrix(i, .ColIndex("PercentInstall")) = Round(val(rs!PercentInstall & ""), 2)
              .TextMatrix(i, .ColIndex("RealTotal")) = Round(val(rs!RealTotal & ""), 2)
              
              
              If val(.TextMatrix(i, .ColIndex("BDet_Qun"))) <> 0 Then
                mPercent = Round(val(.TextMatrix(i, .ColIndex("BdQun"))) / val(.TextMatrix(i, .ColIndex("BDet_Qun"))) * 100, 2)
             End If
              .TextMatrix(i, .ColIndex("Percent")) = Round(val(rs!Percent & ""), 2)
              
            
         
                rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub


Private Sub cmdAdminPer_Click()
If user_id = UserAdminAll Or user_id = UserAdmin Then
    FillGrid3 , True
End If
End Sub

Private Sub CmdPrint_Click()


print_report3 , mOld
End Sub

Private Sub CmdHelp_Click()
          clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
toDate.value = Date
Fromdate.value = Date
End Sub

Private Sub CmdSave_Click()
Dim s As String
s = "Select * from TblProductLineDistribution "

saveGrid s, grd, "UserId", ID
End Sub

Private Sub Command2_Click()
    If grd.Row > 0 Then
        FillGridRows grd.Row
        FramRows.Visible = True
    End If
End Sub

Private Sub DBCboClientName_Change()
'    TxtSearchCode.text = ""
''
   ' Dim DefaultSalesPersonId As Integer
'    Dim fullcode As String

   ' GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

   ' TxtSearchCode.text = fullcode
 
End Sub
Function cahngelang()
    EleHeader.Caption = " Alarm Installment Assets "
    Me.Caption = EleHeader.Caption
    lbl(3).Caption = "Vendor"
    lbl(2).Caption = "Invoice Date"
    lbl(1).Caption = "Invoice No."
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
    Frame5.Caption = "Due Period"
    lbl(4).Caption = "Update All"
    lbl(5).Caption = "Vendor Invoice No."
    lbl(6).Caption = "Installment No."
    lbl(7).Caption = "Insatllment Value"
    
    chkPrintDirect.Caption = "Drirect Print"
   CmdHelp.Caption = "Clear"
   Cmd(5).Caption = "Refresh"
   Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("QsrID")) = "Installment Num"
    .TextMatrix(0, .ColIndex("Value")) = "Installment Value"
    .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
    .TextMatrix(0, .ColIndex("CusName")) = "Vendor"
    .TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill Num"
    .TextMatrix(0, .ColIndex("View")) = "View"
     .TextMatrix(0, .ColIndex("too")) = "Vendor invoice No."
    End With
    With grd
        .TextMatrix(0, .ColIndex("LineID22")) = "Serial"
        .TextMatrix(0, .ColIndex("ItemName")) = "Product name"
        .TextMatrix(0, .ColIndex("widtj")) = "Width "
        .TextMatrix(0, .ColIndex("hight")) = "Height "
        .TextMatrix(0, .ColIndex("Qty1")) = "Quantity "
        .TextMatrix(0, .ColIndex("Start")) = "Start "
        .TextMatrix(0, .ColIndex("End")) = "End "
        .TextMatrix(0, .ColIndex("PrintStiker")) = "Print the stiker "
        .TextMatrix(0, .ColIndex("Convert")) = "Conversion "
        .TextMatrix(0, .ColIndex("LineType")) = "Type of line "
        .TextMatrix(0, .ColIndex("ProductLineName")) = "Product line"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty "
        .TextMatrix(0, .ColIndex("UserName")) = "User Name "
        .TextMatrix(0, .ColIndex("TimeEnd")) = "Time "
        
        .TextMatrix(0, .ColIndex("IDDefCIT")) = "Order No"
        .TextMatrix(0, .ColIndex("NoteSerial13")) = "Sales Order"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name "
        .TextMatrix(0, .ColIndex("lowering")) = "lowering "
        .TextMatrix(0, .ColIndex("increase")) = "increase "
        .TextMatrix(0, .ColIndex("TimeStart")) = "Start time"
        .TextMatrix(0, .ColIndex("TimeEnd")) = "End time "
        .TextMatrix(0, .ColIndex("IsPrinted")) = "Printed"
        .TextMatrix(0, .ColIndex("BasedLineName")) = "Based Line"
        
        .TextMatrix(0, .ColIndex("DateStart")) = "Production start date"
        .TextMatrix(0, .ColIndex("DateEnd")) = "Production end date"
        

    End With
    
    With FGRows
        .TextMatrix(0, .ColIndex("IsAdd")) = "Add"
         .TextMatrix(0, .ColIndex("itemcode")) = "item code "
        .TextMatrix(0, .ColIndex("itemname")) = "item name "
        .TextMatrix(0, .ColIndex("unitname")) = "unit name "
        .TextMatrix(0, .ColIndex("cost")) = "cost"
        .TextMatrix(0, .ColIndex("FlgX")) = "FlgX "
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("name1")) = "Disc"
        .TextMatrix(0, .ColIndex("Amout1")) = "Value "
        .TextMatrix(0, .ColIndex("OldPrice")) = "Old Price"
        .TextMatrix(0, .ColIndex("Price")) = "Price "
        
        .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("name2")) = "Disc"
        .TextMatrix(0, .ColIndex("Amout2")) = "Amout2"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
   
    End With
    FramRows.Caption = "Rows"
    Command2.Caption = "Show raw materials"
    cmdAdminPer.Caption = "Show the powers of the admin"
    lbl(9).Caption = "Refresh"
    Cmd(1).Caption = "Refresh"
    TabMain.TabCaption(0) = "Agreement Alerts"
    TabMain.TabCaption(1) = "Asset Premium Alerts"
    TabMain.TabCaption(2) = "Production orders"
End Function



Public Sub FillGrid(Optional str As String)
Dim cont1 As Double
Dim cont As Double
Dim Typ As Integer
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 
 MySQL = ""
MySQL = MySQL & "   SELECT dbo.TblQestFexed.Value, dbo.TblQestFexed.Due_Date, dbo.TblQestFexed.QestID, dbo.notes_all.NoteID, dbo.TblCustemers.CusID, dbo.notes_all.NoteSerial1,"
MySQL = MySQL & "                     dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.notes_all.too, dbo.notes_all.NoteDateH, dbo.notes_all.NoteDate"
MySQL = MySQL & "  , inst_NO  FROM     dbo.notes_all INNER JOIN"
MySQL = MySQL & "                     dbo.TblQestFexed ON dbo.notes_all.NoteID = dbo.TblQestFexed.Ind LEFT OUTER JOIN"
MySQL = MySQL & "                     dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID"

 MySQL = MySQL & "    where 1 =1 "
  
 If IsNumeric(Text2.Text) And Text2.Text <> "" Then
    MySQL = MySQL & "  and  NoteSerial1 =  " & Text2.Text
 End If
  
 If IsNumeric(Text3.Text) And Text3.Text <> "" Then
    MySQL = MySQL & "   and too  =  " & Text3.Text
 End If
 
 'If IsNumeric(Text5.Text) And Text5.Text <> "" Then
 '   MySQL = MySQL & "  and  value  =  " & Text5.Text
 'End If
  
  If IsNumeric(Text4.Text) And Text4.Text <> "" Then
    MySQL = MySQL & "  and  Inst_No  =  " & Text4.Text
 End If
  
  
 If Not (IsNull(Me.Fromdate.value)) Then
 MySQL = MySQL + " and (dbo.TblQestFexed.Due_Date >='" & SQLDate(Fromdate.value) & "')"
 End If
 
 If Not (IsNull(Me.toDate.value)) Then
 MySQL = MySQL + " and (dbo.TblQestFexed.Due_Date <='" & SQLDate(toDate.value) & "')"
 End If

 If Not (IsNull(Me.DTPicker1.value)) Then
 MySQL = MySQL + " and (  notedate   = '" & SQLDate(toDate.value) & "')"
 End If

If Me.DBCboClientName.Text <> "" And val(Me.DBCboClientName.BoundText) <> 0 Then
MySQL = MySQL + "and notes_all.CusID =" & val(Me.DBCboClientName.BoundText) & ""
End If

MySQL = MySQL + "   order by  dbo.TblQestFexed.Ind "
  
  
  
  
  
Dim ActualTotal As Double
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable
Dim j As Integer
Dim Notstr As String
j = 0
Notstr = ""
        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
               .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value))
              .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("Value").value), 0, Round(rs.Fields("Value").value, 2)))
              .TextMatrix(i, .ColIndex("Due_Date")) = (IIf(IsNull(rs.Fields("Due_Date").value), Date, rs.Fields("Due_Date").value))
              .TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
              .TextMatrix(i, .ColIndex("QsrID")) = (IIf(IsNull(rs.Fields("inst_No").value), "", rs.Fields("inst_No").value))
              .TextMatrix(i, .ColIndex("too")) = (IIf(IsNull(rs.Fields("too").value), Date, rs.Fields("too").value))
            If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
           Else
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
           End If
         
        rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub

Private Sub Coloring()
    Dim i As Integer
 
     With grd

        For i = .FixedRows To .Rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 43) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 43) = vbWhite
            End If
 
        Next i

    End With
        If HidLowering Then
            Dim j As Long
             For i = 1 To grd.Rows - 1
                If val(grd.TextMatrix(i, grd.ColIndex("lowering"))) <> 0 Then
                    grd.Col = grd.ColIndex("widtj")
                    grd.Row = i
                    grd.CellBackColor = &HC0FFFF
                    grd.CellFontBold = True
                Else
                     grd.CellFontBold = False
'                    grd.Col = grd.ColIndex("widtj")
'                    grd.CellBackColor = vbRed
                 End If
            Next
        End If
        grd.Col = 1
        
End Sub



Public Sub FillGrid3(Optional str As String, Optional ByVal IsAdmin As Boolean = False)
Dim cont1 As Double
Dim cont As Double
Dim Typ As Integer
Dim mPercent  As Double
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
'    Dim rs As ADODB.Recordset
'
'    Set rs = New ADODB.Recordset
 If SystemOptions.UserInterface = EnglishInterface Then
     MySQL = ""
    
    
    MySQL = " SELECT DISTINCT T.ID,TblDefComItem.RecordDate,TblDefComItem.RecDate, IsPrinted = (           CASE ISNULL(PrintDate, '') WHEN  '' THEN 0 ELSE 1 END),T.IDDefCIT,T2.FormPrint,T.ProductLineID,T.LineID, T2.name   ProductLineName,T.SalesID,case  T2.IsBasicLine When 1 Then 1 Else 2 End as LineType,"
    MySQL = MySQL & "       T.GroupID, g.GroupNamee GroupName,tblItems.ItemNamee ItemName,TblItems.ItemCode,tblItems.lowering,tblItems.increase,"
    MySQL = MySQL & "       T.ItemNameID,T.UnitId,tu.UnitNamee UnitName,"
    'MySQL = MySQL & "       lowering = (select Top 1 lowering FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       increase = (select Sum( increase) FROM TblDefComItemDet DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID2  and IsNull(DD.IsDeleted,0) <> 1 ),"
    
    MySQL = MySQL & "       BasedLineName = (Select  Name From TblProductLine Where TblProductLine.Id = BaseProductLineID),BaseProductLineID,"
    'MySQL = MySQL & "       BaseProductLineID2 = (Select ProductLineId From TblGroupItemProductLine Where TblGroupItemProductLine.GroupID = T.GroupID),"
    
    MySQL = MySQL & "       BuiltinItemName = (select Top 1 tblItems.ItemNamee ItemName FROM TblDefComItemData DD Inner Join tblItems On DD.BuiltinItemID =tblItems.ItemID  Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID )"
Else
     MySQL = ""
    
    
    MySQL = " SELECT DISTINCT T.ID,TblDefComItem.RecordDate,TblDefComItem.RecDate,IsPrinted = (           CASE ISNULL(PrintDate, '') WHEN  '' THEN 0 ELSE 1 END),T.IDDefCIT,T2.FormPrint,T.ProductLineID,T.LineID, T2.name  ProductLineName,T.SalesID,case  T2.IsBasicLine When 1 Then 1 Else 2 End as LineType,"
    MySQL = MySQL & "       T.GroupID, g.GroupName ,tblItems.ItemName,TblItems.ItemCode,tblItems.lowering,tblItems.increase,"
    MySQL = MySQL & "       T.ItemNameID,T.UnitId,tu.UnitName,"
    
    MySQL = MySQL & "       ItemName2 = (select Top 1 ItemName FROM tblItems Inner join  TblDefComItemData DD On tblItems.ItemID = DD.ItemID2 Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       ItemName5 = (select Top 1 ItemName FROM tblItems Inner join  TblDefComItemData DD On tblItems.ItemID = DD.ItemID5 Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"

   ' MySQL = MySQL & "       lowering = (select Sum( lowering ) FROM TblDefComItemDet DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID2  and IsNull(DD.IsDeleted,0) <> 1 ),"
    'MySQL = MySQL & "       lowering = (select Top 1 lowering FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       increase = (select Sum( increase) FROM TblDefComItemDet DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID2  and IsNull(DD.IsDeleted,0) <> 1 ),"
    
    MySQL = MySQL & "       BasedLineName = (Select Name From TblProductLine Where TblProductLine.Id = BaseProductLineID),BaseProductLineID,"
    'MySQL = MySQL & "       BaseProductLineID2 = (Select ProductLineId From TblGroupItemProductLine Where TblGroupItemProductLine.GroupID = T.GroupID),"
    
    MySQL = MySQL & "       BuiltinItemName = (select Top 1 tblItems.ItemName FROM TblDefComItemData DD Inner Join tblItems On DD.BuiltinItemID =tblItems.ItemID  Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID )"
End If
    MySQL = MySQL & "       ,T.Qty,T.Qty1,T.PrintDate ,T.PrintTime,"
    
    
    MySQL = MySQL & "       CountItem2 = (select Top 1 CountItem2 FROM TblDefComItemData DD Where T.LineID = dd.LineID and DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       CountItem5 = (select Top 1 CountItem5 FROM TblDefComItemData DD Where T.LineID = dd.LineID and DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
   
    MySQL = MySQL & "       t.DateStart ,T2.StoreID,"
    MySQL = MySQL & "       Start =case IsNull(t.DateStart,0) When 0 then 0 else 1 end, "
    MySQL = MySQL & "       [End] =case IsNull(t.DateEnd,0) When 0 then 0 else 1 end "
    MySQL = MySQL & "       ,  t.DateEnd,TimeEnd,TimeStart,T.UserId,Users.UserName,TblCustemers.CusNamee CusName,TblDefComItem.NoteSerial13,"
    MySQL = MySQL & "       LineID22 =T.LineID ,"
    'MySQL = MySQL & "       LineID2 = (select Top 1 LineID FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    If HidLowering Then
        MySQL = MySQL & "       widtj = (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID)-IsNull(tblItems.lowering,0),"
        MySQL = MySQL & "       loweringWi =CASE ISNULL(tblItems.lowering2, 0) WHEN 0 THEN 0 ELSE    (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID)-IsNull(tblItems.lowering,0) -  IsNull(tblItems.lowering2,0) End,"
    Else
        MySQL = MySQL & "       widtj = (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID ),"
        MySQL = MySQL & "       loweringWi =CASE ISNULL(tblItems.lowering2, 0) WHEN 0 THEN 0 ELSE  (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID ) -  IsNull(tblItems.lowering2,0) End ,"
    End If
    MySQL = MySQL & "       hight = (select Top 1 hight FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID  ),"
    MySQL = MySQL & "       Remark = (select Top 1 Remark FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID AND DD.LineID = T.LineID  )"

    

MySQL = MySQL & " FROM   TblProductLineDistribution       AS T "
MySQL = MySQL & "       Inner JOIN TblProductLine  AS T2"
MySQL = MySQL & "            ON  T2.id = T.ProductLineID"

MySQL = MySQL & "       Inner JOIN TblDefComItem "
MySQL = MySQL & "            ON  TblDefComItem.id = T.IDDefCIT"
MySQL = MySQL & "       Left Outer JOIN TblCustemers "
MySQL = MySQL & "            ON  TblDefComItem.CusID= TblCustemers.CusID"

MySQL = MySQL & "       Inner JOIN tblItems"
MySQL = MySQL & "            ON  tblItems.ItemID = T.ItemNameID"
MySQL = MySQL & "       Inner JOIN TblUnites        AS tu"
MySQL = MySQL & "            ON  tu.UnitID = T.UnitID"
MySQL = MySQL & "       LEFT OUTER JOIN Groups           AS g"
MySQL = MySQL & "            ON  G.GroupID = T.GroupID"
MySQL = MySQL & "       LEFT OUTER JOIN TblUsers            AS Users"
MySQL = MySQL & "            ON  Users.UserId = T.UserId"

MySQL = MySQL & "       LEFT OUTER JOIN TblUsersProductLine            "
MySQL = MySQL & "            ON  TblUsersProductLine.ProductLineId = T2.Id"

MySQL = MySQL & "         Where  IsNull(tblItems.ItemType,0) = 0 and  IsNull(T.Qty,0) <> 0  and IsNull(IsNotShowAlarm,0) = 0 "
'T.ProductLineID In   (SELECT LineID FROM TblProductLineWorker "
'MySQL = MySQL & "         WHERE EmpId IN (SELECT EmpID FROM TblUsers WHERE UserID = " & user_id & ")) and "

mIsByAdmin = IsAdmin
If Not IsAdmin Then
    
    'MySQL = MySQL & "         (TblUsersProductLine.UserId = " & user_id & " Or TblUsersProductLine.UserId = " & UserAdminAll & " Or TblUsersProductLine.UserId = " & UserAdmin & "        OR ISNULL(TblUsersProductLine.UserId ,0) = 0)"
    
   'wael
    MySQL = MySQL & "     and    (TblUsersProductLine.UserId = " & user_id & " ) AND IsNull(TblUsersProductLine.typeLine,0) = 0   "
    MySQL = MySQL & "     and    IsNull(BaseProductLineID,0) Not In (Select  GG.ProductLineId From TblUsersProductLine GG Where IsNull(GG.typeLine,0) = 1 and IsNull(GG.ShowAlarm,0) = 1  AND GG.userid = " & user_id & " )"
    
    'MySQL = MySQL & "     and "
    'and IsNull(TblUsersProductLine.ShowAlarm,0) = 0 "
   
End If
MySQL = MySQL & " Order By TblDefComItem.RecordDate Desc"

loadgrid MySQL, grd, True, False
Coloring


'Exit Sub
i = 1
Dim j As Long
Dim ItemNameID As Long
Dim IDDefCIT As Long
Dim LineID As Long
Dim ProductLineName As String

Dim ItemNameID2 As Long
Dim IDDefCIT2 As Long
Dim LineID2 As Long
Dim ProductLineName2 As String
Dim mwidtj As Double
Dim mhight As Double

Dim mwidtj2 As Double
Dim mhight2 As Double
Again:
i = 1
Dim mProductLineID As Long
For i = 1 To grd.Rows - 1
    If i = 11 Then
        i = i
    End If
    ItemNameID = val(grd.TextMatrix(i, grd.ColIndex("ItemNameID")))
    IDDefCIT = val(grd.TextMatrix(i, grd.ColIndex("IDDefCIT")))
    LineID = val(grd.TextMatrix(i, grd.ColIndex("LineID")))
    ProductLineName = Trim(grd.TextMatrix(i, grd.ColIndex("ProductLineName")))
    mProductLineID = val(Trim(grd.TextMatrix(i, grd.ColIndex("ProductLineID"))))
    mwidtj = val(Trim(grd.TextMatrix(i, grd.ColIndex("widtj"))))
    mhight = val(Trim(grd.TextMatrix(i, grd.ColIndex("hight"))))
    For j = 1 To grd.Rows - 1
        ItemNameID2 = val(grd.TextMatrix(j, grd.ColIndex("ItemNameID")))
        IDDefCIT2 = val(grd.TextMatrix(j, grd.ColIndex("IDDefCIT")))
        LineID2 = val(grd.TextMatrix(j, grd.ColIndex("LineID")))
        ProductLineName2 = Trim(grd.TextMatrix(j, grd.ColIndex("ProductLineName")))
        mwidtj2 = val(Trim(grd.TextMatrix(j, grd.ColIndex("widtj"))))
        mhight2 = val(Trim(grd.TextMatrix(j, grd.ColIndex("hight"))))
        
'        If j = 11 Then
'            j = I
'        End If
        If mProductLineID <> 1 Then
            If ItemNameID2 = ItemNameID And IDDefCIT2 = IDDefCIT And ProductLineName = ProductLineName2 And mwidtj2 = mwidtj And mhight = mhight2 And LineID = LineID2 And j <> i Then
                'grd.RowHidden(j) = True
                grd.RemoveItem j
               ' MsgBox j & "  " & I
               GoTo Again
            End If
        End If
    Next j
Next i

End Sub






Private Sub Form_Load()
Dim StrSQL As String
   Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    'Dcombos.GetBranches Me.Dcbranch
    'Dcombos.GetFixedAssets DcbFixed
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " Select CusID,CusName From TblCustemers Where Type=2 or CustomerandVendor=1"
    Else
        StrSQL = " Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"
    End If

    fill_combo Me.DBCboClientName, StrSQL
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
       cahngelang
    End If

grd.ColComboList(grd.ColIndex("LineType")) = "#1;Œÿ √”«”Ì |#2;Œÿ ›—⁄Ï|"

grd.ColHidden(grd.ColIndex("lowering")) = Not HidLowering
grd.ColHidden(grd.ColIndex("increase")) = Not HidLowering

If user_id = UserAdminAll Or user_id = UserAdmin Then
    cmdAdminPer.Visible = True
Else
    cmdAdminPer.Visible = False
End If
Fromdate.value = Date
toDate.value = Date
DTPicker1.value = Date

'Fromdate.value = Null
'todate.value = Null
DTPicker1.value = Null


'FillGrid
FillGrid2
FillGrid3
If IsNumeric(Text1(0).Text) Then
    Timer1.interval = 1 * 60 * 100
End If
 
If IsNumeric(Text1(1).Text) Then
    Timer2.interval = 1 * 60 * 100
End If
 
If IsNumeric(Text1(2).Text) Then
    Timer2.interval = 1 * 60 * 100
End If
 
End Sub


Private Sub FromDate_Change()
FillGrid
End Sub

Private Sub grd_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date
    Select Case grd.ColKey(Col)
 

 Case "Start"
        If grd.ValueMatrix(Row, Col) Then
            mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
            mUserID = val(grd.TextMatrix(Row, grd.ColIndex("UserId")))
            grd.TextMatrix(Row, grd.ColIndex("DateStart")) = Date
            grd.TextMatrix(Row, grd.ColIndex("TimeStart")) = Time
            mTime = Time
            mDate = Date
    
       
            s = "Update TblProductLineDistribution Set UserId =  " & mUserID
            If Trim(grd.TextMatrix(Row, grd.ColIndex("TimeStart"))) <> "" Then
                s = s & " ,TimeStart = '" & Format(mTime, "hh:mm:ss") & "'"
            End If
            If Trim(grd.TextMatrix(Row, grd.ColIndex("DateStart"))) <> "" Then
                s = s & " , DateStart = " & SQLDate(mDate, True) & ""
            End If
            s = s & " Where Id = " & mIDD
            Cn.Execute s
            grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
        Else
            s = "Update TblProductLineDistribution Set UserId =  " & mUserID
            s = s & " ,TimeStart = Null"
            s = s & " , DateStart = Null"
            
            s = s & " Where Id = " & mIDD
            Cn.Execute s
        End If
        grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
    Case "End"
        If grd.ValueMatrix(Row, Col) Then
             mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
            mUserID = val(grd.TextMatrix(Row, grd.ColIndex("UserId")))
            grd.TextMatrix(Row, grd.ColIndex("DateEnd")) = Date
            grd.TextMatrix(Row, grd.ColIndex("TimeEnd")) = Time
            mTime = Time
            mDate = Date
            s = "Update TblProductLineDistribution Set UserId =  " & mUserID
           If Trim(grd.TextMatrix(Row, grd.ColIndex("TimeEnd"))) <> "" Then
                s = s & " ,TimeEnd = '" & Format(mTime, "hh:mm:ss") & "'"
           End If
           If Trim(grd.TextMatrix(Row, grd.ColIndex("DateEnd"))) <> "" Then
               s = s & " , DateEnd = " & SQLDate(mDate, True) & ""
           End If
           s = s & " Where Id = " & mIDD
           Cn.Execute s
           grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
        Else
            s = "Update TblProductLineDistribution Set UserId =  " & mUserID
            s = s & " ,TimeEnd = null"
            s = s & " , DateEnd = null           "
            s = s & " Where Id = " & mIDD
            Cn.Execute s
        End If
        grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
        Case "PrintStiker"
            grd.TextMatrix(Row, grd.ColIndex("PrintStiker")) = ""
          PrintStiker Row
    
          mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
          s = "Update TblProductLineDistribution Set UserId =  " & user_id
          s = s & " ,PrintTime = '" & Format(Time, "hh:mm:ss") & "'"
          s = s & " , PrintDate = " & SQLDate(Date, True) & ""
          s = s & " Where Id = " & mIDD
          Cn.Execute s
          FillGrid3
        Case "Convert"
            mType = val(grd.TextMatrix(Row, grd.ColIndex("LineType")))
            If Trim(grd.TextMatrix(Row, grd.ColIndex("PrintDate"))) <> "" Then
                If mType = 1 Then
                    TransferItems Row
                Else
                    TransferItems Row, True
                End If
                FillGrid3
            Else
               
            End If
        Case "ShowRows"
            FramRows.Visible = True
        End Select
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  On Error Resume Next
  Select Case grd.ColKey(Col)
 
    
    Case "TimeStart"
        grd.TextMatrix(Row, grd.ColIndex("TimeStart")) = Time
        grd.TextMatrix(Row, grd.ColIndex("DateStart")) = Date
        grd.TextMatrix(Row, grd.ColIndex("Start")) = 1
        grd_AfterEdit Row, grd.ColIndex("Start")
       ' FillGrid3
    Case "TimeEnd"
        grd.TextMatrix(Row, grd.ColIndex("TimeEnd")) = Time
        grd.TextMatrix(Row, grd.ColIndex("DateEnd")) = Date
        grd.TextMatrix(Row, grd.ColIndex("End")) = 1
        grd_AfterEdit Row, grd.ColIndex("End")
       ' FillGrid3
  Case "PrintStiker"
        Cancel = False

        grd.TextMatrix(Row, grd.ColIndex("PrintStiker")) = ""
        grd.TextMatrix(Row, grd.ColIndex("IsPrinted")) = 1
         PrintStiker Row
    
          mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
          s = "Update TblProductLineDistribution Set UserId =  " & user_id
          s = s & " ,PrintTime = '" & Format(Time, "hh:mm:ss") & "'"
          s = s & " , PrintDate = " & SQLDate(Date, True) & ""
          s = s & " Where Id = " & mIDD
          Cn.Execute s
          FillGrid3 "", mIsByAdmin
        Cancel = True
        
        
    Case "Convert"
        If val(grd.TextMatrix(Row, grd.ColIndex("IsConverted"))) = 1 Then Exit Sub
        mType = val(grd.TextMatrix(Row, grd.ColIndex("LineType")))
        If Trim(grd.TextMatrix(Row, grd.ColIndex("PrintDate"))) <> "" Then
            If mType = 1 Then
                TransferItems Row
            Else
                TransferItems Row, True
            End If
            FillGrid3
        End If
        On Error Resume Next
        grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = 1
        Cancel = True
    End Select
End Sub

Private Sub grd_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
   Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date
    Select Case grd.ColKey(Col)
   
    Case "DateStart"
        
        Frm.Index = 33
        Me.LngRow = Row
        Frm.show 1
        
        mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
        mUserID = val(grd.TextMatrix(Row, grd.ColIndex("UserId")))
        mDate = Trim(grd.TextMatrix(Row, grd.ColIndex("DateStart")))
        If grd.TextMatrix(Row, grd.ColIndex("TimeStart")) <> "" Then
            mTime = Trim(grd.TextMatrix(Row, grd.ColIndex("TimeStart")))
        End If
    
       
        s = "Update TblProductLineDistribution Set UserId =  " & mUserID
        If Trim(grd.TextMatrix(Row, grd.ColIndex("TimeStart"))) <> "" Then
            s = s & " ,TimeStart = '" & Format(mTime, "hh:mm:ss") & "'"
        End If
        If Trim(grd.TextMatrix(Row, grd.ColIndex("DateStart"))) <> "" Then
            s = s & " , DateStart = " & SQLDate(mDate, True) & ""
        End If
        s = s & " Where Id = " & mIDD
        grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
        Cn.Execute s
          Case "ShowRows"
            FillGridRows Row
            FramRows.Visible = True
    Case "DateEnd"
        Me.LngRow = Row
        Frm.Index = 34
        Frm.show 1
        
        mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
        mUserID = val(grd.TextMatrix(Row, grd.ColIndex("UserId")))
        mDate = Trim(grd.TextMatrix(Row, grd.ColIndex("DateEnd")))
        mTime = Trim(grd.TextMatrix(Row, grd.ColIndex("TimeEnd")))
   
        s = "Update TblProductLineDistribution Set UserId =  " & mUserID
        If Trim(grd.TextMatrix(Row, grd.ColIndex("TimeStart"))) <> "" Then
             s = s & " ,TimeEnd = '" & Format(mTime, "hh:mm:ss") & "'"
        End If
        If Trim(grd.TextMatrix(Row, grd.ColIndex("DateEnd"))) <> "" Then
            s = s & " , DateEnd = " & SQLDate(mDate, True) & ""
        End If
        s = s & " Where Id = " & mIDD
        Cn.Execute s
        grd.TextMatrix(Row, grd.ColIndex("IsConverted")) = ""
    Case "PrintStiker"
        PrintStiker Row
  
        mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
        s = "Update TblProductLineDistribution Set UserId =  " & user_id
        s = s & " ,PrintTime = '" & Format(Time, "hh:mm:ss") & "'"
        s = s & " , PrintDate = " & SQLDate(Date, True) & ""
        s = s & " Where Id = " & mIDD
        Cn.Execute s
        FillGrid3
    Case "Convert"
        mType = val(grd.TextMatrix(Row, grd.ColIndex("LineType")))
        If Trim(grd.TextMatrix(Row, grd.ColIndex("PrintDate"))) <> "" Then
            If mType = 1 Then
                TransferItems Row
            Else
                TransferItems Row, True
            End If
            FillGrid3
        End If
        
        
    End Select
End Sub


Private Sub TransferItems(ByVal Row As Long, Optional ByVal IsFinal As Boolean = False)
    Dim mQty2 As Double, mProductLineName  As String, ss As String
    If Row = 0 Then Exit Sub
   ' For i = 1 To grd.Rows - 1
        mQtyTotal = 0
          
        mQty2 = val(grd.TextMatrix(Row, grd.ColIndex("Qty")))
        If Trim(grd.TextMatrix(Row, grd.ColIndex("DateEnd"))) = "" Or Trim(grd.TextMatrix(Row, grd.ColIndex("TimeEnd"))) = "" Then
            MsgBox "·« Ì„ﬂ‰ «· ÕÊÌ· ﬁ»·  ”ÃÌ· »œ«Ì… Ê‰Â«Ì… «·«„—": Exit Sub
        End If
        
        If mQty2 = 0 Then MsgBox "·« ÌÊÃœ ﬂ„Ì«  ·Ì „  ÕÊÌ·Â«": Exit Sub
        If Not IsFinal Then
            If SystemOptions.CompilingBasedTable Then
                SaveProductionCompilingBasedTable True, mQty2, Row
            Else
            
                SaveItemsProduction3 True, mQty2, Row
            End If
            If mQtyTotal <> val(mQty2) Then
                mTotalSecond = Abs(val(mQty2) - mQtyTotal)
                If SystemOptions.CompilingBasedTable Then
                    SaveProductionCompilingBasedTable False, mQty2, Row
                Else
                    SaveItemsProduction3 False, mQty2, Row
                End If
            End If
        End If
        Dim mDateEnd As Date
        Dim mStoreId As Long
        mProductLineID = val(grd.TextMatrix(Row, grd.ColIndex("ProductLineID")))
        TxtTransSerial = val(grd.TextMatrix(Row, grd.ColIndex("IDDefCIT")))
        TxtNoteSerial13 = val(grd.TextMatrix(Row, grd.ColIndex("SalesID")))
        mItemNo = val(grd.TextMatrix(Row, grd.ColIndex("ItemNameID")))
        mUnitNo = val(grd.TextMatrix(Row, grd.ColIndex("UnitID")))
        mGroupID = val(grd.TextMatrix(Row, grd.ColIndex("GroupID")))
        mLineID = val(grd.TextMatrix(Row, grd.ColIndex("LineID")))
        mStoreId = val(grd.TextMatrix(Row, grd.ColIndex("StoreId")))
        
        mDateEnd = (grd.TextMatrix(Row, grd.ColIndex("DateEnd")))
        mBaseProductLineID = val(grd.TextMatrix(Row, grd.ColIndex("BaseProductLineID")))
        mProductLineName = Trim(grd.TextMatrix(Row, grd.ColIndex("BasedLineName")))
        
        
        s = "  update TblProductLineDistribution Set Qty = 0"
        s = s & "  Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        
        s = s & " and IDDefCIT = " & TxtTransSerial
        s = s & " and ProductLineID = " & mProductLineID
        s = s & " and IsNull(BaseProductLineID,0) = " & mBaseProductLineID
        Cn.Execute s
        
            If SystemOptions.DontCreateOut And IsFinal Then
                
                createVoucher CDbl(branch_id), 0, mDateEnd, 27, 0, val(user_id), 0, 2, CDbl(mStoreId), 0, 0, "”‰œ  ’—› »‰«¡ ⁄·Ì  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, val(TxtTransSerial), mItemNo, mUnitNo, mLineID, mQty2, mProductLineName
                        
            
            End If
            
        s = " SELECT name FROM TblProductLine where Id = " & IIf(mProdId = 0, mProductLineID, mProdId)
        Dim rsDummy As New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            MsgBox " „ «· ÕÊÌ· ⁄·Ï ÿ«Ê·… " & Trim(rsDummy!Name & "")
        End If
   ' Next
End Sub

Private Sub createVoucher(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, Optional invoice As Long = 0, Optional ByVal mItemNo As Long = 0, Optional ByVal mUnitNo As Long = 0, Optional ByVal mLineID As Long = 0, Optional ByVal mQty2 As Double = 0, Optional ByVal mProductLineName As String = "")
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim TxtTransSerial As Long
TxtTransSerial = invoice
 Dim rsTestQty As New ADODB.Recordset
 Dim rsTestQty2 As New ADODB.Recordset
'BillTOTAL = 0
CostTOTAL = 0
Dim ss As String

'Check
  'NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 10, 180, , 27)
    
    If Transaction_Type = 27 Then
            ss = "Select * from TblDefComItemDet Where IsNull(IsDeleted,0) = 0 and IDDefCIT = " & TxtTransSerial & " And ItemId2  = " & mItemNo & " And abs(IsNull(Qty,0)  - IsNull(Qtyout,0)) >= 1 "
            ' and LineID = " & mLineID
            rsTestQty.Open ss, Cn, adOpenKeyset, adLockOptimistic, adCmdText
            If rsTestQty.EOF Then
               ' ss = "Update Transaction_Details "
                ss = "Select * from TblDefComItemDet Where IsNull(IsDeleted,0) = 0 and IDDefCIT = " & TxtTransSerial & " And ItemId2  = " & mItemNo & " And abs(IsNull(Qty,0)  - IsNull(Qtyout,0)) = 0 "
                rsTestQty.Close
                rsTestQty.Open ss, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                Dim mLineRemark As String
                Do While Not rsTestQty.EOF
                    ss = "Select * from Transaction_Details "
                    ss = ss & " WHERE Transaction_ID IN ("
                    ss = ss & " SELECT Transaction_ID From Transactions AS t WHERE t.IDDefCIT = " & TxtTransSerial & ")"
                    ss = ss & " And Item_ID = " & val(rsTestQty!ItemID & "")
                    ss = ss & " And UnitID = " & val(rsTestQty!UnitID & "")
                    ss = ss & " And LineID = " & val(rsTestQty!ID & "")
                    ss = ss & " AND RemarksLine Not LIKE '%" & mProductLineName & "%'"
                    Set rsTestQty2 = New ADODB.Recordset
                    rsTestQty2.Open ss, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    If Not rsTestQty2.EOF Then
                        'mLineRemark = IIf(Trim(rsTestQty2!RemarksLine & "") = "", "", ",")
                        rsTestQty2!RemarksLine = Trim(rsTestQty2!RemarksLine & "") & "," & mProductLineName
                        rsTestQty2.update
                    End If
                   
            
       
                    rsTestQty.MoveNext
                    
                Loop
                Exit Sub
            End If
         NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 18, 240, , CInt(Transaction_Type), , CDbl(StoreId))              '’—› „Ê«œ Œ«„
    Else
    

        NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , CInt(Transaction_Type))    '’—› „Ê«œ Œ«„
    End If
                
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „Ê«œ Œ«„ ··«‰ «Ã  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " You can not add a raw material bond to a new production because you have exceeded the limit on which you have selected the bonds ": Exit Sub
                    End If
            
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
 
   If Trim(StoreId) = 0 Then
         MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
   End If
  
  
 
           'CostAccount = get_account_code_branch(137, CInt(BranchID))
           CostAccount = get_account_code_branch(1, CInt(BranchID))
        
            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›…   «·„»Ì⁄«   ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If

             Exit Sub
              End If
              
              

    StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
      If StoreAccount = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section", vbCritical
                End If
           Exit Sub
            End If
          Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double

 'end Check

        
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
Transaction_serial = NoteSerial1
'    If Transaction_Type <> 19 Then
'        TXTTransactionID1.Text = Transaction_ID
'        TxtNoteSerial11.Text = NoteSerial1
'    Else
'        TXTTransactionID5.Text = Transaction_ID
'        TxtNoteSerial15.Text = NoteSerial1
'
'    End If
Dim s As String
Dim mCust As Long
Dim rsDummyChkCust As New ADODB.Recordset
Dim rsDummy As ADODB.Recordset
sql = "Select * from TblCustemers Where CusId = " & CusID

rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
If rsDummyChkCust.EOF Then
    sql = "Select Top 1 CusId from TblCustemers "
    rsDummyChkCust.Close
    rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
    CusID = val(rsDummyChkCust!CusID & "")
End If
        
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,InvoiceOrderNo,"
 sql = sql & " TransactionComment,IDDefCIT )"
 
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & BoxID & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & Transaction_Type & " ,"
sql = sql & " 2 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & CusID & " ,"
sql = sql & " " & StoreId & " ,"
sql = sql & " 0 ,"
sql = sql & " " & Emp_id & " ," & val(TxtTransSerial) & ","
 sql = sql & "'" & TransactionComment & "'," & val(TxtTransSerial) & ")"
 

         Cn.Execute sql
 
s = "Select MaxNo2,BranchID,UserID From TblDefComItem Where Id = " & val(TxtTransSerial)
Dim mMaxNo As String
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    mMaxNo = Trim(rsDummy!MaxNo2 & "")
End If
Dim mTotal As Double
mTotal = 0
 
        Dim RSTransDetails As New ADODB.Recordset
     Set rsDummy = New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Transaction_Type = 19 Then
            
       
             
            s = "Select * from TblDefComItemData Where IDDefCIT = " & TxtTransSerial & " And ItemId  = " & mItemNo & " and UnitId = " & mUnitNo & " "
            'and LineID = " & mLineID
            rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
           
                
                
                Do While Not rsDummy.EOF
                
                        RSTransDetails.AddNew
                        RSTransDetails("Transaction_ID").value = Transaction_ID
                     
                        RSTransDetails("ColorID").value = 1
                        RSTransDetails("ItemSize").value = 1
                        RSTransDetails("ClassId").value = 1
                        RSTransDetails("Item_ID").value = val(rsDummy!ItemID & "")
                        RSTransDetails("UnitID").value = val(rsDummy!UnitID & "")
                       RSTransDetails("SHOWQTY").value = val(rsDummy!Qty & "")
                       RSTransDetails("showPrice").value = val(rsDummy!Price & "")
                       RSTransDetails("LineID").value = val(rsDummy!ID & "")
                      
                      
        
                
                    LngCurItemID = val(rsDummy!ItemID & "")
                    LngUnitID = val(rsDummy!UnitID & "")
                    DblQty = val(rsDummy!Qty & "")
                    costPrice = val(rsDummy!cost & "")
               '     costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
          ' costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
         'costPrice = 20
          ' CostTOTAL = CostTOTAL + costPrice * DblQty
          
                    ' FG2.TextMatrix(RowNum, FG2.ColIndex("cost")) = costPrice
                          
                  'RSTransDetails("ShowPrice").value = costPrice
                  RSTransDetails("showPrice").value = Round(costPrice / IIf(DblQty <> 0, DblQty, 1), 3)
                 RSTransDetails("ShowQty").value = DblQty
                            
                  
        
                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                    Set RsUnitData = New ADODB.Recordset
                    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                'fg2.TextMatrix(RowNum, fg2.ColIndex("Price")) = 0
        
                    If Not (rs.BOF Or rs.EOF) And Not RsUnitData.EOF Then
         
                        RSTransDetails("QtyBySmalltUnit").value = IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
                        RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                          RSTransDetails("Price").value = Round(costPrice / RSTransDetails("QtyBySmalltUnit").value, 3)
                    
                    End If
                    RSTransDetails("CostPrice").value = val(rsDummy!Price & "")
                    
                    CostTOTAL = CostTOTAL + (val(Round(val(RSTransDetails("showPrice").value) / RSTransDetails("QtyBySmalltUnit").value, 3)) * DblQty)
                    
                    RSTransDetails.update
                    

            rsDummy.MoveNext
        Loop
    Else
    
    
            
            s = "Select * from TblDefComItemDet Where IsNull(IsDeleted,0) = 0 and IDDefCIT = " & TxtTransSerial & " And ItemId2  = " & mItemNo & " And abs(IsNull(Qty,0)  - IsNull(Qtyout,0)) >= .005 "
            ' and LineID = " & mLineID
            rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
            Dim rsDummy3 As New ADODB.Recordset
                Dim rsDummy2 As New ADODB.Recordset
             Dim lowering As Double, increase As Double, MethodCalc As Double, ForUnit As Double, PartItemQty As Double, mwidtj As Double, mhight As Double, Qty As Double, mFlgX As Double, mQtyGrid As Double, thickness As Double, mDiameter As Double
             mQtyGrid = mQty2
            Do While Not rsDummy.EOF
                If (rsDummy!ItemID = 35 And val(rsDummy!Qty & "") = 1) Or (rsDummy!ItemID = 20 And val(rsDummy!Qty & "") = 31000) Then
                    s = s
                End If
                s = "Select * from TblDefComItemData Where ItemId  = " & mItemNo & " And IDDefCIT = " & TxtTransSerial & " and LineID = " & val(rsDummy!LineID & "")
                Set rsDummy3 = New ADODB.Recordset
                rsDummy3.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy3.EOF Then
                    mwidtj = val(rsDummy3!widtj & "")
                    mhight = val(rsDummy3!hight & "")
                    thickness = val(rsDummy3!thickness & "")
                    mDiameter = val(rsDummy3!Diameter & "")
                    
                End If
                
                StrSQL = "select    dbo.TblItemsParts.PartItemQty ,TableID, ForUnit,    TblItemsParts.lowering,TblItemsParts.increase,MethodCalc from TblItemsParts"
                StrSQL = StrSQL & " Where dbo.TblItemsParts.ItemID = " & val(rsDummy!ItemID2 & "")
                StrSQL = StrSQL & " and PartItemID = " & val(rsDummy!ItemID & "")
                StrSQL = StrSQL & " and UnitID = " & val(rsDummy!UnitID & "")
                If val(rsDummy!TableID & "") <> 0 Then
                    StrSQL = StrSQL & " and TableID = " & val(rsDummy!TableID)
                End If
                Set rsDummy2 = New ADODB.Recordset
                rsDummy2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If val(rsDummy!ItemID & "") = 20 Then
                    Qty = Qty
                End If
                If Not rsDummy2.EOF Then
                  
                    PartItemQty = val(rsDummy2!PartItemQty & "")
                    Qty = PartItemQty
                    mFlgX = PartItemQty
                    ForUnit = val(rsDummy2!ForUnit & "")
                    If val(ForUnit) = 0 Then ForUnit = 1
                    If val(rsDummy!lowering & "") = 0 Then
                        lowering = val(rsDummy2!lowering & "")
                    Else
                        lowering = val(rsDummy!lowering & "")
                    End If

                    If val(rsDummy!increase & "") = 0 Then
                        increase = val(rsDummy2!increase & "")
                    Else
                        increase = val(rsDummy!increase & "")
                    End If

              
                    MethodCalc = val(rsDummy2!MethodCalc & "")
                   ' mFlgX = val(rsDummy2!FlgX & "")
'                    mFlgX = val(rsDummy2!FlgX & "")
'                    mFlgX = val(rsDummy2!FlgX & "")
                 Else
                    MethodCalc = 0
                    mFlgX = 1
                 End If
                
                
                    If MethodCalc = 1 Then 'ﬂ„Ì…
                        
                    ElseIf MethodCalc = 2 Then '⁄—÷
                      Qty = ((val(mwidtj) / ForUnit) * Qty) - lowering
                        mFlgX = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                    ElseIf MethodCalc = 3 Then 'ÿÊ·
                    Qty = ((val(mhight) / ForUnit) * Qty) - lowering
                    
                        mFlgX = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                    
                     ElseIf MethodCalc = 4 Then 'ÿÊ·+⁄—÷
                     Qty = ((val(mwidtj) + val(mhight)) / ForUnit * Qty) - lowering
                      mFlgX = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                      ElseIf MethodCalc = 5 Then 'ÿÊ·*⁄—÷
                               Qty = ((val(mwidtj) * val(mhight)) / ForUnit * Qty) - lowering
                                mFlgX = Round(Qty, 2)
                            mQty = Round(mFlgX * val(mQtyGrid), 2)
                   ElseIf MethodCalc = 6 Then ' «·ÿÊ· ·ﬂ· ⁄—÷
                        Qty = ((val(mhight) / ForUnit) - lowering) * Qty * val(mwidtj)
                  
                      '.TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                     ElseIf MethodCalc = 7 Then ' «·⁄—÷ ·ﬂ· ÿÊ·
                    
                     Qty = ((val(mwidtj) / ForUnit * Qty) * val(mhight)) - lowering   ' ((val(mwidtj) +  / ForUnit * Qty) - lowering
                      mFlgX = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                     
                     
                       
                         
                        
                     ElseIf MethodCalc = 8 Then ' «·«— ›«⁄ * «·⁄—÷ * ÿÊ·
                    Qty = (((val(mwidtj) * val(mhight) * val(mLength))) / ForUnit * Qty) - lowering
                     
                      mFlgX = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                                           
                    ElseIf MethodCalc = 9 Then 'ÿÊ·+⁄—÷
                            Qty = (val(mhight) * 3.14 * ((val(mDiameter) / 2) ^ 2) / ForUnit * Qty) - lowering
                                
                    ElseIf MethodCalc = 10 Then 'ÿÊ·+⁄—÷
                            Qty = (((val(mwidtj) * val(mhight) * val(thickness))) / ForUnit * Qty) - lowering
                                   
                        End If
                    If MethodCalc <> 1 Then
                        If val(Qty) = 0 Then
                          '.TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        Qty = Round(mFlgX * val(mQtyGrid), 2)
                        End If
                            
                
                    
                    Else
                    '.TextMatrix(i, .ColIndex("FlgX")) = Qty
                    Qty = mFlgX * val(mQtyGrid)
                            
                            ' .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
                          
                        End If
            '    End If
                
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1

              ' RSTransDetails("showPrice").value = IIf((fg.TextMatrix(RowNum, fg.ColIndex("Price")) = ""), Null, val(fg.TextMatrix(RowNum, fg.ColIndex("Price"))))
                RSTransDetails("Item_ID").value = val(rsDummy!ItemID & "")
                RSTransDetails("UnitID").value = val(rsDummy!UnitID & "")
                RSTransDetails("SHOWQTY").value = Qty
                RSTransDetails("LineID").value = val(rsDummy!ID & "")
                RSTransDetails!RemarksLine = "Œÿ «‰ «Ã :" & mProductLineName
'                RSTransDetails("showPrice").value = val(rsDummy!Price & "")
                      
            LngCurItemID = val(rsDummy!ItemID & "")
            LngUnitID = val(rsDummy!UnitID & "")
            DblQty = Qty
            costPrice = val(rsDummy!cost & "")
                          '«·ÊÕœ« 
           
            Dim mIsFromMix As Boolean
             costPrice = GetCostFromMix2(1, LngCurItemID, mItemNo, LngUnitID, mMaxNo)
             
             If costPrice = 0 Then
                'costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill, val(Me.TXTTransactionID2.Text), LngUnitID, val(Me.DCboStore2Name.BoundText))
                mIsFromMix = False
                costPrice = val(rsDummy!cost & "")
            Else
                mIsFromMix = True
             '   getItemCostData XPDtbBill.value, CLng(LngCurItemID), val(DCboStore2Name.BoundText), val(Me.TXTTransactionID2.Text), OldQty, OldCost, NewQty, NewCost
             End If
  
            
            
       '     costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
             ' costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
            'costPrice = 20
              CostTOTAL = CostTOTAL + costPrice * DblQty
             mTotal = costPrice + mTotal
             
            If mIsFromMix Then
                
            Else
                costPrice = costPrice '* DblQty
            End If
            ' FG.TextMatrix(RowNum, FG.ColIndex("cost")) = costPrice
                  
          RSTransDetails("ShowPrice").value = costPrice
          
         RSTransDetails("ShowQty").value = DblQty
                    
          

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
       ' fg.TextMatrix(RowNum, fg.ColIndex("Price")) = 0

            If Not RsUnitData.EOF Then
 
                RSTransDetails("QtyBySmalltUnit").value = DblQty ' IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
                RSTransDetails("Quantity").value = DblQty ' RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                 RSTransDetails("Price").value = Round(costPrice / RSTransDetails("QtyBySmalltUnit").value, 3)
            RSTransDetails("Price").value = costPrice
            End If
            RSTransDetails("CostPrice").value = costPrice
            
 
            
                RSTransDetails.update
            dd = val(rsDummy!Qty & "")
            rsDummy!QtyOut = val(rsDummy!QtyOut & "") + DblQty
            rsDummy.update
NextRow:
                  rsDummy.MoveNext
        Loop
    
  End If
             UpdateTransactionsCost CStr(Transaction_ID)
             
'Exit Sub
 
NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If Transaction_Type = 27 Then
    CreateNotes NoteID, Transaction_Date, CInt(BranchID), 240, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, ToHijriDate(Transaction_Date)
Else
    CreateNotes NoteID, Transaction_Date, CInt(BranchID), 180, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ›« Ê—… „»Ì⁄«  —ﬁ„ " & TxtNoteSerial13, ToHijriDate(Transaction_Date)
End If

'TxtNoteSerial11
'***********************
'         If Transaction_Type = 19 Then
'            StrSQL = "UPDATE TblDefComItem SET  TransactionID5=" & val(Transaction_ID) & ",  NoteSerial15='" & NoteSerial1 & "' WHERE ID  =" & val(TxtTransSerial)
'            Cn.Execute StrSQL
'
'            StrSQL = "UPDATE Transactions SET  Nots=" & val(TXTTransactionID3) & ",BillBasedOn =2,nots2 = '" & Trim(TxtNoteSerial13.Text) & "',Closed = 1   WHERE Transaction_ID  =" & val(TXTTransactionID5)
'            Cn.Execute StrSQL
'
'        Else
'            StrSQL = "UPDATE TblDefComItem SET  TransactionID1=" & val(Transaction_ID) & ",  NoteSerial11='" & NoteSerial1 & "' WHERE ID  =" & val(TxtTransSerial)
'
'            Cn.Execute StrSQL
'            TxtNoteSerial1 = NoteSerial1
'        End If
'***********************

  CREATE_VOUCHER_GE1 Transaction_ID, NoteSerial1, "", NoteID, branch_id, StoreId, Transaction_Date, 0, CInt(invoice), mItemNo, mUnitNo, mLineID
       
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  'MsgBox " „   «·‰ﬁ·"
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:

End Sub

 
Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreId As Double, Transaction_Date As Date, BoxID As Double, Optional invoice As Integer = 0, Optional ByVal mItemNo As Long = 0, Optional ByVal mUnitNo As Long = 0, Optional ByVal mLineID As Long = 0)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim Line1 As Double
    Dim Line2 As Double
    Dim OtherInformation As New ClsGLOther
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim s As String
    Dim mCust As Long
    
    Dim rsDummy As ADODB.Recordset
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = CostTOTAL
 
    If SngTemp > 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
OtherInformation.NextAccount_Code = get_store_Account(val(StoreId), "Account_Code")
        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            Dim UseCustomerAcc As Integer

    
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
   

            DebitAccount = StrTempAccountCode
    
            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & TxtNoteSerialV & "  " & TxtBillComment
            Else
                StrTempDes = "Issue Voucher No.  " & TxtNoteSerialV & "  " & TxtBillComment
            End If

            Line1 = setfoxy_Line
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
            '«·„Œ“Ê‰ ›Ì «·›—⁄
            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                                    MsgBox "The inventory cost calculation in the branch is not specified for this process", vbCritical
                End If
                    GoTo ErrTrap
         
                End If
            End If
        
           
                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
          

            CreditAccount = StrTempAccountCode
    
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & TxtNoteSerialV & "  " & TxtBillComment
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerialV & "  " & TxtBillComment
            End If
    
            LngDevNO = LngDevNO + 1
            Line2 = setfoxy_Line

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            
     'salimhere
     If invoice = 0 Then '« «Ã
     Account_Code_dynamic = get_account_code_branch(37, CInt(BranchID))
        Else
        
        Account_Code_dynamic = get_account_code_branch(1, my_branch) '„»Ì⁄« 
        End If
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·«‰ «Ã ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                         MsgBox "The production cost calculation is not determined in the section for this process", vbCritical
                    End If
                    GoTo ErrTrap
         
                End If
            End If

           
            StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
          
            DebitAccount = StrTempAccountCode
            
            Line1 = setfoxy_Line

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & TxtNoteSerialV
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerialV
            End If
    
            LngDevNO = LngDevNO + 1
       Dim project_id As Integer
'        project_id = IIf(Me.DcbProject.BoundText = "", 0, Me.DcbProject.BoundText)
             If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
            SngTemp = CostTOTAL

            
            Account_Code_dynamic = get_store_Account(val(StoreId), "Account_Code")
            
        
            If Account_Code_dynamic = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section  ", vbCritical
                End If
                
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            CreditAccount = StrTempAccountCode
OtherInformation.NextAccount_Code = DebitAccount
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & TxtNoteSerialV & "  " & TxtBillComment
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerialV & "  " & TxtBillComment
            End If

            Line2 = setfoxy_Line
         
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single
            s = "Select * from TblDefComItemDet Where IsNull(IsDeleted,0) = 0 and IDDefCIT = " & invoice & " And ItemId2  = " & mItemNo & "  "
            'and LineID = " & mLineID
            rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
           
                
                
                Do While Not rsDummy.EOF

                                ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_in_branch(rsDummy!itemcode & "", val(my_branch), 1)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   ﬂ·›… ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = val(rsDummy!Price & "") * val(rsDummy!Qty & "")
                        
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & TxtNoteSerialV
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerialV
                        End If
    
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , , , , , , , , , , val(branch_id)) = False Then
                            GoTo ErrTrap
                        End If
    
                    
NextRow2:
                rsDummy.MoveNext
                Loop

            

             s = "Select * from TblDefComItemDet Where IsNull(IsDeleted,0) = 0 and IDDefCIT = " & invoice & " And ItemId2  = " & mItemNo & " "
             'and LineID = " & mLineID
            rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
           
                
                
                Do While Not rsDummy.EOF


                
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(rsDummy!itemcode & "", CInt(StoreId), 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        
                        line_value = val(rsDummy!Price & "") * val(rsDummy!Qty & "")
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & TxtNoteSerialV & "  " & TxtBillComment
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerialV & "  " & TxtBillComment
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Transaction_Date, user_id, Transaction_ID, , , , , , , , , , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                        End If
    
NextRow:
                            rsDummy.MoveNext
                Loop


        End If

        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If
   ' ute StrSQL
ErrTrap:
End Function


 Private Function GetCostFromMix2(ByVal mRow As Long, ByVal mItemNo As Long, ByVal mItemNo2 As Long, ByVal mUnitId As Long, ByVal mMaxNo As String) As Double
    If Trim(mMaxNo) = "" Then Exit Function

    
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    
        s = " SELECT tdcid.cost,tdcid.Price"
    s = s & " FROM   TblDefComItemDet AS tdcid"
    s = s & "        INNER JOIN TblDefComItemData"
    s = s & "             ON  tdcid.IDDefCIT = TblDefComItemData.IDDefCIT"
    s = s & "             AND tdcid.ItemID2 = TblDefComItemData.ItemID"
    s = s & "        RIGHT OUTER JOIN TblDefComItem"
    s = s & "             ON  TblDefComItemData.IDDefCIT = TblDefComItem.ID"
    s = s & " Where MaxNo = N'" & Trim(mMaxNo) & "'"
    s = s & " AND tdcid.itemId= " & mItemNo
    s = s & " AND tdcid.UnitID =" & mUnitId
    s = s & " AND tdcid.ItemID2 =" & mItemNo2
    
    s = " SELECT TblDefComItemData.cost,TblDefComItemData.Price FROM TblDefComItemData"
    s = s & " Inner Join"
    s = s & " TblDefComItem"
    s = s & " ON TblDefComItem.ID = TblDefComItemData.IDDefCIT"
    s = s & " Where MaxNo = N'" & Trim(mMaxNo) & "'"
    s = s & " AND TblDefComItemData.ItemID = " & mItemNo
    s = s & " AND TblDefComItemData.UnitID =" & mUnitId
    
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        GetCostFromMix2 = val(rsDummy!cost & "")
        If val(rsDummy!cost & "") <> 0 Then
          '  Fg.TextMatrix(mRow, Fg.ColIndex("cost")) = val(rsDummy!cost & "")
            GetCostFromMix2 = val(rsDummy!cost & "")
        End If
        'If val(rsDummy!Price & "") <> 0 Then
        '    Fg.TextMatrix(mRow, Fg.ColIndex("Price")) = val(rsDummy!Price & "")
        'End If
        
    End If
    


End Function

   
  

Private Sub SaveItemsProduction(ByVal IsFirst As Boolean, ByVal mQty22 As Double, ByVal mRow As Long)
    Dim TxtTransSerial As Long
    Dim TxtNoteSerial13 As Long
    
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    Dim mItemNo As Long, mUnitNo As Long, mGroupID As Long
    mItemNo = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
    mUnitNo = val(grd.TextMatrix(mRow, grd.ColIndex("UnitID")))
    mGroupID = val(grd.TextMatrix(mRow, grd.ColIndex("GroupID")))
    mLineID = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
    mProductLineID = val(grd.TextMatrix(mRow, grd.ColIndex("ProductLineID")))
    TxtTransSerial = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
    TxtNoteSerial13 = val(grd.TextMatrix(mRow, grd.ColIndex("SalesID")))
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
        s = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        End If
        If mCount = 0 Then
            
            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " and IDDefCIT = " & TxtTransSerial
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(mUnitNo)
        s = s & " and ItemNameID = " & val(mItemNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " Where IsNull(IsBasicLine,0) = 0"
        s = s & " Group BY ItemNameID,UnitID,T2.ID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "

        Dim isFirstTime As Boolean
'        RsDataLine.Close
        RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If RsDataLine.EOF Then
            RsDataLine.Close
            isFirstTime = True
            s = "SELECT *,Qty = 0 FROM TblProductLine "
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
        Do While Not RsDataLine.EOF
            i = i + 1
            If i = 1 Then
                mQtyCompare = val(RsDataLine!Qty & "")
            End If
            If IsFirst Then
                If isFirstTime Then
                    mPart = Round(val(mQty22) / mCount)
                    mQtyNew = mPart
                Else
                    mQtyNew = mQtyCompare - val(RsDataLine!Qty & "")
                End If
                
                    If i = mCount Then
                        If isFirstTime And Not IsFirst Then
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                        End If
                    End If
'                    If i = mCount Then
'                        If Not isFirstTime And IsFirst Then
'                            mQtyNew = val(mQty22) - mQtyTotal
'                        End If
'                    End If
                    
                        
                mQtyTotal = mQtyTotal + mQtyNew
                If mQtyTotal > val(mQty22) Then
                    
                    mQtyTotal = mQtyTotal - mQtyNew
                    mQtyNew = 0
                End If
            End If
         '   If (mQtyTotal > val(mQty22) And Not isFirstTime) And Not IsFirst Then GoTo ExitLoop
                If IsFirst Then
                    RsData.AddNew
                    RsData!ItemNameID = val(mItemNo)
                    RsData!UnitID = val(mUnitNo)
                    RsData!GroupID = val(mGroupID)
                    RsData!LineID = val(mLineID)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!SalesID = val(TxtNoteSerial13)
                    RsData!Qty1 = val(mQty22)
                   ' mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                   ' mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
        
                    'RsData!Qty = mQtyNew
                    If i <> mCount Then
                        RsData!Qty = mQtyNew
                    Else
                        If isFirstTime Then
                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
                        Else
                            RsData!Qty = mQtyNew
                        End If
                    End If
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
                Else
                    RsData.Close
                    
                    s = "SELECT * FROM TblProductLineDistribution Where "
                    s = s & "  ItemNameID = " & val(mItemNo)
                    s = s & " and UnitID = " & val(mUnitNo)
                    s = s & " and LineID = " & val(mLineID)
                    s = s & " and IDDefCIT = " & val(TxtTransSerial)
                    s = s & " and ProductLineID = " & val(RsDataLine!ID)
                    
                    Set RsData = New ADODB.Recordset
                    Cn.CommandTimeout = 10000
                    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    
                    mPart = Round(mTotalSecond / mCount)
                    Total = Total + mPart
                    If Total <= mTotalSecond Then
                    
                        If i <> mCount Then
                            RsData!Qty = RsData!Qty + mPart
                        Else
                            RsData!Qty = Abs(RsData!Qty + (mPart - ((mPart * mCount) - mTotalSecond)))
                        End If
                    Else
                     RsData!Qty = mTotalSecond - (Total - mPart)
                    
                    End If
                    
                    RsData.update
                    
                End If
ExitLoop:

            RsDataLine.MoveNext
        Loop
   
End Sub



Private Sub SaveItemsProduction3(ByVal IsFirst As Boolean, ByVal mQty22 As Double, ByVal mRow As Long)
    Dim TxtTransSerial As Long
    Dim TxtNoteSerial13 As Long
    
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    Dim mItemNo As Long, mUnitNo As Long, mGroupID As Long

    mItemNo = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
    mUnitNo = val(grd.TextMatrix(mRow, grd.ColIndex("UnitID")))
    mGroupID = val(grd.TextMatrix(mRow, grd.ColIndex("GroupID")))
    mLineID = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
    mProductLineID = val(grd.TextMatrix(mRow, grd.ColIndex("ProductLineID")))
    mBaseProductLineID2 = val(grd.TextMatrix(mRow, grd.ColIndex("BaseProductLineID2")))
    
    TxtTransSerial = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
    TxtNoteSerial13 = val(grd.TextMatrix(mRow, grd.ColIndex("SalesID")))
    
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
    
'        S = "Select Count(*) CC  from TblGroupItemProductLine Where GroupID = " & val(mGroupID)
'        Dim isFirstTime As Boolean
''        RsDataLine.Close
'        RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If RsDataLine.EOF Then
'            RsDataLine.Close
'            'isFirstTime = True
'            'S = "SELECT *,Qty = 0 FROM TblProductLine Where IsBasicLine = 1"
'            S = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
'            RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
'
'
'
'        RsData.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If Not RsData.EOF Then
'            mCount = val(RsData!CC & "")
'        End If
'
'
'        If mCount = 0 Then
'
'            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
'            Exit Sub
'        End If
'
'
         
         
            s = " SELECT Count(*) CC FROM TblProductLine"

        s = s & " LEFT OUTER JOIN TblGroupItemProductLine ON TblGroupItemProductLine.GroupID = " & mGroupID
        s = s & " AND TblGroupItemProductLine.ProductLineId = TblProductLine.id"
        s = s & " Where IsNull(IsBasicLine, 0) = 0"
        s = s & " AND TblGroupItemProductLine.GroupID = " & mGroupID

        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        Else
            s = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
            Set RsData = New ADODB.Recordset
            RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            mCount = val(RsData!CC & "")
        End If
        If mCount = 0 Then
            
            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " and IDDefCIT = " & TxtTransSerial
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID,BaseProductLineID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(mUnitNo)
        s = s & " and ItemNameID = " & val(mItemNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " Where IsNull(IsBasicLine,0) = 0"
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        s = s & "           AND T.IDDefCIT = " & TxtTransSerial
        s = s & " Group BY ItemNameID,UnitID,T2.ID,BaseProductLineID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "
        
        Dim isFirstTime As Boolean
            isFirstTime = True

            s = "Select ProductLineId ID,Qty = 0  from TblGroupItemProductLine Where GroupID = " & val(mGroupID)
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If RsDataLine.EOF Then
                s = "SELECT *,Qty = 0 FROM TblProductLine Where  IsNull(IsBasicLine,0) = 0"
                RsDataLine.Close
                RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            End If
            
'
'        S = " SELECT * FROM ("
'        S = S & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID,BaseProductLineID FROM TblProductLineDistribution T"
'        S = S & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"
'
'        S = S & " and UnitId = " & val(mUnitNo)
'        S = S & " and ItemNameID = " & val(mItemNo)
'        S = S & " and LineID = " & val(mLineID)
'        S = S & " Where IsNull(IsBasicLine,0) = 0"
'        S = S & " and IsNull(BaseProductLineID,0) = " & mProductLineID
'        S = S & "           AND T.IDDefCIT = " & TxtTransSerial
'        S = S & " Group BY ItemNameID,UnitID,T2.ID,BaseProductLineID"
'        S = S & " ) T "
'        S = S & " Order BY T.Qty DESC "
'
'        Dim isFirstTime As Boolean
''        RsDataLine.Close
'        RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If RsDataLine.EOF Then
'            RsDataLine.Close
'            isFirstTime = True
'            S = "SELECT *,Qty = 0 FROM TblProductLine Where  IsNull(IsBasicLine,0) = 0"
'            RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
            
        
       
        Do While Not RsDataLine.EOF
            i = i + 1
            If i = 1 Then
                mQtyCompare = val(RsDataLine!Qty & "")
            End If
            If IsFirst Then
                If isFirstTime Then
                    mPart = Round(val(mQty22) / mCount)
                    mQtyNew = mPart
                Else
                    mQtyNew = mQtyCompare - val(RsDataLine!Qty & "")
                End If
                
                    If i = mCount Then
                        If isFirstTime And Not IsFirst Then
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                        End If
                    End If
'                    If i = mCount Then
'                        If Not isFirstTime And IsFirst Then
'                            mQtyNew = val(mQty22) - mQtyTotal
'                        End If
'                    End If
                    
                        
                mQtyTotal = mQtyTotal + mQtyNew
                If mQtyTotal > val(mQty22) Then
                    
                    mQtyTotal = mQtyTotal - mQtyNew
                    mQtyNew = 0
                End If
            End If
         '   If (mQtyTotal > val(mQty22) And Not isFirstTime) And Not IsFirst Then GoTo ExitLoop
                If IsFirst Then
                    RsData.AddNew
                    RsData!ItemNameID = val(mItemNo)
                    RsData!UnitID = val(mUnitNo)
                    RsData!GroupID = val(mGroupID)
                    RsData!LineID = val(mLineID)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!SalesID = val(TxtNoteSerial13)
                    RsData!BaseProductLineID = val(mProductLineID)
                    RsData!Qty1 = val(mQty22)
                   ' mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                   ' mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
                    mProdId = val(RsDataLine!ID & "")
                    'RsData!Qty = mQtyNew
                    If i <> mCount Then
                        RsData!Qty = mQtyNew
                    Else
                        If isFirstTime Then
                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
                            mQtyTotal = mQtyTotal - mQtyNew + RsData!Qty
                        Else
                            RsData!Qty = mQtyNew
                        End If
                    End If
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
                Else
                    RsData.Close
                    
                    s = "SELECT * FROM TblProductLineDistribution Where "
                    s = s & "  ItemNameID = " & val(mItemNo)
                    s = s & " and UnitID = " & val(mUnitNo)
                    s = s & " and LineID = " & val(mLineID)
                    s = s & " and IDDefCIT = " & val(TxtTransSerial)
                    s = s & " and ProductLineID = " & val(RsDataLine!ID)
                    s = s & " and  IsNull(BaseProductLineID,0) = " & mProductLineID
                    mProdId = val(RsDataLine!ID & "")
                    Set RsData = New ADODB.Recordset
                    Cn.CommandTimeout = 10000
                    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    
                    mPart = Round(mTotalSecond / mCount)
                    Total = Total + mPart
                    If Total <= mTotalSecond Then
                    
                        If i <> mCount Then
                            RsData!Qty = RsData!Qty + mPart
                        Else
                            RsData!Qty = Abs(RsData!Qty + (mPart - ((mPart * mCount) - mTotalSecond)))
                        End If
                    Else
                     RsData!Qty = mTotalSecond - (Total - mPart)
                    
                    End If
                    
                    RsData.update
                    
                End If
ExitLoop:

            RsDataLine.MoveNext
        Loop
   
End Sub




Private Sub SaveProductionCompilingBasedTable(ByVal IsFirst As Boolean, ByVal mQty22 As Double, ByVal mRow As Long)
    Dim TxtTransSerial As Long
    Dim TxtNoteSerial13 As Long
    
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    Dim mItemNo As Long, mUnitNo As Long, mGroupID As Long

    mItemNo = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
    mUnitNo = val(grd.TextMatrix(mRow, grd.ColIndex("UnitID")))
    mGroupID = val(grd.TextMatrix(mRow, grd.ColIndex("GroupID")))
    mLineID = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
    mProductLineID = val(grd.TextMatrix(mRow, grd.ColIndex("ProductLineID")))
    mBaseProductLineID2 = val(grd.TextMatrix(mRow, grd.ColIndex("BaseProductLineID2")))
    
    TxtTransSerial = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
    TxtNoteSerial13 = val(grd.TextMatrix(mRow, grd.ColIndex("SalesID")))
    
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
    
'        S = "Select Count(*) CC  from TblGroupItemProductLine Where GroupID = " & val(mGroupID)
'        Dim isFirstTime As Boolean
''        RsDataLine.Close
'        RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If RsDataLine.EOF Then
'            RsDataLine.Close
'            'isFirstTime = True
'            'S = "SELECT *,Qty = 0 FROM TblProductLine Where IsBasicLine = 1"
'            S = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
'            RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
'
'
'
'        RsData.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If Not RsData.EOF Then
'            mCount = val(RsData!CC & "")
'        End If
'
'
'        If mCount = 0 Then
'
'            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
'            Exit Sub
'        End If
'
'
         
         
            s = " SELECT Count(*) CC FROM TblProductLine"

        s = s & " LEFT OUTER JOIN TblGroupItemProductLine ON TblGroupItemProductLine.GroupID = " & mGroupID
        s = s & " AND TblGroupItemProductLine.ProductLineId = TblProductLine.id"
        s = s & " Where IsNull(IsBasicLine, 0) = 0"
        s = s & " AND TblGroupItemProductLine.GroupID = " & mGroupID

        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        Else
            s = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
            Set RsData = New ADODB.Recordset
            RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            mCount = val(RsData!CC & "")
        End If
        mCount = 1
        If mCount = 0 Then
            
            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " and IDDefCIT = " & TxtTransSerial
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID,BaseProductLineID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(mUnitNo)
        s = s & " and ItemNameID = " & val(mItemNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " Where IsNull(IsBasicLine,0) = 0"
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        s = s & "           AND T.IDDefCIT = " & TxtTransSerial
        s = s & " Group BY ItemNameID,UnitID,T2.ID,BaseProductLineID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "
        
        Dim isFirstTime As Boolean
            isFirstTime = True

            s = "Select ProductLineId ID,Qty = 0  from TblGroupItemProductLine Where GroupID = " & val(mGroupID)
            
       
       
            s = " SELECT TOP 1 *"
            s = s & "        FROM   ("
            s = s & "              SELECT        TblProductLine.ID, COUNT(TblProductLineDistribution.IDDefCIT) AS Qty"
            
            
            
            s = s & "                From TblGroupItemProductLine"
            s = s & "                                LEFT OUTER JOIN TblProductLineDistribution"
            s = s & "                                     ON  TblProductLineDistribution.ProductLineID = TblGroupItemProductLine.ProductLineId"
            s = s & "                                LEFT OUTER JOIN TblProductLine"
            s = s & "                                     ON  TblGroupItemProductLine.ProductLineId = TblProductLine.id"
            s = s & "                                     AND TblGroupItemProductLine.GroupID = " & val(mGroupID)
            s = s & "                                LEFT OUTER JOIN TblUsersProductLine"
            s = s & "                                     ON  TblUsersProductLine.ProductLineId = TblProductLine.id"
                             
                             

            
            s = s & "              Where  IsNull(IsBasicLine,0) = 0 "
            's = s & " and    (TblUsersProductLine.UserId = " & user_id & " )"
            s = s & "              AND (TblGroupItemProductLine.GroupID =  " & val(mGroupID) & " )"
            s = s & "              AND (TblProductLineDistribution.IDDefCIT =  " & val(TxtTransSerial) & " )"
            
            s = s & "        GROUP BY TblProductLineDistribution.IDDefCIT,  TblProductLine.id"
            s = s & "               ) T"
            s = s & "        Order By"
            s = s & "               Qty"
            Set RsDataLine = New ADODB.Recordset
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If RsDataLine.EOF Then
                
                s = "SELECT Top 1 TblProductLine.ID,Qty = 0 FROM TblProductLine "
                s = s & "                                LEFT OUTER JOIN TblGroupItemProductLine"
                s = s & "                                     ON  TblGroupItemProductLine.ProductLineId = TblProductLine.id"
                s = s & "                                     AND TblGroupItemProductLine.GroupID = " & val(mGroupID)
                s = s & "                                LEFT OUTER JOIN TblUsersProductLine"
                s = s & "                                     ON  TblUsersProductLine.ProductLineId = TblProductLine.id"

                s = s & " Where IsNull(IsBasicLine,0) = 0"
                s = s & "                                     AND TblGroupItemProductLine.GroupID = " & val(mGroupID)
                s = s & " and TblProductLine.id Not In (Select TblProductLineDistribution.ProductLineID from TblProductLineDistribution Where (TblProductLineDistribution.IDDefCIT <>  " & val(TxtTransSerial) & " ))"
                s = s & "                                     ORDER BY        TblProductLine.ID"
                 RsDataLine.Close
                RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If RsDataLine.EOF Then
                
                                
                    s = "SELECT Top 1 TblProductLine.ID,Qty = 0 FROM TblProductLine "
                    s = s & "                                LEFT OUTER JOIN TblGroupItemProductLine"
                    s = s & "                                     ON  TblGroupItemProductLine.ProductLineId = TblProductLine.id"
                    s = s & "                                     AND TblGroupItemProductLine.GroupID = " & val(mGroupID)
                    s = s & "                                LEFT OUTER JOIN TblUsersProductLine"
                    s = s & "                                     ON  TblUsersProductLine.ProductLineId = TblProductLine.id"
    
                    s = s & " Where IsNull(IsBasicLine,0) = 0"
                    s = s & "                                     AND TblGroupItemProductLine.GroupID = " & val(mGroupID)
                    
                    s = s & "                                     ORDER BY        TblProductLine.ID"
                     RsDataLine.Close
                    RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If RsDataLine.EOF Then
                    s = "SELECT Top 1 TblProductLine.ID,Qty = 0 FROM TblProductLine Where  IsNull(IsBasicLine,0) = 0"
                  '  s = s & " and TblProductLine.id Not In (Select TblProductLineDistribution.ProductLineID from TblProductLineDistribution Where (TblProductLineDistribution.IDDefCIT <>  " & val(TxtTransSerial) & " ))"
                    RsDataLine.Close
                    RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    End If
                End If
            End If
            
'
'        S = " SELECT * FROM ("
'        S = S & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID,BaseProductLineID FROM TblProductLineDistribution T"
'        S = S & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"
'
'        S = S & " and UnitId = " & val(mUnitNo)
'        S = S & " and ItemNameID = " & val(mItemNo)
'        S = S & " and LineID = " & val(mLineID)
'        S = S & " Where IsNull(IsBasicLine,0) = 0"
'        S = S & " and IsNull(BaseProductLineID,0) = " & mProductLineID
'        S = S & "           AND T.IDDefCIT = " & TxtTransSerial
'        S = S & " Group BY ItemNameID,UnitID,T2.ID,BaseProductLineID"
'        S = S & " ) T "
'        S = S & " Order BY T.Qty DESC "
'
'        Dim isFirstTime As Boolean
''        RsDataLine.Close
'        RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If RsDataLine.EOF Then
'            RsDataLine.Close
'            isFirstTime = True
'            S = "SELECT *,Qty = 0 FROM TblProductLine Where  IsNull(IsBasicLine,0) = 0"
'            RsDataLine.Open S, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
            
        
       
        Do While Not RsDataLine.EOF
            i = i + 1
            If i = 1 Then
                mQtyCompare = val(RsDataLine!Qty & "")
            End If
            If IsFirst Then
                If isFirstTime Then
                    mPart = Round(val(mQty22) / mCount)
                    mQtyNew = mPart
                Else
                    mQtyNew = mQtyCompare - val(RsDataLine!Qty & "")
                End If
                
                    If i = mCount Then
                        If isFirstTime And Not IsFirst Then
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                        End If
                    End If
'                    If i = mCount Then
'                        If Not isFirstTime And IsFirst Then
'                            mQtyNew = val(mQty22) - mQtyTotal
'                        End If
'                    End If
                    
                        
                mQtyTotal = mQtyTotal + mQtyNew
                If mQtyTotal > val(mQty22) Then
                    
                    mQtyTotal = mQtyTotal - mQtyNew
                    mQtyNew = 0
                End If
            End If
         '   If (mQtyTotal > val(mQty22) And Not isFirstTime) And Not IsFirst Then GoTo ExitLoop
                If IsFirst Then
                    RsData.AddNew
                    RsData!ItemNameID = val(mItemNo)
                    RsData!UnitID = val(mUnitNo)
                    RsData!GroupID = val(mGroupID)
                    RsData!LineID = val(mLineID)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!SalesID = val(TxtNoteSerial13)
                    RsData!BaseProductLineID = val(mProductLineID)
                    RsData!Qty1 = val(mQty22)
                   ' mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                   ' mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
                    mProdId = val(RsDataLine!ID & "")
                    'RsData!Qty = mQtyNew
                    If i <> mCount Then
                        RsData!Qty = mQtyNew
                    Else
                        If isFirstTime Then
                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
                            mQtyTotal = mQtyTotal - mQtyNew + RsData!Qty
                        Else
                            RsData!Qty = mQtyNew
                        End If
                    End If
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
                Else
                    RsData.Close
                    
                    s = "SELECT * FROM TblProductLineDistribution Where "
                    s = s & "  ItemNameID = " & val(mItemNo)
                    s = s & " and UnitID = " & val(mUnitNo)
                    s = s & " and LineID = " & val(mLineID)
                    s = s & " and IDDefCIT = " & val(TxtTransSerial)
                    s = s & " and ProductLineID = " & val(RsDataLine!ID)
                    s = s & " and  IsNull(BaseProductLineID,0) = " & mProductLineID
                    mProdId = val(RsDataLine!ID & "")
                    Set RsData = New ADODB.Recordset
                    Cn.CommandTimeout = 10000
                    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    
                    mPart = Round(mTotalSecond / mCount)
                    Total = Total + mPart
                    If Total <= mTotalSecond Then
                    
                        If i <> mCount Then
                            RsData!Qty = RsData!Qty + mPart
                        Else
                            RsData!Qty = Abs(RsData!Qty + (mPart - ((mPart * mCount) - mTotalSecond)))
                        End If
                    Else
                     RsData!Qty = mTotalSecond - (Total - mPart)
                    
                    End If
                    
                    RsData.update
                    
                End If
ExitLoop:

            RsDataLine.MoveNext
        Loop
        'mProdId = val(RsDataLine!ID & "")
   
End Sub





Private Sub SaveItemsProduction2(ByVal IsFirst As Boolean, ByVal mQty22 As Double, ByVal mRow As Long)
    Dim TxtTransSerial As Long
    Dim TxtNoteSerial13 As Long
          '  DB_CreateField "TblProductLineDistribution", "BaseProductLineID", adInteger, adColNullable, , , "  ", False, True

    Dim rsDistributionDet As New ADODB.Recordset
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    Dim mItemNo As Long, mUnitNo As Long, mGroupID As Long
    Dim mMasterID As Long
    mItemNo = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
    mUnitNo = val(grd.TextMatrix(mRow, grd.ColIndex("UnitID")))
    mGroupID = val(grd.TextMatrix(mRow, grd.ColIndex("GroupID")))
    mLineID = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
    mMasterID = val(grd.TextMatrix(mRow, grd.ColIndex("ID")))
    mProductLineID = val(grd.TextMatrix(mRow, grd.ColIndex("ProductLineID")))
    TxtTransSerial = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
    TxtNoteSerial13 = val(grd.TextMatrix(mRow, grd.ColIndex("SalesID")))
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
        s = "SELECT Count(*) CC FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        End If
        If mCount = 0 Then
            
            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        s = s & " and IDDefCIT = " & TxtTransSerial
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(mUnitNo)
        s = s & " and ItemNameID = " & val(mItemNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " Where IsNull(IsBasicLine,0) = 0"
        s = s & " and IsNull(BaseProductLineID,0) = " & mProductLineID
        s = s & " Group BY ItemNameID,UnitID,T2.ID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "

        Dim isFirstTime As Boolean
'        RsDataLine.Close
        RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If RsDataLine.EOF Then
            RsDataLine.Close
            isFirstTime = True
            s = "SELECT *,Qty = 0 FROM TblProductLine Where IsNull(IsBasicLine,0) = 0"
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
        Do While Not RsDataLine.EOF
            i = i + 1
            If i = 1 Then
                mQtyCompare = val(RsDataLine!Qty & "")
            End If
            If IsFirst Then
                If isFirstTime Then
                    mPart = Round(val(mQty22) / mCount)
                    mQtyNew = mPart
                   ' mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                Else
                    mQtyNew = mQtyCompare - val(RsDataLine!Qty & "")
                End If
                
                    If i = mCount Then
                        If isFirstTime And Not IsFirst Then
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                        End If
                    End If
'                    If i = mCount Then
'                        If Not isFirstTime And IsFirst Then
'                            mQtyNew = val(mQty22) - mQtyTotal
'                        End If
'                    End If
                    
                        
                mQtyTotal = mQtyTotal + mQtyNew
                If mQtyTotal > val(mQty22) Then
                    
                    mQtyTotal = mQtyTotal - mQtyNew
                    mQtyNew = 0
                End If
            End If
         '   If (mQtyTotal > val(mQty22) And Not isFirstTime) And Not IsFirst Then GoTo ExitLoop
                If IsFirst Then
                    RsData.AddNew
                    RsData!ItemNameID = val(mItemNo)
                    RsData!UnitID = val(mUnitNo)
                    RsData!GroupID = val(mGroupID)
                    RsData!LineID = val(mLineID)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!SalesID = val(TxtNoteSerial13)
                    RsData!BaseProductLineID = val(mProductLineID)
                    RsData!Qty1 = val(mQty22)
                   ' mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                   ' mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
        
                    'RsData!Qty = mQtyNew
                    If i <> mCount Then
                        RsData!Qty = mQtyNew
                    Else
                        If isFirstTime Then
                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
                            'mQtyTotal = RsData!Qty
                        Else
                            RsData!Qty = mQtyNew
                            'mQtyTotal = mQtyTotal + RsData!Qty
                        End If
                    End If
                    
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
                    
'                    If val(RsData!Qty & "") <> 0 Then
'                        mId2 = CStr(new_id("TblProductLineDistributionDet", "ID", "", True))
'                        s = "Select * from TblProductLineDistributionDet Where MasterID = " & val(mMasterID)
'                        Set rsDistributionDet = New ADODB.Recordset
'                        rsDistributionDet.Open s, Cn, adOpenKeyset, adLockOptimistic
'                        rsDistributionDet.AddNew
'                        rsDistributionDet!IDDefCIT = RsData!IDDefCIT
'                        rsDistributionDet!ProductLineID = RsData!ProductLineID
'                        rsDistributionDet!BaseProductLineID = mProductLineID
'                        rsDistributionDet!MasterID = val(mMasterID)
'                        rsDistributionDet!Qty = val(RsData!Qty & "")
'                        rsDistributionDet!ID = mId2
'                        rsDistributionDet.update
'                    End If
                Else
                    RsData.Close
                    
                    s = "SELECT * FROM TblProductLineDistribution Where "
                    s = s & "  ItemNameID = " & val(mItemNo)
                    s = s & " and UnitID = " & val(mUnitNo)
                    s = s & " and LineID = " & val(mLineID)
                    s = s & " and IDDefCIT = " & val(TxtTransSerial)
                    s = s & " and ProductLineID = " & val(RsDataLine!ID)
                    s = s & " and  IsNull(BaseProductLineID,0) = " & mProductLineID
                    Set RsData = New ADODB.Recordset
                    Cn.CommandTimeout = 10000
                    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    
                    mPart = Round(mTotalSecond / mCount)
                    Total = Total + mPart
                    If Total <= mTotalSecond Then
                    
                        If i <> mCount Then
                            RsData!Qty = RsData!Qty + mPart
                        Else
                            RsData!Qty = Abs(RsData!Qty + (mPart - ((mPart * mCount) - mTotalSecond)))
                        End If
                    Else
                     RsData!Qty = mTotalSecond - (Total - mPart)
                    
                    End If
'                     s = "Update  TblProductLineDistributionDet Set Qty = " & val(RsData!Qty & "") & "  Where MasterID = " & val(RsData!ID & "")
'                     Cn.Execute s
                    
                    RsData.update
                    If val(RsData!Qty & "") <> 0 Then
                        
                        s = "Select * from TblProductLineDistributionDet Where MasterID = " & val(mMasterID) & " and TblProductLineDistributionDet.ProductLineId  =" & val(RsData!ProductLineID & "") & " and IDDefCIT =" & val(RsData!IDDefCIT & "")
                        Set rsDistributionDet = New ADODB.Recordset
                        rsDistributionDet.Open s, Cn, adOpenKeyset, adLockOptimistic
                        If rsDistributionDet.EOF Then
                            mId2 = CStr(new_id("TblProductLineDistributionDet", "ID", "", True))
                            rsDistributionDet.AddNew
                            rsDistributionDet!IDDefCIT = RsData!IDDefCIT
                            rsDistributionDet!ProductLineID = RsData!ProductLineID
                            rsDistributionDet!BaseProductLineID = mProductLineID
                            rsDistributionDet!MasterID = val(mMasterID)
                            rsDistributionDet!Qty = val(RsData!Qty & "")
                            rsDistributionDet!ID = mId2
                            rsDistributionDet.update
                        End If
                    End If
                    
                    If val(RsData!Qty & "") = 0 Then
                        s = "Update   TblProductLineDistributionDet Set Qty = " & val(RsData!Qty & "") & "    Where TblProductLineDistributionDet.ProductLineId  =" & val(RsData!ProductLineID & "") & " and  MasterID = " & val(mMasterID)
                        Cn.Execute s
                    End If
                    
                End If
ExitLoop:

            RsDataLine.MoveNext
        Loop
   
End Sub


Private Sub PrintStiker(ByVal mRow As Long)
Dim StrSQL As String
Dim StrWhere As String


   Dim mQty As Double
   Dim mQty2 As Double
    Dim MySQL As String
    
    Dim rsDummy As New ADODB.Recordset
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim mPrintType As Long
    Dim mLineID As Long
Dim mSerId As Long
Dim ItemNameID As Long
mSerId = val(grd.TextMatrix(mRow, grd.ColIndex("IDDefCIT")))
mFormPrint = val(grd.TextMatrix(mRow, grd.ColIndex("FormPrint")))
mLineID = val(grd.TextMatrix(mRow, grd.ColIndex("LineID")))
mQty = val(grd.TextMatrix(mRow, grd.ColIndex("Qty")))
mQty2 = val(grd.TextMatrix(mRow, grd.ColIndex("Qty1")))
ItemNameID = val(grd.TextMatrix(mRow, grd.ColIndex("ItemNameID")))
Dim mLineType As Integer
mLineType = val(grd.TextMatrix(mRow, grd.ColIndex("LineType")))
txtMySQL = ""
'If mLineType = 1 Then
    
'    LoadPrintData mSerId, mLineID, mFormPrint, mQty, mQty2, txtMySQL
'Else
    Cn.Execute "Delete TmpItemsQty"
   
    rsDummy.Open "Select * from TmpItemsQty ", Cn, adOpenKeyset, adLockOptimistic
    For i = 1 To mQty
        rsDummy.AddNew
        rsDummy!LineID = mLineID
        rsDummy!ID = i
        rsDummy!Qty = i
        rsDummy.update
        
        'LoadPrintData mSerId, mLineID, mFormPrint, i, mQty2, txtMySQL
'        If i <> mQty Then
'            txtMySQL = txtMySQL & vbNewLine & " Union all "
'        End If
    Next
     LoadPrintData mSerId, mLineID, mFormPrint, mQty, mQty2, txtMySQL, ItemNameID
'End If






'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
 
  StrSQL = txtMySQL
  
  print_report2 StrSQL, 0, mFormPrint
End Sub
Private Sub LoadPrintData(ByVal mSerId As Long, ByVal mLineID As Long, ByVal mFormPrint As Long, ByVal mQty As Double, ByVal mQty2 As Double, ByRef txtMySQL As TextBox, Optional ItemNameID As Long)


'txtMySQL = " SELECT   dbo.TblDefComItem.id as IDDD, TblDefComItem.PaymentType, Grou.GroupName,TblDefComItem.PaymentType ,TblDefComItem.id Transaction_ID, " & mQty & "  Qty1,  tdcid.ItemID, tdcid.ItemCode ItemCode2, dbo.TblItems.ItemName, "

txtMySQL = txtMySQL & " SELECT  dbo.TblDefComItem.id ,TblDefComItem.NoteSerial13 NoteSerial11 ,Grou.GroupName,       TblDefComItem.PaymentType,tdcid.Vat2,tdcid.TotalWithVat,       TblDefComItem.id    Transaction_ID1,  TmpItemsQty.Qty   Qty1," & mQty & "   Qty2,"
txtMySQL = txtMySQL & vbNewLine & "        dbo.TblItems.ItemCode, TblDefComItem.MaxName NoteSerial12,      dbo.TblItems.ItemName,       dbo.TblItems.ItemNamee,dbo.TblDefComItem.RecordDate,"
txtMySQL = txtMySQL & vbNewLine & "        dbo.TblDefComItem.CusID,dbo.TblCustemers.CusName MaxName,dbo.TblCustemers.CusNamee,dbo.TblDefComItem.BranchID,dbo.TblBranchesData.branch_name,tdcid.Remark ItemCode,"
txtMySQL = txtMySQL & vbNewLine & "       dbo.TblBranchesData.branch_nameE,TblDefComItem.ItemNameID,tdcid.widtj,"
txtMySQL = txtMySQL & vbNewLine & "      tdcid.hight,tdcid.Price           ,tdcid.TotalAdd,"
txtMySQL = txtMySQL & vbNewLine & "       tdcid.TotalDisc,tdcid.Net,"
txtMySQL = txtMySQL & vbNewLine & "       tu.UnitName UnitName2,tdcid.LineID MaxNo,tdcid.LineID,TblItems.barCodeNO"
       
txtMySQL = txtMySQL & vbNewLine & " From dbo.TblItems"

     
txtMySQL = txtMySQL & vbNewLine & "       RIGHT OUTER JOIN dbo.TblBranchesData"
     
txtMySQL = txtMySQL & vbNewLine & "       RIGHT OUTER JOIN dbo.TblDefComItem"
     
txtMySQL = txtMySQL & vbNewLine & "            ON  dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID"
txtMySQL = txtMySQL & vbNewLine & "                        LEFT OUTER JOIN TblDefComItemData AS tdcid"
txtMySQL = txtMySQL & vbNewLine & "                        ON  tdcid.IDDefCIT = TblDefComItem.ID"
txtMySQL = txtMySQL & vbNewLine & "       LEFT OUTER JOIN dbo.TblCustemers"
txtMySQL = txtMySQL & vbNewLine & "            ON  dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID"
     
txtMySQL = txtMySQL & vbNewLine & "            ON  dbo.TblItems.ItemID = tdcid.ItemID"

txtMySQL = txtMySQL & vbNewLine & "       LEFT OUTER JOIN TblUnites  AS tu"
txtMySQL = txtMySQL & vbNewLine & "            ON  tu.UnitID = tdcid.UnitID"
txtMySQL = txtMySQL & vbNewLine & "       LEFT OUTER JOIN Groups     AS Grou"
txtMySQL = txtMySQL & vbNewLine & "            ON  Grou.GroupID = tdcid.GroupID "

txtMySQL = txtMySQL & vbNewLine & "       LEFT OUTER JOIN TmpItemsQty     "
txtMySQL = txtMySQL & vbNewLine & "            ON  tdcid.LineID = TmpItemsQty.LineID "
txtMySQL = txtMySQL & vbNewLine & "  Where (dbo.TblDefComItem.id = " & val(mSerId) & ") and tdcid.LineID = " & mLineID
txtMySQL = txtMySQL & vbNewLine & "     and TblItems.ItemID= " & ItemNameID


'StrSQL = txtMySQL & StrWhere



End Sub
Function print_report2(Optional NoteSerial As String, Optional Ind As Integer = 0, Optional ByVal mFormPrint As Long = 0)
   
   
     
    'Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
     Set RsData = New ADODB.Recordset

    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If mFormPrint = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemStiker1.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemStiker1.rpt"
            End If
        ElseIf mFormPrint = 2 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemStiker2.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemStiker2.rpt"
            End If
        Else
           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemStiker.rpt"
        End If
  
  
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    Cn.CommandTimeout = 10000
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    
     Dim i As Long
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@Title}"
         '   xReport.FormulaFields.Item(i).Text = "” Ìﬂ—"
        End Select
    Next i
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    If chkPrintDirect.value = vbChecked Then
        xReport.PrintOut
    Else
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    End If
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub grd_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Select Case grd.ColKey(Col)
  
   Case "Start", "End"
    grd.TextMatrix(Row, Col) = Not grd.ValueMatrix(Row, Col)
'
'  Case "PrintStiker"
'        Cancel = False
'        grd.TextMatrix(Row, grd.ColIndex("PrintStiker")) = ""
'         PrintStiker Row
'
'          mIDD = val(grd.TextMatrix(Row, grd.ColIndex("ID")))
'          s = "Update TblProductLineDistribution Set UserId =  " & user_id
'          s = s & " ,PrintTime = '" & Format(Time, "hh:mm:ss") & "'"
'          s = s & " , PrintDate = " & SQLDate(Date, True) & ""
'          s = s & " Where Id = " & mIDD
'          Cn.Execute s
'          FillGrid3
'        Cancel = True
    Case "Convert"
        If Trim(grd.TextMatrix(Row, grd.ColIndex("PrintDate"))) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Can not convert before printing"
            Else
                MsgBox "·« Ì„ﬂ‰ «· ÕÊÌ· ﬁ»· «·ÿ»«⁄…"
            End If
        End If
    Case "IsPrinted"
        Cancel = True
    End Select
End Sub

Private Sub GridInstallments_Click()
On Error Resume Next

If GridInstallments.Col = GridInstallments.ColIndex("View") Then
Dim i As Integer
i = GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("NoteID"))
If i > 0 Then
FrmExpenses4.Retrive (i)
End If

End If
End Sub

Private Sub Label20_Click()
FramRows.Visible = False
End Sub

Private Sub Option1_Click(Index As Integer)
      With Me.FG
       .Rows = 1
End With
mOld = False
ProgressBar1.Visible = True
: ProgressBar1.value = 10
FillGrid2
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0

    
End Sub

Private Sub Text1_Change(Index As Integer)
If IsNumeric(Text1(Index).Text) Then
    If Index = 0 Then
        Timer1.interval = 1 * 60 * 100
    ElseIf Index = 1 Then
        Timer2.interval = 1 * 60 * 100
    ElseIf Index = 2 Then
        Timer3.interval = 1 * 60 * 100
    End If
End If
 
 
End Sub

Private Sub Timer1_Timer()
'FillGrid
End Sub

Private Sub Timer3_Timer()
'FillGrid3
End Sub

Private Sub ToDate_Change()
FillGrid
End Sub



Public Sub LoadGrid2(ByVal Sqlstmt As String, _
                          ByRef tGrd As Control, _
                          Optional ResetRows As Boolean = True, _
                          Optional InsertRow As Boolean = False, _
                          Optional mReCreateColumns As Boolean = False)
    Dim tRs As New ADODB.Recordset
  
    
    tRs.Open Sqlstmt, Cn, adOpenStatic, adLockReadOnly, adCmdText
  
    ' ******************************************
    If ResetRows Then tGrd.Rows = tGrd.FixedRows
    ' ******************************************
    If mReCreateColumns Then
        tGrd.Cols = 1
        tGrd.Cols = tRs.Fields.count + 1
        For i = 1 To tGrd.Cols - 1
            tGrd.ColKey(i) = tRs.Fields.Item(i - 1).Name
            tGrd.TextMatrix(0, i) = tRs.Fields.Item(i - 1).Name
        Next
    End If
    ' ******************************************
    ' ******************************************
    tGrd.Redraw = flexRDNone
    ' ******************************************
    i = tGrd.Rows
    sCur = 0
    Do While Not tRs.EOF
        tGrd.AddItem i - tGrd.FixedRows + 1
        For j = 0 To tRs.Fields.count - 1
            
            
            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                If tRs.Fields.Item(j).Type = adCurrency And mWithMyFormat Then
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = (val(tRs.Fields.Item(j).value & ""))
                Else
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = Trim(tRs.Fields.Item(j).value & "")
                End If
            End If
        Next
        i = i + 1
        sCur = sCur + 1

        tRs.MoveNext
    Loop
    tRs.Close
    Set tRs = Nothing

    If InsertRow Then tGrd.AddItem tGrd.Rows - tGrd.FixedRows + 1
    tGrd.Redraw = flexRDDirect
End Sub


Public Sub saveGrid(ByVal Sqlstmt As String, ByRef tGrd As vsFlexGrid, ByVal ChekPoint As String, ByVal Index As String, ParamArray FieldValue())
    On Error GoTo Err
    Dim tRs As New ADODB.Recordset

    
    tRs.Open Sqlstmt, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    ' *******************************************
    Dim II As Integer
    II = 0
    For i = tGrd.FixedRows To tGrd.Rows - 1
        If ChekPoint <> "" Then
            If Trim(tGrd.TextMatrix(i, tGrd.ColIndex(ChekPoint))) = "" Then GoTo NextStep
        End If
        '**********************
        tRs.AddNew
        II = II + 1
        If Index <> "" Then tRs(Index) = II
        For k = 0 To UBound(FieldValue) Step 2
            tRs.Fields.Item(FieldValue(k)).value = FieldValue(k + 1)
            'Debug.Print FieldValue(k) & " " & tRs.Fields.Item(FieldValue(k)).Value
        Next
        '*************************
        'Debug.Print "fields count " & tRs.Fields.count
        For j = 0 To tRs.Fields.count - 1

            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                If tRs.Fields.Item(j).Type = adInteger Or tRs.Fields.Item(j).Type = adCurrency Or tRs.Fields.Item(j).Type = adBoolean Or tRs.Fields.Item(j).Type = adSmallInt Or tRs.Fields.Item(j).Type = adBigInt Or tRs.Fields.Item(j).Type = adTinyInt Or tRs.Fields.Item(j).Type = adUnsignedTinyInt Or tRs.Fields.Item(j).Type = adNumeric Then
                    If tRs.Fields.Item(j).Type = adBoolean Then
                        tRs.Fields.Item(j).value = (UCase(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "TRUE") Or (UCase(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "-1") Or (val(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = -1)
                    Else
'                        If tGrd.ColComboList(tGrd.ColIndex(tRS.Fields.Item(j).Name)) <> "" Then
'                            tRS.Fields.Item(j).Value = tGrd.ValueMatrix(i, tGrd.ColIndex(tRS.Fields.Item(j).Name))
'                        Else
                            'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                            tRs.Fields.Item(j).value = val(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)))
                            'End If
'                        End If
                    End If
                Else
                    If tRs.Fields.Item(j).Type = adDBTimeStamp Or tRs.Fields.Item(j).Type = adDBTime Or tRs.Fields.Item(j).Type = adDBDate Then
                        If Not IsDate(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) Then
                            tRs.Fields.Item(j).value = Null
                        Else
                            tRs.Fields.Item(j).value = tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))
                        End If
                    Else
                        'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                        tRs.Fields.Item(j).value = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))
                        'End If
                    End If
                End If
            End If
            'Debug.Print tRs.Fields.Item(j).Name & " = " & tRs.Fields.Item(j).Value
        Next

NextStep:
    Next
    tRs.Close
    Exit Sub
Err:
    If Err.Number = -2147217887 Then        ' one item is empty
        Resume Next
    End If
    '    Resume Next
End Sub

