VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOpeningBalance 
   Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   HelpContextID   =   90
   Icon            =   "FrmOpeningBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOpeningBalance.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   12240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7845
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   12240
      _cx             =   21590
      _cy             =   13838
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
      AutoSizeChildren=   8
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
      GridRows        =   6
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmOpeningBalance.frx":0714
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   5
         Left            =   15
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   6840
         Width           =   12180
         _cx             =   21484
         _cy             =   767
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   8700
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   345
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3255
            TabIndex        =   17
            Top             =   75
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   9210
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   0
            Width           =   1980
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   315
            Index           =   63
            Left            =   7380
            TabIndex        =   67
            Top             =   120
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6540
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   9165
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   0
            Width           =   1950
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·—’Ìœ"
            Height          =   255
            Index           =   3
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   120
            Width           =   1500
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   6
            Left            =   5115
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   105
            Width           =   795
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   135
            Width           =   765
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   2055
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   105
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   360
            Index           =   5
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   480
            Index           =   4
            Left            =   2715
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   0
            Width           =   525
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4785
         Index           =   3
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   12210
         _cx             =   21537
         _cy             =   8440
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
         AutoSizeChildren=   8
         BorderWidth     =   2
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
         GridRows        =   3
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmOpeningBalance.frx":078C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.Toolbar TBr 
            Height          =   630
            Left            =   30
            TabIndex        =   31
            Top             =   4395
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
            Begin VSFlex8Ctl.VSFlexGrid grdExcel 
               Height          =   2040
               Index           =   1
               Left            =   2160
               TabIndex        =   75
               Top             =   270
               Width           =   12255
               _cx             =   21616
               _cy             =   3598
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
               GridLines       =   3
               GridLinesFixed  =   2
               GridLineWidth   =   5
               Rows            =   2
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmOpeningBalance.frx":07EE
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
               ExplorerBar     =   3
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
               Begin VSFlex8Ctl.VSFlexGrid grdExcel 
                  Height          =   2040
                  Index           =   0
                  Left            =   630
                  TabIndex        =   76
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   12255
                  _cx             =   21616
                  _cy             =   3598
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
                  GridLines       =   3
                  GridLinesFixed  =   2
                  GridLineWidth   =   5
                  Rows            =   2
                  Cols            =   16
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOpeningBalance.frx":0A46
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
                  ExplorerBar     =   3
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
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   4
            Left            =   30
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   30
            Width           =   12150
            _cx             =   21431
            _cy             =   1217
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
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   615
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   300
               Width           =   1500
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   3690
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   300
               Width           =   1875
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2205
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   300
               Width           =   1410
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   5610
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   300
               Width           =   1755
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   7365
               TabIndex        =   6
               Top             =   270
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   9735
               TabIndex        =   5
               Top             =   300
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   420
               Left            =   30
               TabIndex        =   11
               Top             =   210
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   741
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
               ButtonImage     =   "FrmOpeningBalance.frx":0CDE
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ·›…"
               Height          =   270
               Index           =   26
               Left            =   780
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   270
               Index           =   27
               Left            =   2415
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   30
               Width           =   1200
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ì—Ì«·"
               Height          =   390
               Index           =   28
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   30
               Width           =   1785
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   270
               Index           =   29
               Left            =   5715
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   30
               Width           =   1650
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   270
               Index           =   30
               Left            =   7575
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   30
               Width           =   2160
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   270
               Index           =   31
               Left            =   9945
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   30
               Width           =   2190
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   3645
            Left            =   30
            TabIndex        =   65
            Top             =   735
            Width           =   12150
            _cx             =   21431
            _cy             =   6429
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
            Cols            =   31
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmOpeningBalance.frx":1078
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
         Begin VB.Label LblItemsCount 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   495
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   4395
            Width           =   11220
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1380
         Index           =   1
         Left            =   15
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   645
         Width           =   12180
         _cx             =   21484
         _cy             =   2434
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.CommandButton cmdLoadFile_TrialBalanceExcel 
            Caption         =   "«” Ì—«œ"
            Height          =   285
            Left            =   2910
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1110
            Width           =   1335
         End
         Begin VB.CheckBox chkMeth2 
            Alignment       =   1  'Right Justify
            Caption         =   "«” Ì—«œ 2"
            Height          =   195
            Left            =   750
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1170
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton CMDSelectFile 
            Caption         =   "Õœœ «·„·›"
            Height          =   315
            Left            =   4290
            TabIndex        =   72
            Top             =   1110
            Width           =   1230
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   210
            TabIndex        =   71
            Top             =   825
            Visible         =   0   'False
            Width           =   4410
         End
         Begin VB.CommandButton CmdImport 
            Caption         =   "«” Ì—«œ «·„·›"
            Height          =   375
            Left            =   1860
            TabIndex        =   70
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9630
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   930
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2655
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1560
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   885
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   0
            Width           =   5685
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   240
               TabIndex        =   40
               Top             =   180
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   240
               TabIndex        =   41
               Top             =   510
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label LblAccountInterval 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   6780
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   510
               Width           =   255
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   180
               Width           =   15
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   285
               Index           =   9
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1710
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   285
               Index           =   8
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1740
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   7
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   510
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   13695
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   90
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   8550
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   60
            Width           =   1965
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1020
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   630
            Visible         =   0   'False
            Width           =   750
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   5910
            TabIndex        =   4
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   100073475
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   5910
            TabIndex        =   0
            Top             =   480
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   5910
            TabIndex        =   1
            Top             =   840
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   -6120
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
            Height          =   1830
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   2445
            _cx             =   4313
            _cy             =   3228
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
            Rows            =   50
            Cols            =   40
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„·›"
            Height          =   210
            Index           =   15
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Tag             =   "53"
            Top             =   840
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   375
            Index           =   2
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   840
            Width           =   1545
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   10515
            TabIndex        =   64
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   375
            Index           =   0
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   105
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   375
            Index           =   1
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   75
            Width           =   1545
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   15
         Width           =   12210
         _cx             =   21537
         _cy             =   1085
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
         Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
            Left            =   1575
            TabIndex        =   49
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "FrmOpeningBalance.frx":1550
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
            Left            =   900
            TabIndex        =   50
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
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
            ButtonImage     =   "FrmOpeningBalance.frx":18EA
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
            Left            =   2430
            TabIndex        =   51
            Top             =   120
            Width           =   555
            _ExtentX        =   979
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
            ButtonImage     =   "FrmOpeningBalance.frx":1C84
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
            TabIndex        =   52
            Top             =   120
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmOpeningBalance.frx":201E
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7305
         Width           =   12240
         _cx             =   21590
         _cy             =   952
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
         Align           =   2
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
            Height          =   360
            Index           =   0
            Left            =   10980
            TabIndex        =   54
            Top             =   105
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
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
            Height          =   360
            Index           =   1
            Left            =   9630
            TabIndex        =   55
            Top             =   105
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   635
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
            Height          =   360
            Index           =   2
            Left            =   8235
            TabIndex        =   56
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
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
            Height          =   360
            Index           =   3
            Left            =   6975
            TabIndex        =   57
            Top             =   105
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   635
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
            Height          =   360
            Index           =   4
            Left            =   5295
            TabIndex        =   58
            Top             =   105
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
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
            Height          =   360
            Index           =   5
            Left            =   4140
            TabIndex        =   59
            Top             =   105
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   635
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
            Height          =   360
            Index           =   6
            Left            =   30
            TabIndex        =   60
            Top             =   105
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
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
            Height          =   360
            Index           =   7
            Left            =   2655
            TabIndex        =   61
            Top             =   105
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
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
            Height          =   360
            Left            =   1485
            TabIndex        =   62
            Top             =   105
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   635
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
   End
End
Attribute VB_Name = "FrmOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim BalanceReport As ClsOpeningBalanceReport
Dim cSearchDcbo As clsDCboSearch
Dim NewGrid As New ClsGrid
Dim FirstPeriodDateInthisYear  As Date
Public Sub RetriveSerials(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
    ' For i = 1 To Fg.Rows - 2
    '        If Fg.TextMatrix(i, Fg.ColIndex("Code")) = ItemID Then
    '         Me.Fg.RemoveItem (i)
    '         i = 1
    '        End If
    'NewGrid.Grid_AfterEdit Num, Fg.ColIndex("Code")
    ' Next i
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    For intX = 0 To UBound(astrSplitItems)
   
        FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
        
        ' FG.TextMatrix(Num, FG.ColIndex("Name")) = itemname
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        NewGrid.Grid_AfterEdit Num, FG.ColIndex("Code")
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
  
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        

 
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.rows = FG.rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.text = "F"
    'TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Private Sub cmdLoadFile_Click()
'ExportToExcel Me, Grd, "TT", , "grdExcel"
tmpGrd.rows = 1

Dim i As Long
Dim s As String
Dim mIndex As Long


Dim rsDummy As New ADODB.Recordset
Dim rsDummy2 As New ADODB.Recordset
Dim mCode As String
Dim mGroupID As Long
Dim mUnitId As Long
Dim mUnitPurPrice As Double
Dim mUnitSalesPrice As Double
Dim mRatePur As Double
Dim mRateSale As Double
Dim mNewCode  As String
Dim mMaxId As Long
Dim mName As String
Dim mbarCode As String
Dim mUnitWholeSalePrice As Double
Dim mUnitName As String
Dim rsDummyUnit As New ADODB.Recordset
Dim mQty As Double
Dim StrSQL As String
mIndex = 1
Dim rs2 As New ADODB.Recordset
Dim LngItemID As Long
    grdExcel(1).rows = 1
    FromExcel grdExcel(1), tmpGrd, Me, , , txtFile.text, "TblEmployee"
'Dim StrSQL As String

    

For i = 1 To grdExcel(mIndex).rows - 1
    mCode = Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("Fullcode")))
    If mCode = "986203V000" Then
        mCode = mCode
    End If
    mbarCode = Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("barCodeNO")))
    mQty = val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("ShowQty")))
    mUnitName = Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("UnitName")))
    mUnitWholeSalePrice = val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("UnitWholeSalePrice")))
    Set rsDummyUnit = New ADODB.Recordset
    s = "Select UnitName,UnitId from TblUnites Where UnitName Like '%" & Trim(mUnitName) & "%'"
    rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummyUnit.EOF Then
        mUnitId = val(rsDummyUnit!UnitID & "")
    End If
    
    
    mUnitPurPrice = val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("UnitPurPrice")))
    mName = Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("ItemName")))
   
    
    mUnitSalesPrice = val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("UnitSalesPrice")))
   
   
    StrSQL = "  SELECT     dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, TblItems.ItemName,dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
     StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
     StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
      StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
     StrSQL = StrSQL & "  WHERE     (dbo.TblItems.Fullcode = N'" & mCode & "' oR dbo.TblItems.Fullcode = N'0" & mCode & "' Or dbo.TblItems.barCodeNO = N'" & mCode & "'  Or dbo.TblItems.ItemCode = N'" & mCode & "') "
     
    ' StrSQL = StrSQL & "  Or (TblItems.ItemName Like '" & Trim(mName) & "'  Or TblItems.ItemNamee Like '" & Trim(mName) & "')) AND (dbo.TblItemsUnits.UnitID = " & mUnitId & ")"
     StrSQL = StrSQL & "   "
     StrSQL = StrSQL & "  GROUP BY dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID,TblItems.ItemName, dbo.TblItemsUnits.ItemID"
     Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then

         LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
         mUnitId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
    Else
    
         StrSQL = "  SELECT     dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, TblItems.ItemName,dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
         StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
         StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
          StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
         StrSQL = StrSQL & "  WHERE     (dbo.TblItems.Fullcode = N'" & mCode & "' oR dbo.TblItems.Fullcode = N'0" & mCode & "' Or dbo.TblItems.barCodeNO = N'" & mCode & "'  Or dbo.TblItems.ItemCode = N'" & mCode & "' "
         
         StrSQL = StrSQL & "  Or (TblItems.ItemName Like '" & Trim(mName) & "'  Or TblItems.ItemNamee Like '" & Trim(mName) & "')) AND (dbo.TblItemsUnits.UnitID = " & mUnitId & ")"
         StrSQL = StrSQL & "   "
         StrSQL = StrSQL & "  GROUP BY dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID,TblItems.ItemName, dbo.TblItemsUnits.ItemID"
         Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs2.RecordCount > 0 Then
            LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            mUnitId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
        Else
            Debug.Print mCode & " " & mName & " " & mUnitPurPrice & " " & mQty; vbNewLine
            GoTo NextRow
        End If
        
    End If
'         ParrtNoCode = "" 'IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
'        ItemDetailedCode = "" 'IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID = 0 Then GoTo NextRow
    If mCode = "" Then GoTo NextRow
       
         
    With Me.FG

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If
        .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
        .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
        .TextMatrix(.rows - 1, FG.ColIndex("Count")) = mQty
       ' .TextMatrix(.Rows - 1, FG.ColIndex("DiscountType")) = 1
        
        .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
        .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
         FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = 1 ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = 1 ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = 1 ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
         FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = mUnitId ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
         'FG.TextMatrix(.Rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
         'FG.TextMatrix(.Rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
        .TextMatrix(.rows - 1, FG.ColIndex("Price")) = mUnitPurPrice
       ' .TextMatrix(.Rows - 1, FG.ColIndex("ShowPrice")) = mUnitPurPrice
        .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))
If SystemOptions.UserInterface = ArabicInterface Then
             FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
Else
    FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
End If

     End With


NextRow:
Next
Dim mFlag As String
mFlag = TxtModFlg.text
If FG.rows > 1 Then
    mQty = val(FG.TextMatrix(FG.rows - 1, FG.ColIndex("Count")))
    
   ' TxtModFlg.Text = "E"
    
    FG.TextMatrix(FG.rows - 1, FG.ColIndex("Count")) = mQty
    'FG_AfterEdit FG.Rows - 1, 1
    mUnitPurPrice = val(FG.TextMatrix(FG.rows - 1, FG.ColIndex("Price")))
    NewGrid.Grid_AfterEdit FG.rows - 1, 1
    
    FG.TextMatrix(FG.rows - 1, FG.ColIndex("Price")) = mUnitPurPrice
    
End If
'TxtModFlg.Text = mFlag

's = " UPDATE Groups SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and GroupID <> 1"
' Cn.Execute s
'
' s = " UPDATE groups SET Code = fullCode WHERE ISNULL(code,'') = ''"
' Cn.Execute s
 
 
End Sub



Private Sub cmdLoadFile_TrialBalanceExcel_Click()

    On Error GoTo ErrTrap

    Dim xlApp As Object, xlWB As Object, xlWs As Object
    Dim lastrow As Long, r As Long

    Dim mCode As String, mName As String
    Dim Qty As Double, valAmount As Double, unitCost As Double
Dim UnitName As String
    Dim ItemID As Long, UnitID As Long

    If Trim$(txtFile.text) = "" Then Exit Sub

    '==============================
    ' «› Õ „·› «·«ﬂ”Ì·
    '==============================
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False
    xlApp.Visible = False

    Set xlWB = xlApp.Workbooks.Open(txtFile.text)
    Set xlWs = xlWB.Worksheets(1)

    '==============================
    ' ¬Œ— ’› »‰«¡ ⁄·Ï ⁄„Êœ U («·ﬂÊœ)
    '==============================
    Dim lastCell As Object
Set lastCell = xlWs.cells.Find(What:="*", LookIn:=-4163, LookAt:=1, _
                               SearchOrder:=1, SearchDirection:=2) ' xlByRows=1 , xlPrevious=2 , xlFormulas=-4163
If Not lastCell Is Nothing Then
    lastrow = lastCell.Row
Else
    lastrow = 0
End If


    If Me.FG.rows < 2 Then Me.FG.rows = 2

    ' €«·»« «·œ« «  »œ√ „‰ 6 (“Ì «·’Ê—…) - ·Ê ⁄‰œﬂ €Ì—Â« ⁄œ· Â‰«
    For r = 6 To lastrow

        ' U = «·ﬂÊœ
        mCode = Trim$(CStr(xlWs.cells(r, 21).text))
mCode = Replace(mCode, CHR(160), "") ' Ì‘Ì· «·„”«›… €Ì— «·„—∆Ì… NBSP

       If mCode = "" Then
    Debug.Print "Skipped row (empty code): "; r; "  Name=" & mName
    GoTo NextRow
End If


        ' P = «·«”„ («Œ Ì«—Ì)
        mName = Trim$(CStr(xlWs.cells(r, 16).value2))

        ' B = «·ﬂ„Ì… «·Œ «„Ì…
        Qty = val(CStr(xlWs.cells(r, 13).value2))

        ' A = «·ﬁÌ„… «·Œ «„Ì…
        valAmount = val(CStr(xlWs.cells(r, 12).value2))

        '  Ã«Â· «·√’‰«› «··Ì —’ÌœÂ« ’›—
  '  If Qty = 0 And valAmount = 0 Then GoTo NextRow

If Qty = 0 Then
    unitCost = valAmount / 1

Else
    unitCost = valAmount / Qty
End If

        
        '==============================
        ' „«» «·ﬂÊœ -> ItemID
        '==============================
        ItemID = GetItemIdByAnyCode(mCode, mName)
        If ItemID = 0 Then GoTo NextRow

        ' ÊÕœ… ‘—«¡ «› —«÷Ì…
     '   UnitID = GetDefaultItemUnit(ItemID)
        
   
        GetDefaultItemUnit ItemID, UnitID, UnitName
        If UnitID = 0 Then UnitID = 1

        '==============================
        ' ‰“¯· ⁄·Ï FG
        '==============================
        With Me.FG
            If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                .rows = .rows + 1
            End If

            .TextMatrix(.rows - 1, .ColIndex("Code")) = ItemID
            .TextMatrix(.rows - 1, .ColIndex("Name")) = ItemID   ' ‰›” √”·Ê»ﬂ «·ﬁœÌ„
            .TextMatrix(.rows - 1, .ColIndex("Count")) = Qty
            .TextMatrix(.rows - 1, .ColIndex("Price")) = unitCost
            .TextMatrix(.rows - 1, .ColIndex("Valu")) = Qty * unitCost

            .cell(flexcpData, .rows - 1, .ColIndex("UnitID")) = UnitID
            
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = UnitName
            Else
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = UnitName
            End If
   
        End With

NextRow:
    Next r

CleanUp:
    On Error Resume Next
    xlWB.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrTrap:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume CleanUp

End Sub
'=========================================================
'  —Ã⁄ True ·Ê ·ﬁ  Qty/Value ’«·Õ… (Qty>0 √Ê Value<>0)
' qtyCol = ⁄„Êœ «·ﬂ„Ì…° valCol = ⁄„Êœ «·ﬁÌ„…
'=========================================================
Private Function ReadPair(ByVal WS As Object, ByVal r As Long, _
                          ByVal qtyCol As Long, ByVal valCol As Long, _
                          ByRef outQty As Double, ByRef outVal As Double) As Boolean

    Dim q As Double, v As Double

    q = val(CStr(WS.cells(r, qtyCol).value2))
    v = val(CStr(WS.cells(r, valCol).value2))

    ' «⁄ »—Â« ’«·Õ… ·Ê ›ÌÂ Õ—ﬂ…/—’Ìœ
    If (q <> 0) Or (v <> 0) Then
        outQty = q
        outVal = v
        ReadPair = True
    Else
        ReadPair = False
    End If

End Function


'=========================================================
' „«» «·ﬂÊœ ·√Ì ’‰›
'=========================================================
Private Function GetItemIdByAnyCode(ByVal mCode As String, ByVal mName As String) As Long

    On Error GoTo EH
    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT TOP 1 ItemID " & _
          "FROM dbo.TblItems " & _
          "WHERE Fullcode = N'" & Replace(mCode, "'", "''") & "' " & _
          "   OR Fullcode = N'0" & Replace(mCode, "'", "''") & "' " & _
          "   OR barCodeNO = N'" & Replace(mCode, "'", "''") & "' " & _
          "   OR ItemCode = N'" & Replace(mCode, "'", "''") & "' "

    ' fallback »«·«”„ («Œ Ì«—Ì) ·Ê  Õ»Â:
    ' sql = sql & " OR ItemName = N'" & Replace(mName, "'", "''") & "' OR ItemNamee = N'" & Replace(mName, "'", "''") & "'"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        GetItemIdByAnyCode = CLng(rs.Fields(0).value)
    Else
        GetItemIdByAnyCode = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    GetItemIdByAnyCode = 0
End Function


'=========================================================
' ÌÃÌ» ÊÕœ… «·‘—«¡ «·«› —«÷Ì… ··’‰›
' (⁄œ¯· ORDER BY Õ”» ÕﬁÊ·ﬂ ≈‰ ÊÃœ  IsDefaultPur/IsMain/Ö)
'=========================================================
Private Function GetDefaultPurchaseUnitId(ByVal ItemID As Long) As Long

    On Error GoTo EH
    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT TOP 1 UnitID " & _
          "FROM dbo.TblItemsUnits " & _
          "WHERE ItemID = " & ItemID & " " & _
          "ORDER BY ISNULL(IsDefaultPur,0) DESC, ISNULL(IsMain,0) DESC, UnitID"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        GetDefaultPurchaseUnitId = CLng(rs.Fields(0).value)
    Else
        GetDefaultPurchaseUnitId = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    GetDefaultPurchaseUnitId = 0
End Function




Public Sub FromExcel(ByRef mGrid As Object, _
                     ByRef mtmpGrd As Object, _
                     Frm As Form, _
                     Optional MainFormName As String = "", _
                     Optional ProgressBar As Object = Nothing, Optional ByVal XlsFileName As String = "", Optional ByVal MainTableName As String = "")


    ' If Not i Then Exit Sub
       Dim cProgress As ClsProgress
       Dim Hide As Integer
       Dim i As Long
       Dim j As Long
       Dim jj As Long
       Dim H As Long
    '    Dim mtmpGrd As VSFlexGrid
    If XlsFileName = "" Then
    MsgBox "Õœœ «·„·› «Ê·«", vbCritical
    Exit Sub
        'XlsFileName = GetGridFileName(mGrid, MainFormName)
    End If
    If FileExists(XlsFileName) Then

        mtmpGrd.FixedCols = 0
        mtmpGrd.FixedRows = 0

        mtmpGrd.loadgrid XlsFileName, flexFileExcel

        mtmpGrd.backcolor = &HFFFFFF
        mtmpGrd.BackColorAlternate = &HE9E9E9
        mtmpGrd.BackColorBkg = &H8000000C
        mtmpGrd.BackColorFixed = &H8000000F
        mtmpGrd.BackColorFrozen = &HC0FFFF
        mtmpGrd.BackColorSel = &H8000000D
        mtmpGrd.ForeColor = &H80000008
        mtmpGrd.ForeColorFixed = &HFF0000
        mtmpGrd.ForeColorSel = &H8000000E
        mtmpGrd.GridColor = &H8000000F
        mtmpGrd.GridColorFixed = &H80000010
        mtmpGrd.FixedCols = 1
        mtmpGrd.FixedRows = 1
        '·«‰ Loaded ÌŒ ›Ì
        mtmpGrd.Cols = mGrid.Cols + 1
        mtmpGrd.ColKey(mtmpGrd.Cols - 1) = "Loaded"
        mtmpGrd.ColHidden(mtmpGrd.Cols - 1) = True
        mtmpGrd.AutoSize 0, mtmpGrd.Cols - 1
    End If
    mGrid.rows = 1
    mGrid.rows = mtmpGrd.rows

    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Min = 1
        ProgressBar.Max = IIf(mGrid.rows > 2, mGrid.rows - 1, 2)    ' mGrid.Rows - 1
        ProgressBar.Visible = True
        '********************************
    End If
        Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    



       
        For i = 1 To mtmpGrd.rows - 1
        '********************************
        If Not ProgressBar Is Nothing Then
            ProgressBar.value = i
            DoEvents
            ProgressBar.Refresh
        End If
        cProgress.StartProgress
       DoEvents
        '********************************
        jj = 0
        For j = 1 To mGrid.Cols - 1
            If j = 18 Then
                j = 18
            End If
            If Not mGrid.ColHidden(j) Then
                jj = jj + 1
                       If mGrid.ColKey(j) = "MainGroumName" Then
                    j = j
                End If
                If i > mGrid.rows - 1 Then Exit Sub
                Debug.Print i & " " & mGrid.TextMatrix(i, j)
                If InStr(1, mGrid.ColComboList(j), "#") Then
                    Hide = 0
                    For H = j - 1 To 1 Step -1
                        Hide = Hide + IIf(mGrid.ColHidden(H), 1, 0)
                    Next
                    mGrid.TextMatrix(i, j) = mtmpGrd.TextMatrix(i, j - Hide)
                    'Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                Else
                    mGrid.TextMatrix(i, j) = Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                End If
                If Trim(mGrid.ColEditMask(j)) = "Date" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid
                End If
                'pValue = Split(G.ColComboList(j), ";")
            Else
                j = j
                If j = 34 Then
                j = j
                End If
                If Trim(mGrid.ColEditMask(j)) <> "" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid, MainTableName
                End If
                If Trim(mGrid.ColComboList(j)) <> "" Then
                    GetIDCombo Trim(mGrid.ColComboList(j)), i, j, mGrid
                End If
            End If
            If Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 1)), "'", "")) = "" Then
                mGrid.rows = i + 1:  Exit Sub
            End If
        Next
        ' DisplayOrderTotals
NextRow:
    Next
    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Visible = False
    End If
           DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    MsgBox " „ «·«œ—«Ã"
    '********************************
End Sub


Private Sub CmdImport_Click()
 
 chkMeth2.value = vbChecked
If chkMeth2.value = vbChecked Then
    cmdLoadFile_Click
    Exit Sub
End If


If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim itemcode As String
Dim account_serial As String
Dim des As String
Dim DebitValue As String
Dim CreditValue As String
  Dim Price As String
 Dim unitcode As String
 Dim UnitName As String
 
Dim ItemName As String
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
Dim RsUnit As New ADODB.Recordset
Dim rsITem As New ADODB.Recordset

    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
      FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 1000
    Dim mItemName As String
    With ExcelSheet
    i = 2
    Do Until .cells(i, 1) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
   itemcode = .cells(i, 1)
    mItemName = .cells(i, 2)
    Set rsITem = New ADODB.Recordset
    s = "Select * from tblItems where ItemName Like '" & Trim(mItemName) & "' "
    rsITem.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsITem.EOF Then
        MsgBox "«·’‰› ”ÿ— " & i - 1 & " " & mItemName & "€Ì— „”Ã· ›Ï „·› «·’‰›"
    End If
    Set RsUnit = New ADODB.Recordset
    s = "SELECT * FROM TblItemsUnits AS tiu WHERE tiu.DefaultUnit = 1 AND tiu.ItemID = " & val(rsITem!ItemID & "")
    RsUnit.Open s, Cn, adOpenStatic, adLockReadOnly
    Qty = .cells(i, 3)
    Price = .cells(i, 5)
    itemcode = val(rsITem!ItemID & "")
    
    
    unitcode = RsUnit!UnitID
    UnitName = .cells(i, 4)
    
         des = .cells(i, 7)
    '    DebitValue = .Cells(i, 5)
    '     CreditValue = .Cells(i, 6)
         
     
 With FG

     
'  .TextMatrix(i, .ColIndex("des")) = (des)
  
  
   FG.TextMatrix(i, FG.ColIndex("Code")) = itemcode
                .TextMatrix(i, FG.ColIndex("Name")) = itemcode
                .TextMatrix(i, FG.ColIndex("Count")) = val(Qty)
              .TextMatrix(i, FG.ColIndex("Price")) = val(Price)
 .TextMatrix(i, FG.ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Count")))


        FG.TextMatrix(1, FG.ColIndex("ColorID")) = 1
            FG.TextMatrix(1, FG.ColIndex("ItemSize")) = 1
            FG.TextMatrix(1, FG.ColIndex("ClassID")) = 1
        
            FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = unitcode
            '        Fg.TextMatrix(RowNum, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(1, FG.ColIndex("UnitID")) = UnitName
            Else
                FG.TextMatrix(1, FG.ColIndex("UnitID")) = UnitName
            End If
   


     
  ' Fg_Journal_AfterEdit i, .ColIndex("account_serial")
   

'    Fg_Journal_AfterEdit i, .ColIndex("BranchId")
       
   
 End With
        i = i + 1
    Loop

    End With

    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing


End Sub

Private Sub CMDSelectFile_Click()
CD1.ShowOpen
txtFile.text = CD1.FileName

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 8
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
    On Error Resume Next
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If val(DCboStoreName.BoundText) <> 0 Then
        Dcbranch.BoundText = GetInventoryBranch(DCboStoreName.BoundText)
    End If


 
    
    
    WriteDev
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
    DCboStoreName_Change
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName
    End If

End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches Dcbranch
    End If

End Sub

Private Sub Ele_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 2 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , 0, TxtTransSerial, 0
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , 0, TxtTransSerial, 101
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), 0, TxtTransSerial, 101

    End If

End Sub

Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

Private Sub Form_Load()
    Dim RsItems As New ADODB.Recordset
    Dim StrSQL As String
    Dim BGround As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim Msg As String

    ScreenNameArabic = "—’Ìœ √›  «ÕÌ"
    ScreenNameEnglish = "Opening Balance"
    
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 0
    
    On Error GoTo ErrTrap
    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
 
 
    If SystemOptions.usertype <> UserAdmin Then
            If checkmanyBranches = False Then
                   Me.Dcbranch.Enabled = True
             End If
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Resize_Form Me, TransactionSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FG.WallPaper = BGround.Picture
    AddTip
    SetDtpickerDate XPDtbBill
    NewGrid.GridTrans = OpeningBalance
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.Grid = FG
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = cmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.txtPrice = txtPrice
    Set NewGrid.GrdTBar = Me.TBr
    ' Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.txtTotal = Me.XPTxtSum
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.LblTotalQty = Me.LblTotalQty
        Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBranches Dcbranch
    
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName

    StrSQL = "Select * From Transactions where Transaction_Type=3"
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
         If SystemOptions.usertype <> UserAdminAll Then
 '       StrSQL = StrSQL & " AND   BranchId=" & Current_branch
    End If
    
    
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    XPBtnMove_Click 2
    TxtModFlg.text = "R"
DoEvents
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    Msg = Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRight, App.Title
End Sub

Private Sub WriteDev()
    On Error GoTo errortrap
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

        If Account_Code_dynamic = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
            Me.DcboDebitSide.BoundText = ""
            Exit Sub
        End If
        
        Me.DcboDebitSide.BoundText = Account_Code_dynamic 'Õ”«» «·„Œ“Ê‰
        'Me.DcboDebitSide.BoundText = "a1a2a5"'
    
        Account_Code_dynamic = get_account_code_branch(19, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        Else

            If Account_Code_dynamic = "NO account" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ   Õ”«» Ê”Ìÿ «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
              Else
              MsgBox " Specify Opening Balnace For Store", vbCritical
              End If
                Exit Sub
         
            End If
        End If
        
        Me.DcboCreditSide.BoundText = Account_Code_dynamic 'Ã”«» Ê”Ìÿ «›  «ÕÌ
        'Me.DcboCreditSide.BoundText = "a2a1a1" '
 
    End If

errortrap:
End Sub

Private Sub LblTotal_Change()
    lblTotalView.Caption = Format(val(lblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
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

Private Sub Cmd_Click(Index As Integer)
  '  On Error GoTo ErrTrap

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.XPDtbBill.value = FirstPeriodDateInthisYear
 
    Dim intDef As Integer

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))

            WriteDev
            GridDefaultValue FG.rows - 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            FG.SetFocus
            FG.rows = 2
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
'            Me.dcBranch.BoundText = branch_id



            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                Dcbranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                Dcbranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.Dcbranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
        




      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.Dcbranch.Enabled = True
                                   Else
                                    Me.Dcbranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If
                        
            Me.Dcbranch.BoundText = Current_branch


        Case 1
                                   If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            'If AvailableDeal = True Then
            TxtModFlg.text = "E"
            DCboStoreName_Change
            Me.DCboUserName.BoundText = user_id

            'End If
        Case 2
                                   If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                              
            SaveData

        Case 3
            Call Undo

        Case 4
                                   If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If

            '   If BolPrint = False Then
            '       Exit Sub
            '   End If

            printing

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBalanceSearch.show vbModal

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ"
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
        
            'XPBtnAdd.Enabled = False
            'XPBtnRemove.Enabled = False
            Me.XPDtbBill.Enabled = False
            Me.DCboStoreName.locked = True
            FG.Editable = flexEDNone

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

            'XPFillData.Enabled = False
            ELe(4).Enabled = False

        Case "N"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '  Me.XPBtnMove(0).Enabled = False
            '  Me.XPBtnMove(1).Enabled = False
            '  Me.XPBtnMove(2).Enabled = False
            '  Me.XPBtnMove(3).Enabled = False
        
            Me.XPDtbBill.Enabled = True
            Me.DCboStoreName.locked = False
            XPDtbBill.value = Date
            FG.Editable = flexEDKbdMouse
        
            ELe(4).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ(  ⁄œÌ· )"
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
                
            Me.XPDtbBill.Enabled = True
            FG.Enabled = True
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
            ELe(4).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If AvailableDeal = True Then
                If Not rs.RecordCount < 1 Then
                    rs.delete
                    
                      CuurentLogdata ("D")
                      
                    StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                    Cn.Execute StrSqlDel, , adExecuteNoRecords
       
                    '  Update_opening_balance_screen_accounts
                    '       MsgBox " „ «·Õ–›"
       
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ Â–Â «·»Ì«‰« " & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
       Else
             Msg = "Confirm Undo"
       End If
       
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            
            End If

        Case "E"
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
   Else
             Msg = "Confirm Undo"
       End If
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                Me.TxtModFlg.text = "R"
                Retrive

            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 0
    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set cSearchDcbo = Nothing
    Set rs = Nothing
    Set TTP = Nothing
    Set BalanceReport = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Exit Sub
ErrTrap:
End Sub
Function SaveItemsData()
   Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & (Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            
           If FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
             
                
                          RsgGrantee.AddNew
                         astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
                         RsgGrantee("ItemDetailedCode").value = (astrSplitItems1(0))
                         RsgGrantee("ParrtNoCode").value = (astrSplitItems1(1))
                         RsgGrantee("count").value = val(astrSplitItems1(2))
                         RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = 1
                    RsgGrantee.update
                                    Next intX
                Else
                If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                RsgGrantee.AddNew
              RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
            RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
            RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
           RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
          RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
          RsgGrantee("EffectN").value = 1
           RsgGrantee.update
                  
         End If
         
                   
                   End If
                   

 
                
  
                    
            End If

       

    Next RowNum


End Function

Function SaveGoldData()
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete TblGoldDetail   where Transaction_ID= " & val(Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  TblGoldDetail Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
                        
                  
                        RsgGrantee.AddNew
 
'                        RsgGrantee("itemid").value = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
'    astrSplitItems = Split(AllIDS, strFilterText)
    

 
 
    
    astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
            RsgGrantee("TTypeId").value = val(astrSplitItems1(0))
            RsgGrantee("typeid").value = val(astrSplitItems1(1))
            RsgGrantee("uniteid").value = val(astrSplitItems1(2))
            RsgGrantee("type").value = val(astrSplitItems1(3))
            RsgGrantee("price").value = val(astrSplitItems1(4))
            RsgGrantee("weight").value = val(astrSplitItems1(5))
            RsgGrantee("Count").value = val(astrSplitItems1(6))
            RsgGrantee("InstallPrice").value = val(astrSplitItems1(7))
         RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
         
      
  
               
                         
                         RsgGrantee.update
                  
                       
                Next intX
                    
            End If

        End If

    Next RowNum

End Function




Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„  «·”‰œ " & TxtTransSerial.text & CHR(13) & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·„Œ“‰ " & DCboStoreName.text & CHR(13) & " «Ã„«·Ì «·”‰œ  " & lblTotalView.Caption
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr No " & TxtTransSerial.text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Store  " & DCboStoreName.text & CHR(13)
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 101, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , 0, TxtTransSerial
    Else
        AddToLogFile CInt(user_id), 101, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , 0, TxtTransSerial
    End If
    
End Function

Private Sub SaveData()
    Dim RSTransDetails As New ADODB.Recordset
    Dim RsSerial As New ADODB.Recordset
    Dim RsNotes As New ADODB.Recordset
    Dim RsCheckSerial As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim BegineTrans As Boolean
    Dim IntRes As Integer
    Dim LngDev As Long
    Dim LngNoteID As Long
    Dim StrTemp As String

     On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
CuurentLogdata
    If Trim(Dcbranch.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
            
        Else
           Msg = "Specify Departement"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Dcbranch.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Me.TxtModFlg.text <> "R" Then
        If DCboStoreName.BoundText = "" Then
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
                 Else
                 Msg = "Please Select Store"
                 End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If NewGrid.IsReaptedSerials = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÊÃœ  ﬂ—«— ›Ï √—ﬁ«„ «·”Ì—Ì«· «·„œŒ·… "
            Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·√—ﬁ«„ «·„œŒ·…"
         Else
         Msg = "You have Repeated Serial "
         End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        '-----------------------------------------------
        'Check the Items Grid
          If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        '----------------------------------------------
        Cn.BeginTrans
        BegineTrans = True
    
        If TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))
                      
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.text)
            
        Else
        
         StrSQL = "delete        DOUBLE_ENTREY_VOUCHERS1   Where  opening_balance_voucher_id =" & val(txtopening_balance_voucher_id.text)
     
      Cn.Execute StrSQL
        
        End If
    
      txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
      
    '    RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
 
        
        rs("BranchId").value = IIf(Me.Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
        rs("Transaction_Serial").value = Me.TxtTransSerial.text
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Type").value = 3
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
        rs.update

        If Me.TxtModFlg.text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
          
            StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
    
        End If

 StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
      
        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = XPTxtBillID.text
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

   RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
   If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                    RsCheckSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsCheckSerial.EOF Or RsCheckSerial.BOF) Then
                        If RsCheckSerial("HaveSerial").value = True Then
                            RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                        End If
                    End If

                    RsCheckSerial.Close
                End If
RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("Price").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("BranchId").value = IIf(Me.Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
                ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
                ' RSTransDetails("ItemSize").value = _
                  IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
'new gold

  RSTransDetails("showPrice1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("showPrice1")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("showPrice1"))))
  RSTransDetails("ShowQty1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ShowQty1")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ShowQty1"))))
  RSTransDetails("Salaries1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Salaries1")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Salaries1"))))
 RSTransDetails("Salaries2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Salaries2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Salaries2"))))
                                             
     RSTransDetails("showPrice2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("showPrice2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("showPrice2"))))
  RSTransDetails("ShowQty2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ShowQty2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ShowQty2"))))
                                          
              '  RSTransDetails("ScurrencyID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("ScurrencyID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("ScurrencyID"))))
                RSTransDetails("Scurrenyrate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Scurrenyrate")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Scurrenyrate"))))
                                          RSTransDetails("ScurrencyID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ScurrencyID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ScurrencyID"))))
                                          
'gplllllllllllllllld
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
           
                
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                     RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                     RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                End If

                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                              
                RSTransDetails("showprice").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))

                RSTransDetails("OpeningBurcahseQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty"))))
                RSTransDetails("OpeningBurcahseValue").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue"))))
                RSTransDetails("OpeningSalesQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty"))))
                RSTransDetails("OpeningSalesValue").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue"))))
                RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                '*******************************************************************
                RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
        
        RSTransDetails("GoldDetails").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")))
        RSTransDetails("ItemsDetailsNewidea").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")))
        
        RSTransDetails("Wages").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Wages")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("Wages")))
        
                    Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
       RSTransDetails("OldQty").value = NewQty
       RSTransDetails("OldCost").value = NewCost
       
      RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
      If (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value) <> 0 Then
       RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       Else
      RSTransDetails("NewCost").value = 0
       End If
            RSTransDetails.update
            
                '*******************************************************************
                RSTransDetails.update
            End If

        Next RowNum

        Set RsNotes = New ADODB.Recordset
    
        'RsNotes.Open "NOTES1", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    
 StrSQL = "SELECT     * from dbo.NOTES1 Where (1 = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
        Dim NoteSerial As String
        Dim NoteSerial1 As String

        NoteSerial = year(XPDtbBill.value) & 1
        NoteSerial1 = NoteSerial
   
        '    Dim NoteSerial As String
        '    Dim noteserial1 As String
        '
        '    NoteSerial = ""
        '
        '                  If NoteSerial = "" Then
        '                       If Notes_coding(Val(my_branch), XPDtbBill.value) = "error" Then
        '                       MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        '                       Else
        '
        '                       If Notes_coding(Val(my_branch), XPDtbBill.value) = "" Then
        '                       MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
        '                       Else
        '                       NoteSerial = Notes_coding(Val(my_branch), XPDtbBill.value)
        '                       End If
        '                       End If
        '                End If
        '
        '                If noteserial1 = "" Then
        '                   If Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000) = "error" Then
        '                   MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ «›  «ÕÌ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        '                   Else
        '
        '                   If Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000) = "" Then
        '                   MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        '                   Else
        '                   noteserial1 = Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000)
        '                   End If
        '                   End If
        '                End If
    
        '    RsNotes.AddNew
        '        'LngNoteID = new_id("NOTES", "NoteID", "")
        '        'RsNotes("NoteID").value = LngNoteID
        '        RsNotes("NoteID").value = 1
        '        RsNotes("NoteDate").value = XPDtbBill.value
        '        RsNotes("NoteType").value = 101
        '        RsNotes("NoteSerial").value = NoteSerial ' new_id("NOTES", "NoteSerial", "", True, "NOTETYPE=100")
        '          RsNotes("NoteSerial1").value = noteserial1
        '
        '        RsNotes("Note_Value").value = Val(lbltotal.Caption)
        '        RsNotes("Transaction_ID").value = Val(Me.XPTxtBillID.text)
        '    RsNotes.update
        '«·œ«‘‰
        LngNoteID = 1
        ' LngDev = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        LngDev = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "")

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = "—’Ìœ ≈›  «ÕÏ ··„Œ“‰    " & TxtStoreID & "     " & Trim(Me.DCboStoreName.text) & " ··›—⁄ " & Dcbranch.text
        Else
            StrTemp = "   Opening Balance No:  " & TxtStoreID & "     " & Trim(Me.TxtTransSerial.text) & " Branch " & Dcbranch.text
        End If
    
        If val(Me.lblTotal.Caption) >= 0 Then
            If ModAccounts.AddNewDev(LngDev, 1, Me.DcboCreditSide.BoundText, val(Me.lblTotal.Caption), 1, StrTemp, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Me.XPDtbBill.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            Else
                '  update_account_opening_balance Me.DcboCreditSide.BoundText
            End If
        End If
    
        '    If ModAccounts.AddNewDev(LngDev, 2, Me.DcboCreditSide.BoundText, Val(Me.lbltotal.Caption), 1, _
        '        StrTemp, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Me.XPDtbBill.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
        '        GoTo ErrTrap
        '    End If
    
        Dim LngDevNO  As Integer
        Dim StrTempAccountCode As String
        Dim StrTempDes As String
        Dim SngTemp  As Variant

        Dim Account_Code_dynamic As String
        Dim i As Integer

        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "")
        LngDevNO = 0
        '«·ÿ—› «·„œÌ‰
        SngTemp = (Me.lblTotal.Caption)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Then

                Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…

                ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "—’Ìœ «›  «ÕÌ   ··„Œ«“‰ ·⁄«„" & year(XPDtbBill.value)
                Else
                    StrTempDes = "Opening Balance Year" & year(XPDtbBill.value)
                End If

                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDev, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, Me.txtopening_balance_voucher_id, , , , val(Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                Else
                    '      update_account_opening_balance StrTempAccountCode
                End If

            ElseIf detect_inventory_work_type = 2 Then
                '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
                Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                    GoTo ErrTrap
                End If
    
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "—’Ìœ «›  «ÕÌ ··„Œ«“‰   ·⁄«„ " & year(XPDtbBill.value)
                Else
                    StrTempDes = "Opening Balance For Inventory Year" & year(XPDtbBill.value)
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDev, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id), , , , val(Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                Else
                    '     update_account_opening_balance StrTempAccountCode
            
                End If

            ElseIf detect_inventory_work_type = 3 Then
                Dim groupAccount As String
             
                Dim line_value As Single

                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                            groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "—’Ìœ «›  «ÕÌ  ··„Œ«“‰  ·⁄«„  " & year(XPDtbBill.value)
                            Else
                                StrTempDes = "Opening Balance Year" & year(XPDtbBill.value)
                            End If

                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDev, LngDevNO, groupAccount, line_value, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, Me.txtopening_balance_voucher_id, , , , val(Dcbranch.BoundText)) = False Then
                                GoTo ErrTrap
                            Else
                                '    update_account_opening_balance groupAccount
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        Cn.CommitTrans
        BegineTrans = False
        Me.LblDevID.Caption = LngDev
        Me.LblAccountInterval.Caption = SystemOptions.SysCurrentAccountIntervalID
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
      SaveGoldData
      SaveItemsData
        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Successfully Saved " & CHR(13)
                    Msg = Msg + "Do you want to enter another  New operation"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Successfully Updated", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            
        End Select

        TxtModFlg.text = "R"

        If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Or SystemOptions.SysMainStockCostMethod = LastPurPriceType Then
            '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
            'ÂÊ
            'ModernWeightAverage
            '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
            '«·„ÊÃÊœ… ›Ï «·›« Ê—…
            UpdateTransCost val(Me.XPTxtBillID.text)
        End If
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BegineTrans = True Then
        Cn.RollbackTrans
        BegineTrans = False
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  "
    Msg = Msg & CHR(13) & "" & Err.Description
    Msg = Msg & CHR(13) & "" & Err.Number
    Msg = Msg & CHR(13) & "" & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

'Me.TxtModFlg.text = "R"

    Screen.MousePointer = vbArrowHourglass
    Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", val(rs("BranchId").value))
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    Me.TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
    DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    StrSQL = StrSQL + "order by id"

    'StrSql = "select * From Transaction_Details where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
                FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty").value), "", RsDetails("showqty").value)
                .TextMatrix(RowNum, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
                .TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)

                If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                Else
                    FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
                End If

                .TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice").value), "", RsDetails("showPrice").value)
                .TextMatrix(RowNum, FG.ColIndex("Valu")) = val(.TextMatrix(RowNum, .ColIndex("Price"))) * val(.TextMatrix(RowNum, .ColIndex("Count")))
         'gooooooooooooooooooold
   .TextMatrix(RowNum, FG.ColIndex("showPrice1")) = IIf(IsNull(RsDetails("showPrice1").value), "", RsDetails("showPrice1").value)

  .TextMatrix(RowNum, FG.ColIndex("ShowQty1")) = IIf(IsNull(RsDetails("ShowQty1").value), "", RsDetails("ShowQty1").value)

  .TextMatrix(RowNum, FG.ColIndex("Salaries1")) = IIf(IsNull(RsDetails("Salaries1").value), "", RsDetails("Salaries1").value)

  .TextMatrix(RowNum, FG.ColIndex("Salaries2")) = IIf(IsNull(RsDetails("Salaries2").value), "", RsDetails("Salaries2").value)


'gooooooooooooooooooold
.TextMatrix(RowNum, FG.ColIndex("Value1")) = val(.TextMatrix(RowNum, .ColIndex("showPrice1"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty1")))


   .TextMatrix(RowNum, FG.ColIndex("showPrice2")) = IIf(IsNull(RsDetails("showPrice2").value), "", RsDetails("showPrice2").value)

  .TextMatrix(RowNum, FG.ColIndex("ShowQty2")) = IIf(IsNull(RsDetails("ShowQty2").value), "", RsDetails("ShowQty2").value)
.TextMatrix(RowNum, FG.ColIndex("Value2")) = val(.TextMatrix(RowNum, .ColIndex("showPrice2"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty2")))

.TextMatrix(RowNum, FG.ColIndex("TotalCost")) = val(.TextMatrix(RowNum, .ColIndex("Salaries1"))) + val(.TextMatrix(RowNum, .ColIndex("Salaries2"))) + val(.TextMatrix(RowNum, .ColIndex("Value1"))) + val(.TextMatrix(RowNum, .ColIndex("Value2")))

            
                 FG.TextMatrix(RowNum, FG.ColIndex("ScurrencyID")) = IIf(IsNull(RsDetails("ScurrencyID")), 1, (RsDetails("ScurrencyID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("Scurrenyrate")) = IIf(IsNull(RsDetails("Scurrenyrate")), 1, (RsDetails("Scurrenyrate").value))
       
            End With

 


            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            '        Fg.TextMatrix(RowNum, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If
       
            FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty")) = IIf(IsNull(RsDetails("OpeningBurcahseQty").value), "", RsDetails("OpeningBurcahseQty").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue")) = IIf(IsNull(RsDetails("OpeningBurcahseValue").value), "", RsDetails("OpeningBurcahseValue").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty")) = IIf(IsNull(RsDetails("OpeningSalesQty").value), "", RsDetails("OpeningSalesQty").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue")) = IIf(IsNull(RsDetails("OpeningSalesValue").value), "", RsDetails("OpeningSalesValue").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", RsDetails("FoxyNo").value)
            '******************************************
            FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
        'ItemsDetailsNewidea
            FG.TextMatrix(RowNum, FG.ColIndex("GoldDetails")) = IIf(IsNull(RsDetails("GoldDetails")), "", (RsDetails("GoldDetails").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
        FG.TextMatrix(RowNum, FG.ColIndex("Wages")) = IIf(IsNull(RsDetails("Wages")), "", (RsDetails("Wages").value))
        
            '*************************************************
            RsDetails.MoveNext

            If FG.rows > 10 Then
              '  If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Me.XPTxtSum.text = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Valu"), FG.rows - 1, FG.ColIndex("Valu"))
    Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.rows - 1, FG.ColIndex("Count"))
    
     'ss StrSQL = "Select * From NOTES Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
     'ss  Set RsNotes = New ADODB.Recordset
     'ss  RsNotes.Open StrSQL, Cn

     'ss  If Not (RsNotes.BOF Or RsNotes.EOF) Then
     'ss    LngNoteID = RsNotes("NoteID").value
     'ss     StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & LngNoteID & ""
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
        StrSQL = StrSQL + " Order BY DEV_ID_Line_No"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or RsDev.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.LblAccountInterval.Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    
  'ss End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    NewGrid.CountItems
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BalanceReport = New ClsOpeningBalanceReport
            BalanceReport.ShowOpeningBalanceData XPTxtBillID.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BalanceReport = New ClsOpeningBalanceReport
            BalanceReport.ShowOpeningBalanceData XPTxtBillID.text, True 'Short View
        End If
    End If

    'If XPTxtBillID.text <> "" Then
    '    Set BalanceReport = New ClsOpeningBalanceReport
    '    BalanceReport.ShowOpeningBalanceData XPTxtBillID.text
    'End If
    'Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long
    On Error GoTo ErrTrap
AvailableDeal = True
Exit Function

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    '                StrSql = "select * From QryGardComplete where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                    '                StrSql = StrSql + " AND ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    '                StrSql = StrSql + " AND StoreID=" & DCboStoreName.BoundText
                    '                Set RsTemp = New ADODB.Recordset
                    '                RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    '                If RsTemp.EOF Or RsTemp.BOF Then
                    With FrmAlarm
                        .DealingForm = OpeningBalance
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = New ADODB.Recordset
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("totalqty").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = OpeningBalance
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

Private Sub GridDefaultValue(Crow As Long)
    On Error GoTo ErrTrap

    With FG
        .TextMatrix(Crow, .ColIndex("ItemCase")) = 1
        .TextMatrix(Crow, .ColIndex("Count")) = 1
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            
            End If
        End If
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Opening Balance"
    C1Elastic6.Caption = Me.Caption

    lbl(1).Caption = "ID"
    lbl(0).Caption = "Date"
    Label3.Caption = "Branch"

    lbl(2).Caption = "Store "
    lbl(63).Caption = "Total Qty "

    Fra.Caption = "GL"
    lbl(8).Caption = "GL#"
    lbl(9).Caption = "Interval"
    lbl(32).Caption = "Depit"
    lbl(7).Caption = "Credit"

    lbl(3).Caption = " Total:"
 
    lbl(6).Caption = " By:"
    lbl(4).Caption = "Curr. rec"
    lbl(5).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
   
    With Me.FG
        .TextMatrix(0, .ColIndex("NewItem")) = "NewItem"
        .TextMatrix(0, .ColIndex("order_no")) = "order No"
        .TextMatrix(0, .ColIndex("OrderArrivalDate")) = "Order Arrival Date"
    End With
   
    'NewItem

End Sub

Private Sub XPTxtSum_Change()
    Me.lblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub



Private Sub GetFieldID(ByVal mTableColName As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object, Optional ByVal MainTableName As String = "")
    Dim mTableName As String
    Dim mFieldIDName As String
    Dim mFieldName As String
    Dim xx As Variant
    Dim mValue As String
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummy2 As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function


Private Sub GetIDCombo(ByVal mTableColID As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object)
Dim mTxt As String
mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Select Case mTableColID
Case "sexID"
    If mTxt = "Male" Or mTxt = "–ﬂ—" Then
        mTxt = 1
    Else
        mTxt = 2
    End If
Case "MaritalStatusID"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "√⁄“»" Or mTxt = "Single" Then
        mTxt = 0
    ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
        mTxt = 1
    ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
        mTxt = 2
    ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
        mTxt = 3
        
    End If
Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
    mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
Case ""
End Select
mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub



Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then CheckDateIsHij = 3: Exit Function
    
    If Trim(mDate) = "" Then CheckDateIsHij = 3: Exit Function
    
    If year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function



