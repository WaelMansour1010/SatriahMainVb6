VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmRequest 
   BackColor       =   &H00E2E9E9&
   Caption         =   "«·√’‰«› «·„ÿ·Ê»…"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12825
   HelpContextID   =   410
   Icon            =   "FrmRequuest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12825
      _cx             =   22622
      _cy             =   15690
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8400
         Left            =   0
         TabIndex        =   1
         Top             =   -240
         Width           =   12855
         _cx             =   22675
         _cy             =   14817
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·«’‰«› «· Ì »·€  Õœ «·ÿ·»|«’‰«› „ Êﬁ⁄ «‰  ’· ·Õœ «·ÿ·»|«·„—›ﬁ« "
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
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic EleMain 
            Height          =   7980
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   12765
            _cx             =   22516
            _cy             =   14076
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   1815
               Left            =   6720
               TabIndex        =   41
               Top             =   600
               Width           =   6045
               Begin VB.ListBox ListBranchAll 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":058A
                  Left            =   3240
                  List            =   "FrmRequuest.frx":0591
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   330
                  Width           =   2655
               End
               Begin VB.ListBox ListBranchSelected 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":05A4
                  Left            =   120
                  List            =   "FrmRequuest.frx":05AB
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   330
                  Width           =   2655
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·›—⁄"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   2310
                  TabIndex        =   48
                  Top             =   120
                  Width           =   1440
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   1095
                  Width           =   495
               End
               Begin VB.Label Label10 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   240
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1425
                  Width           =   495
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   750
                  Width           =   495
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   420
                  Width           =   495
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Height          =   1815
               Left            =   840
               TabIndex        =   33
               Top             =   600
               Width           =   5925
               Begin VB.ListBox ListStoreSelect 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":05C3
                  Left            =   120
                  List            =   "FrmRequuest.frx":05CA
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   330
                  Width           =   2610
               End
               Begin VB.ListBox ListAllStore 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":05DF
                  Left            =   3165
                  List            =   "FrmRequuest.frx":05E6
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   360
                  Width           =   2625
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·„Œ“‰"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   2130
                  TabIndex        =   40
                  Top             =   120
                  Width           =   1470
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   360
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   1305
                  Width           =   495
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   975
                  Width           =   495
               End
            End
            Begin C1SizerLibCtl.C1Elastic EleHeader 
               Height          =   405
               Left            =   30
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   180
               Width           =   12855
               _cx             =   22675
               _cy             =   714
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   555
                  Index           =   1
                  Left            =   570
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2595
                  _cx             =   4577
                  _cy             =   979
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "ÿ—Ìﬁ… ⁄—÷ «·√’‰«›"
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
                  Style           =   1
                  TagSplit        =   2
                  PicturePos      =   4
                  CaptionStyle    =   3
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
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ÃœÊ·Ï"
                     Height          =   195
                     Index           =   0
                     Left            =   1365
                     TabIndex        =   6
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   1170
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ‘Ã—Ì"
                     Height          =   315
                     Index           =   1
                     Left            =   90
                     TabIndex        =   5
                     Top             =   180
                     Width           =   1200
                  End
               End
               Begin VB.Label LblCaption 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√’‰«› «· Ì »·€  Õœ «·ÿ·»"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   20.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   2580
                  TabIndex        =   7
                  Top             =   30
                  Width           =   6705
               End
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   0
                  Picture         =   "FrmRequuest.frx":05F8
                  Top             =   30
                  Width           =   480
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   5550
               Left            =   0
               TabIndex        =   8
               Top             =   2430
               Width           =   12735
               _cx             =   22463
               _cy             =   9790
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
               BackColorBkg    =   16777215
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmRequuest.frx":12C2
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
            Begin MSDataListLib.DataCombo DcbStore 
               Bindings        =   "FrmRequuest.frx":14C0
               Height          =   315
               Left            =   630
               TabIndex        =   9
               Top             =   450
               Visible         =   0   'False
               Width           =   4470
               _ExtentX        =   7885
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
            Begin MSDataListLib.DataCombo DcbBranch 
               Bindings        =   "FrmRequuest.frx":14D5
               Height          =   315
               Left            =   7125
               TabIndex        =   27
               Top             =   330
               Visible         =   0   'False
               Width           =   4470
               _ExtentX        =   7885
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
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   1815
               Left            =   120
               TabIndex        =   31
               Top             =   600
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   3201
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
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
               ButtonImage     =   "FrmRequuest.frx":14EA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   300
               Index           =   3
               Left            =   10815
               TabIndex        =   28
               Top             =   450
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Œ“‰"
               Height          =   300
               Index           =   17
               Left            =   4680
               TabIndex        =   10
               Top             =   450
               Visible         =   0   'False
               Width           =   1560
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7980
            Left            =   13500
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   12765
            _cx             =   22516
            _cy             =   14076
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
               BackColor       =   &H00E2E9E9&
               Height          =   1815
               Left            =   6720
               TabIndex        =   57
               Top             =   720
               Width           =   6045
               Begin VB.ListBox ListBranchSelected2 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":1884
                  Left            =   120
                  List            =   "FrmRequuest.frx":188B
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   330
                  Width           =   2655
               End
               Begin VB.ListBox ListBranchAll2 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":18A3
                  Left            =   3240
                  List            =   "FrmRequuest.frx":18AA
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   330
                  Width           =   2655
               End
               Begin VB.Label Label17 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   420
                  Width           =   495
               End
               Begin VB.Label Label18 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   750
                  Width           =   495
               End
               Begin VB.Label Label19 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   240
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   1425
                  Width           =   495
               End
               Begin VB.Label Label20 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1095
                  Width           =   495
               End
               Begin VB.Label Label21 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·›—⁄"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   2310
                  TabIndex        =   60
                  Top             =   120
                  Width           =   1440
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   1815
               Left            =   840
               TabIndex        =   49
               Top             =   720
               Width           =   5925
               Begin VB.ListBox ListAllStore2 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":18BD
                  Left            =   3165
                  List            =   "FrmRequuest.frx":18C4
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   360
                  Width           =   2625
               End
               Begin VB.ListBox ListStoreSelect2 
                  Height          =   1425
                  ItemData        =   "FrmRequuest.frx":18D6
                  Left            =   120
                  List            =   "FrmRequuest.frx":18DD
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   360
                  Width           =   2610
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   975
                  Width           =   495
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   360
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   1305
                  Width           =   495
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   630
                  Width           =   495
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Left            =   2715
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·„Œ“‰"
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Left            =   2130
                  TabIndex        =   52
                  Top             =   120
                  Width           =   1470
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   405
               Left            =   -150
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   180
               Width           =   10185
               _cx             =   17965
               _cy             =   714
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   555
                  Index           =   2
                  Left            =   570
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2595
                  _cx             =   4577
                  _cy             =   979
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "ÿ—Ìﬁ… ⁄—÷ «·√’‰«›"
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
                  Style           =   1
                  TagSplit        =   2
                  PicturePos      =   4
                  CaptionStyle    =   3
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
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ‘Ã—Ì"
                     Height          =   315
                     Index           =   3
                     Left            =   90
                     TabIndex        =   22
                     Top             =   180
                     Width           =   1200
                  End
                  Begin VB.OptionButton Opt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ ÃœÊ·Ï"
                     Height          =   195
                     Index           =   2
                     Left            =   1365
                     TabIndex        =   21
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   1170
                  End
               End
               Begin VB.Image Image2 
                  Height          =   480
                  Left            =   0
                  Picture         =   "FrmRequuest.frx":18F2
                  Top             =   30
                  Width           =   480
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«’‰«› „ Êﬁ⁄ «‰  ’· Õœ «·ÿ·»"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   1005
                  Left            =   3660
                  TabIndex        =   23
                  Top             =   0
                  Width           =   6465
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   6090
               Left            =   0
               TabIndex        =   24
               Top             =   2565
               Width           =   12765
               _cx             =   22516
               _cy             =   10742
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
               BackColorBkg    =   16777215
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmRequuest.frx":25BC
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
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "FrmRequuest.frx":27E6
               Height          =   315
               Left            =   390
               TabIndex        =   25
               Top             =   570
               Visible         =   0   'False
               Width           =   4500
               _ExtentX        =   7938
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
            Begin MSDataListLib.DataCombo DcbBranch1 
               Bindings        =   "FrmRequuest.frx":27FB
               Height          =   315
               Left            =   6735
               TabIndex        =   29
               Top             =   570
               Visible         =   0   'False
               Width           =   4500
               _ExtentX        =   7938
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
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   1815
               Left            =   120
               TabIndex        =   32
               Top             =   720
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   3201
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
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
               ButtonImage     =   "FrmRequuest.frx":2810
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   300
               Index           =   4
               Left            =   11175
               TabIndex        =   30
               Top             =   570
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Œ“‰"
               Height          =   300
               Index           =   2
               Left            =   4800
               TabIndex        =   26
               Top             =   570
               Visible         =   0   'False
               Width           =   1200
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   585
         Index           =   0
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8280
         Width           =   11475
         _cx             =   20241
         _cy             =   1032
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
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   570
            Left            =   3915
            TabIndex        =   12
            Top             =   75
            Width           =   5505
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   435
            Left            =   240
            TabIndex        =   13
            Top             =   135
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   767
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
            BackStyle       =   0
            ButtonImage     =   "FrmRequuest.frx":2BAA
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   480
            Left            =   1890
            TabIndex        =   14
            Top             =   105
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   847
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
            ButtonImage     =   "FrmRequuest.frx":2F44
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
            Height          =   840
            Left            =   9870
            TabIndex        =   15
            Top             =   -150
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1482
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· «·Ï  ÿ·» ‘—«¡ «·Ì"
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   1
            Left            =   14355
            TabIndex        =   17
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«›"
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   0
            Left            =   15900
            TabIndex        =   16
            Top             =   180
            Width           =   2085
         End
      End
   End
End
Attribute VB_Name = "FrmRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    FrmShowPrice.show
End Sub

Private Sub CmdPrint_Click()
    On Error GoTo ErrTrap
    Dim RequestReport As ClsRepoerts
    Dim StrSQL As String
Dim MySQL As String
  '  If SystemOptions.SysDataBaseType = AccessDataBase Then
  '      StrSQL = "Select * From"
   '     StrSQL = StrSQL + "("
      '  StrSQL = StrSQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName, "
       ' StrSQL = StrSQL + " Sum(Qty)as Qty  From RequestItems"
      '  StrSQL = StrSQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName"
     '   StrSQL = StrSQL + ")XTable"
    '    StrSQL = StrSQL + " Where XTable.Qty <= XTable.Requestlimit"
    
   ' Else
        'StrSQL = "Select * From"
      '  StrSQL = StrSQL + "("
       'StrSQL = StrSQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName, "
       'StrSQL = StrSQL + " Sum(Qty)as Qty  From RequestItems"
       ' StrSQL = StrSQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName"
       'StrSQL = StrSQL + ")XTable"
      ' StrSQL = StrSQL + " Where XTable.Qty <= (XTable.Requestlimit-1)"
   ' End If
'
 ' StrSQL = StrSQL + " Order By XTable.GroupID,XTable.ItemID "


   ' Set RequestReport = New ClsRepoerts
'
 '   If Me.Opt(0).value = True Then
   '   RequestReport.RequestItems MySQL, False
  ' ' ElseIf Me.Opt(1).value = True Then
 '   '    RequestReport.RequestItems StrSQL, True
'   ' End If
print_report
    Exit Sub
ErrTrap:
End Sub

Private Sub DataCombo1_Change()
DataCombo1_Click (0)
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If val(DataCombo1.BoundText) <> 0 Then
FillGrid2
End If
End Sub

Private Sub DcbBranch_Change()
DcbBranch_Click (0)
End Sub

Private Sub DcbBranch_Click(Area As Integer)
If val(DcbBranch.BoundText) <> 0 Then
FillGrid
End If
End Sub

Private Sub DcbBranch1_Change()
DcbBranch1_Click (0)
End Sub

Private Sub DcbBranch1_Click(Area As Integer)
If val(DcbBranch1.BoundText) <> 0 Then
FillGrid2
End If
End Sub

Private Sub DcbStore_Change()
DcbStore_Click (0)
End Sub

Private Sub DcbStore_Click(Area As Integer)
If val(DcbStore.BoundText) <> 0 Then
FillGrid
End If
End Sub

Private Sub ISButton1_Click()
FillGrid2
End Sub

Private Sub ISButton2_Click()
FillGrid
End Sub

Private Sub Label10_Click()
ListBranchSelected.Clear
Me.ListAllStore.Clear
Me.ListStoreSelect.Clear
End Sub

Private Sub Label11_Click()
If ListBranchSelected.ListIndex > -1 Then
ListBranchSelected.RemoveItem (ListBranchSelected.ListIndex)
End If

FillMylist2
End Sub

Private Sub Label14_Click()
    Dim i As Integer
    Me.ListStoreSelect2.Clear
    For i = 0 To Me.ListAllStore2.ListCount - 1
        Me.ListStoreSelect2.AddItem ListAllStore2.List(i)
        ListStoreSelect2.ItemData(i) = ListAllStore2.ItemData(i)
    Next i
End Sub

Private Sub Label15_Click()
 If Me.ListAllStore2.ListIndex > -1 Then
    Me.ListStoreSelect2.AddItem ListAllStore2.List(ListAllStore2.ListIndex)
    ListStoreSelect2.ItemData(ListStoreSelect2.NewIndex) = ListAllStore2.ItemData(ListAllStore2.ListIndex)
End If
End Sub

Private Sub Label17_Click()
 If Me.ListBranchAll2.ListIndex > -1 Then
    Me.ListBranchSelected2.AddItem ListBranchAll2.List(ListBranchAll2.ListIndex)
    ListBranchSelected2.ItemData(ListBranchSelected2.NewIndex) = ListBranchAll2.ItemData(ListBranchAll2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label18_Click()
    Dim i As Integer
    Me.ListBranchSelected2.Clear
    For i = 0 To Me.ListBranchAll2.ListCount - 1
        Me.ListBranchSelected2.AddItem ListBranchAll2.List(i)
        ListBranchSelected2.ItemData(i) = ListBranchAll2.ItemData(i)
    Next i
  
   FillMylist3
End Sub

Private Sub Label19_Click()
ListBranchSelected2.Clear
Me.ListAllStore2.Clear
Me.ListStoreSelect2.Clear
End Sub

Private Sub Label2_Click()
If ListStoreSelect2.ListIndex > -1 Then
ListStoreSelect2.RemoveItem (ListStoreSelect2.ListIndex)
End If
End Sub

Private Sub Label20_Click()
If ListBranchSelected2.ListIndex > -1 Then
ListBranchSelected2.RemoveItem (ListBranchSelected2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label3_Click()
Me.ListStoreSelect2.Clear
End Sub

Private Sub Label4_Click()

 If Me.ListBranchAll.ListIndex > -1 Then
    Me.ListBranchSelected.AddItem ListBranchAll.List(ListBranchAll.ListIndex)
    ListBranchSelected.ItemData(ListBranchSelected.NewIndex) = ListBranchAll.ItemData(ListBranchAll.ListIndex)

End If
FillMylist2
End Sub

Private Sub Label5_Click()

If ListStoreSelect.ListIndex > -1 Then
ListStoreSelect.RemoveItem (ListStoreSelect.ListIndex)
End If

End Sub
Private Sub Label6_Click()
Me.ListStoreSelect.Clear
End Sub

Private Sub Label7_Click()

    Dim i As Integer
    Me.ListStoreSelect.Clear
    For i = 0 To Me.ListAllStore.ListCount - 1
        Me.ListStoreSelect.AddItem ListAllStore.List(i)
        ListStoreSelect.ItemData(i) = ListAllStore.ItemData(i)
    Next i

End Sub

Private Sub Label8_Click()

 If Me.ListAllStore.ListIndex > -1 Then
    Me.ListStoreSelect.AddItem ListAllStore.List(ListAllStore.ListIndex)
    ListStoreSelect.ItemData(ListStoreSelect.NewIndex) = ListAllStore.ItemData(ListAllStore.ListIndex)
End If
End Sub

Private Sub Label9_Click()

    Dim i As Integer
    Me.ListBranchSelected.Clear
    For i = 0 To Me.ListBranchAll.ListCount - 1
        Me.ListBranchSelected.AddItem ListBranchAll.List(i)
        ListBranchSelected.ItemData(i) = ListBranchAll.ItemData(i)
    Next i
  
   FillMylist2
End Sub
Function FillMylist()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblBranchesData "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Me.ListBranchAll.Clear
    Me.ListBranchSelected.Clear
     Me.ListBranchAll2.Clear
    Me.ListBranchSelected2.Clear
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBranchAll.AddItem IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
            Else
                ListBranchAll.AddItem IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
            End If
              If SystemOptions.UserInterface = ArabicInterface Then
                ListBranchAll2.AddItem IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
            Else
                ListBranchAll2.AddItem IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
            End If
            ListBranchAll.ItemData(ListBranchAll.NewIndex) = IIf(IsNull(rs2("branch_id").value), 0, rs2("branch_id").value)
            ListBranchAll2.ItemData(ListBranchAll2.NewIndex) = IIf(IsNull(rs2("branch_id").value), 0, rs2("branch_id").value)
            rs2.MoveNext
        Next i
    End If
    rs2.Close
End Function
Function FillMylist2()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListBranchSelected.ListCount - 1
    ActivID = ActivID & "," & Me.ListBranchSelected.ItemData(i)
    Next i
    Me.ListAllStore.Clear
    Me.ListStoreSelect.Clear
    If ActivID = "0" Then Exit Function
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblStore where BranchId in(" & ActivID & ") "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllStore.AddItem IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
            Else
                ListAllStore.AddItem IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
            End If
            ListAllStore.ItemData(ListAllStore.NewIndex) = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
End Function
Function FillMylist3()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListBranchSelected2.ListCount - 1
    ActivID = ActivID & "," & Me.ListBranchSelected2.ItemData(i)
    Next i
    Me.ListAllStore2.Clear
    Me.ListStoreSelect2.Clear
    If ActivID = "0" Then Exit Function
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblStore where BranchId in(" & ActivID & ") "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllStore2.AddItem IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
            Else
                ListAllStore2.AddItem IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
            End If
            ListAllStore2.ItemData(ListAllStore2.NewIndex) = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
End Function

Private Sub Fg_DblClick()

    With Me.fg
     
        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ItemID"))) <> 0 Then
            Load FrmSelectData
            FrmSelectData.DCboItemName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
            FrmSelectData.txtItemCode.Text = val(.TextMatrix(.Row, .ColIndex("ItemCode")))
            FrmSelectData.show
        End If

    End With

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean
    Dim IntGridView As Integer
    LoadIcons
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetBranches Me.DcbBranch1
    Dcombos.GetStores Me.DcbStore
    Dcombos.GetStores Me.DataCombo1
    fg.WallPaper = BGround.Picture
    FillMylist
    '----------------------------------------------------------------------------
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "ShowRequest", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    '----------------------------------------------------------------------------
    IntGridView = GetSetting(SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "GridView", 0)

    If IntGridView = 0 Then
        Me.Opt(0).value = True
        Opt_Click 0
    Else
        Me.Opt(1).value = True
        Opt_Click 1
    End If

    '----------------------------------------------------------------------------
    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Exit Sub
ErrTrap:

End Sub

Private Sub ChangeLang()
Frame1.RightToLeft = False
Frame2.RightToLeft = False
Frame3.RightToLeft = False
Frame4.RightToLeft = False
Label13.Caption = "Select Branch"
Label12.Caption = "Select Store"
ISButton2.Caption = "Show"
Label21.Caption = "Select Branch"
Label16.Caption = "Select Store"
ISButton1.Caption = "Show"
    Me.Caption = "Items Request Alarm"
    LblCaption.Caption = "Items Request Alarm"
    Label1.Caption = "Items Request Alarm"
    C1Tab1.TabCaption(0) = "Items Request Alarm"
    C1Tab1.TabCaption(1) = "Items May Request Alarm"
    lbl(17).Caption = "Store"
    lbl(2).Caption = "Store"
    ELe(1).Caption = "View Type"
    Opt(0).Caption = "Table"
    Opt(1).Caption = "Tree"
    lbl(0).Caption = "No of Items"
    ChkShow.Caption = "Dont Show At Start up"
    cmdPrint.Caption = "Print"
    CmdExit.Caption = "Exit"
    lbl(4).Caption = "Branch"
    lbl(3).Caption = "Branch"
    With fg
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Requst")) = "Requst"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("StoreID")) = "Store ID"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("DefalutPrice")) = "Defalut Price"
        .TextMatrix(0, .ColIndex("RequiredQty")) = "Required Qty"

    End With
    
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("RequiredQty")) = "Required Qty"
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Requst")) = "Requst"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("StoreID")) = "Store ID"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("DefalutPrice")) = "Defalut Price"
        .TextMatrix(0, .ColIndex("OperReQst")) = "Factor"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowRequest", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowRequest", True
    End If

    If Me.Opt(0).value = True Then
        SaveSetting SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "GridView", 0
    ElseIf Opt(1).value = True Then
        SaveSetting SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "GridView", 1
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    With fg
        .Cell(flexcpPicture, 0, .ColIndex("ItemID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("ItemCode")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Qty")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("Item").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Requst")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub LblCaption_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub LoadTreeGrid()

    Dim My_SQL As String
    Dim i As Integer
    Dim IntColName As Integer
    Dim BolRtl As Boolean
    Dim RsData As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim ReCount As Long
    Dim RowNum As Long
    Dim LngParentRow As Long
    Dim DblNodeChildCount As Double

    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With Me.fg
        .Redraw = flexRDNone
        .Rows = 1

        If BolRtl = True Then
            IntColName = 1
            .AddItem "‘Ã—… «·√’‰«›"
        Else
            .AddItem "Items Tree"
            IntColName = 1
        End If

        .Rowdata(.Rows - 1) = "1G"
        .IsSubtotal(.Rows - 1) = True
        .Cell(flexcpFontBold, .Rows - 1, 1) = True
        .GridLines = flexGridFlat
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        '.NodeClosedPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeClose").Picture
        '.NodeOpenPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeOpen").Picture
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        '-----------------------------------------
        .ColHidden(.ColIndex("GroupName")) = True
        .ColPosition(.ColIndex("Name")) = 0
        '-----------------------------------------
        My_SQL = " SELECT Groups.GroupID, Groups.GroupName, Groups.ParentID " & "FROM Groups Where Groups.ParentID=1"
        Set RsData = New ADODB.Recordset
    
        RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call LoadGridTree("1G", RsData, fg, "Groups", "ParentID", "", , IntColName, vbBlue)
        '------------------------------------------------------------
        Set RsTemp = New ADODB.Recordset

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            My_SQL = "Select * From"
            My_SQL = My_SQL + "("
            My_SQL = My_SQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName, "
            My_SQL = My_SQL + " Sum(Qty)as Qty  From RequestItems"
            My_SQL = My_SQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName"
            My_SQL = My_SQL + ")XTable"
            My_SQL = My_SQL + " Where XTable.Qty <= XTable.Requestlimit"
        Else
            My_SQL = "Select * From"
            My_SQL = My_SQL + "("
            My_SQL = My_SQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName, "
            My_SQL = My_SQL + " Sum(Qty)as Qty  From RequestItems"
            My_SQL = My_SQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName"
            My_SQL = My_SQL + ")XTable"
            My_SQL = My_SQL + " Where XTable.Qty <= (XTable.Requestlimit-1)"
        End If

        RsTemp.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        '-----------------------------------------------------------
        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Me.lbl(1).Caption = RsTemp.RecordCount
            RsTemp.MoveFirst

            Do While Not RsTemp.EOF
                LngParentRow = .FindRow(CStr(RsTemp("GroupID").value) & "G", 0, -1, False, True)

                If LngParentRow > 0 Then
                    .AddItem RsTemp(IntColName).value, (LngParentRow + 1)
                    .Rowdata((LngParentRow + 1)) = RsTemp("ItemID").value & "I"
                    .RowOutlineLevel((LngParentRow + 1)) = .RowOutlineLevel(LngParentRow) + 1
                    .Cell(flexcpPicture, LngParentRow + 1, 0) = mdifrmmain.ImgLstTree.ListImages("Item").ExtractIcon
                    '-----------------------------------------------------------------
                    RowNum = (LngParentRow + 1)
                    '-----------------------------------------------------------------
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                    .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                    .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
                    .TextMatrix(RowNum, .ColIndex("DefalutPrice")) = IIf(IsNull(RsTemp("PurchasePrice").value), "0", RsTemp("PurchasePrice").value)
                    .TextMatrix(RowNum, .ColIndex("Requst")) = IIf(IsNull(RsTemp("RequestLimit").value), "0", RsTemp("RequestLimit").value)

                    If Not (IsNull(RsTemp("qty").value)) = True Then
                        .TextMatrix(RowNum, .ColIndex("qty")) = Format(RsTemp("qty").value, SystemOptions.SysDefCurrencyForamt)
                    Else
                        .TextMatrix(RowNum, .ColIndex("qty")) = 0
                    End If
                
                    .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), "0", RsTemp("StoreID").value)
                    .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "0", RsTemp("StoreName").value)
                    .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "0", RsTemp("GroupName").value)
               
                    .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
                    RsTemp.MoveNext
                End If

            Loop

        End If

        '------------------------------------------------------------
        For i = Me.fg.FixedRows To Me.fg.Rows - 1
            Dim XNode As VSFlex8UCtl.VSFlexNode
            Dim StrTemp As String

            If .IsSubtotal(i) = True Then
                Set XNode = fg.GetNode(i)

                If Not XNode Is Nothing Then
                    DblNodeChildCount = ModFgLib.GetNodeChildTotal(fg, XNode, flexSTCount)
                    StrTemp = XNode.Text & " ( " & DblNodeChildCount & " ) "
                    XNode.Text = StrTemp
                End If
            End If

        Next i

        .AutoSize 0, .Cols - 1, False
        .Redraw = True
        .Outline 1
    End With

    RsData.Close
    RsTemp.Close
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Resume
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadTableData()
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As ADODB.Recordset
    On Error GoTo hErr

    Set RsTemp = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "Select * From"
        My_SQL = My_SQL + "("
        My_SQL = My_SQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName, "
        My_SQL = My_SQL + " Sum(Qty)as Qty  From RequestItems"
        My_SQL = My_SQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupID,GroupName"
        My_SQL = My_SQL + ")XTable"
        My_SQL = My_SQL + " Where XTable.Qty <= (XTable.Requestlimit-1)"
    
    Else
        My_SQL = "Select * From"
        My_SQL = My_SQL + "("
        My_SQL = My_SQL + "Select ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupName, "
        My_SQL = My_SQL + " Sum(Qty)as Qty  From RequestItems"
        My_SQL = My_SQL + " Group by ItemID,ItemCode,ItemName,RequestLimit,PurchasePrice,StoreID,StoreName,GroupName"
        My_SQL = My_SQL + ")XTable"
        My_SQL = My_SQL + " Where XTable.Qty <= (XTable.Requestlimit-1)"
    End If

    RsTemp.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Me.lbl(1).Caption = RsTemp.RecordCount

        With fg
            .ColHidden(.ColIndex("GroupName")) = False
            .Rows = .FixedRows
            .ExplorerBar = flexExSortShowAndMove

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
                .TextMatrix(RowNum, .ColIndex("DefalutPrice")) = IIf(IsNull(RsTemp("PurchasePrice").value), "0", RsTemp("PurchasePrice").value)
                .TextMatrix(RowNum, .ColIndex("Requst")) = IIf(IsNull(RsTemp("RequestLimit").value), "0", RsTemp("RequestLimit").value)

                If Not (IsNull(RsTemp("qty").value)) = True Then
                    .TextMatrix(RowNum, .ColIndex("qty")) = Format(RsTemp("qty").value, SystemOptions.SysDefQuantityFormat)
                Else
                    .TextMatrix(RowNum, .ColIndex("qty")) = 0
                End If

                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), "0", RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "0", RsTemp("StoreName").value)
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "0", RsTemp("GroupName").value)
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.lbl(1).Caption = ""
    End If

    Exit Sub
hErr:
End Sub
Private Sub FillGrid2()
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As ADODB.Recordset
    On Error GoTo hErr
   With VSFlexGrid1
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .ExplorerBar = flexExSortShowAndMove
     End With
    Set RsTemp = New ADODB.Recordset
My_SQL = " SELECT     Xb.Qty, Xb.ConsuRateLowQty, Xb.GroupName, Xb.GroupNamee, Xb.ItemName, Xb.ItemNamee, Xb.Fullcode, Xb.UnitName, Xb.UnitNamee, Xb.GropFullcode,"
My_SQL = My_SQL & "                      Xb.StoreId , Xb.ItemID, Xb.ConsuRate, Xb.UnitFactor, Xb.storename, Xb.storenamee, Xb.code, BX.QNty ,xb.OperReQst"
My_SQL = My_SQL & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
My_SQL = My_SQL & "                                              dbo.Groups.Fullcode AS GropFullcode, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
My_SQL = My_SQL & "                                              dbo.TblStore.code,dbo.TblSettsRequestLimitDet.OperReQst"
My_SQL = My_SQL & "                        FROM         dbo.TblStore INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblUnites INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblItems INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Groups INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
My_SQL = My_SQL & "                                              dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
My_SQL = My_SQL & "                                              dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID"
My_SQL = My_SQL & "                        Where (dbo.TblSettsRequestLimitDet.Typ = 0)"
My_SQL = My_SQL & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code,dbo.TblSettsRequestLimitDet.OperReQst) Xb INNER JOIN"
My_SQL = My_SQL & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
My_SQL = My_SQL & "                             FROM         dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
My_SQL = My_SQL & "                             GROUP BY Item_ID, StoreID) BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID and BX.Item_ID in (select ItemID from TblSettsRequestLimitDet)"
My_SQL = My_SQL & " where 1=1"
   Dim SotreID As String
   Dim i As Integer
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect2.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect2.ItemData(i)
    Next i
    If SotreID <> "0" Then
     My_SQL = My_SQL & " and   Xb.StoreID in (" & SotreID & ")"
     End If
'   If val(Me.DcbBranch1.BoundText) <> 0 Then
'   My_SQL = My_SQL & " and   Xb.StoreID in(select StoreID"
'   My_SQL = My_SQL & " From dbo.TblStore"
'   My_SQL = My_SQL & " WHERE     (BranchId = " & val(DcbBranch1.BoundText) & ") OR"
'    My_SQL = My_SQL & "  (BranchId IS NULL) )"
'End If
'    If SystemOptions.usertype = UserAdminAll Then
'  If val(DataCombo1.BoundText) <> 0 Then
'        My_SQL = My_SQL & " and   Xb.StoreID =" & val(DataCombo1.BoundText) & ""
' End If
' Else
'         My_SQL = My_SQL & " and   Xb.StoreID =" & val(DataCombo1.BoundText) & ""
' End If
 My_SQL = My_SQL & " and  (BX.QNty-xb.OperReQst-Xb.Qty)>0 "
    RsTemp.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
  Dim UnitFactor As Double
  Dim Requst As Double
  Dim Qty As Double
  Dim RequiredQty As Double
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Me.lbl(1).Caption = RsTemp.RecordCount

        With VSFlexGrid1
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .ExplorerBar = flexExSortShowAndMove

            For ReCount = 1 To RsTemp.RecordCount
            UnitFactor = IIf(IsNull(RsTemp("UnitFactor").value), 0, RsTemp("UnitFactor").value)
            Requst = IIf(IsNull(RsTemp("Qty").value), 0, RsTemp("Qty").value)
            Qty = IIf(IsNull(RsTemp("QNty").value), "0", RsTemp("QNty").value)
            If UnitFactor <> 0 Then
            Qty = Round(Qty / UnitFactor, 2)
            End If
            RequiredQty = Requst - Qty
            If RequiredQty > 0 Then
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                .TextMatrix(RowNum, .ColIndex("UnitFactor")) = UnitFactor
                .TextMatrix(RowNum, .ColIndex("Requst")) = Requst
                .TextMatrix(RowNum, .ColIndex("qty")) = Qty
                .TextMatrix(RowNum, .ColIndex("RequiredQty")) = RequiredQty
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitName").value), "", RsTemp("UnitName").value)
                Else
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitNamee").value), "", RsTemp("UnitNamee").value)
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupNamee").value), "", RsTemp("GroupNamee").value)
                 .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreNamee").value), "", RsTemp("StoreNamee").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemNamee").value), "", RsTemp("ItemNamee").value)
                End If
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), "", RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                .TextMatrix(RowNum, .ColIndex("OperReQst")) = IIf(IsNull(RsTemp("OperReQst").value), 0, RsTemp("OperReQst").value)
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), 0, RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
                End If
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.lbl(1).Caption = ""
    End If

    Exit Sub
hErr:
End Sub
Private Sub FillGrid()
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As ADODB.Recordset
    Dim UnitFactor As Double
    Dim Requst As Double
    Dim Qty As Double
    Dim RequiredQty As Double
    On Error GoTo hErr

    Set RsTemp = New ADODB.Recordset
     With fg
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .ExplorerBar = flexExSortShowAndMove
     End With
 ' My_SQL = " SELECT     TOP 100 PERCENT dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
 ' My_SQL = My_SQL & "                    dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode AS GropFullcode,"
 ' My_SQL = My_SQL & "                    dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
 ' My_SQL = My_SQL & "                    dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID) AS QNty, dbo.TblSettsRequestLimitDet.UnitFactor,"
 ' My_SQL = My_SQL & "                    dbo.TblStore.storename , dbo.TblStore.storenamee, dbo.TblStore.code"
 ' My_SQL = My_SQL & " FROM         dbo.Groups RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblItems RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblUnites RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblStore RIGHT OUTER JOIN"
 ' My_SQL = My_SQL & "                    dbo.TblSettsRequestLimitDet INNER JOIN"
'  My_SQL = My_SQL & "                    dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
'  My_SQL = My_SQL & "                    dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID ON dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON"
'  My_SQL = My_SQL & "                    dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID"
'  My_SQL = My_SQL & " Where (dbo.TblSettsRequestLimitDet.typ = 0)"


'My_SQL = My_SQL & " GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.ItemName, "
'My_SQL = My_SQL & "                      dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode, dbo.TblSettsRequestLimitDet.StoreID,"
'My_SQL = My_SQL & "                      dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName,"
'My_SQL = My_SQL & "                      dbo.TblStore.storenamee , dbo.TblStore.code, dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreId, dbo.TblSettsRequestLimitDet.ItemID)"
'My_SQL = My_SQL & " ORDER BY dbo.TblSettsRequestLimitDet.StoreID"
My_SQL = " SELECT     Xb.Qty, Xb.ConsuRateLowQty, Xb.GroupName, Xb.GroupNamee, Xb.ItemName, Xb.ItemNamee, Xb.Fullcode, Xb.UnitName, Xb.UnitNamee, Xb.GropFullcode,"
My_SQL = My_SQL & "                      Xb.StoreId , Xb.ItemID, Xb.ConsuRate, Xb.UnitFactor, Xb.storename, Xb.storenamee, Xb.code, BX.QNty"
My_SQL = My_SQL & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
My_SQL = My_SQL & "                                              dbo.Groups.Fullcode AS GropFullcode, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
My_SQL = My_SQL & "                                              dbo.TblStore.code"
My_SQL = My_SQL & "                        FROM         dbo.TblStore INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblUnites INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblItems INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Groups INNER JOIN"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet INNER JOIN"
My_SQL = My_SQL & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
My_SQL = My_SQL & "                                              dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
My_SQL = My_SQL & "                                              dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID"
My_SQL = My_SQL & "                        Where (dbo.TblSettsRequestLimitDet.Typ = 0)"
My_SQL = My_SQL & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
My_SQL = My_SQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
My_SQL = My_SQL & "                                              dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code) Xb INNER JOIN"
My_SQL = My_SQL & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
My_SQL = My_SQL & "                             FROM         dbo.Transactions INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
My_SQL = My_SQL & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
My_SQL = My_SQL & "                             GROUP BY Item_ID, StoreID)  BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID and BX.Item_ID in (select ItemID from TblSettsRequestLimitDet)"
My_SQL = My_SQL & " where 1=1 "
 '   If SystemOptions.usertype = UserAdminAll Then
 ' If val(DcbStore.BoundText) <> 0 Then
 '       My_SQL = My_SQL & " and   Xb.StoreID =" & val(DcbStore.BoundText) & " "
'
' End If
' Else
'         My_SQL = My_SQL & " and   Xb.StoreID =" & val(DcbStore.BoundText) & ""
' End If
'   If val(Me.DcbBranch.BoundText) <> 0 Then
'   My_SQL = My_SQL & " and   Xb.StoreID in(select StoreID"
'   My_SQL = My_SQL & " From dbo.TblStore"
'   My_SQL = My_SQL & " WHERE     (BranchId = " & val(DcbBranch.BoundText) & ") OR"
'    My_SQL = My_SQL & "  (BranchId IS NULL) )"
'End If
   Dim SotreID As String
   Dim i As Integer
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect.ItemData(i)
    Next i
    If SotreID <> "0" Then
     My_SQL = My_SQL & " and   Xb.StoreID in (" & SotreID & ")"
     End If
''/////
    RsTemp.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

     
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Me.lbl(1).Caption = RsTemp.RecordCount

        With fg
            .ColHidden(.ColIndex("GroupName")) = False
             .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .ExplorerBar = flexExSortShowAndMove

            For ReCount = 1 To RsTemp.RecordCount
                UnitFactor = IIf(IsNull(RsTemp("UnitFactor").value), 0, RsTemp("UnitFactor").value)
                Requst = IIf(IsNull(RsTemp("Qty").value), 0, RsTemp("Qty").value)
                Qty = IIf(IsNull(RsTemp("QNty").value), "0", RsTemp("QNty").value)
                If UnitFactor <> 0 Then
                   Qty = Round(Qty / UnitFactor, 2)
                End If
                RequiredQty = Requst - Qty
                If RequiredQty > 0 Then
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                .TextMatrix(RowNum, .ColIndex("UnitFactor")) = UnitFactor
                .TextMatrix(RowNum, .ColIndex("Requst")) = Requst
                .TextMatrix(RowNum, .ColIndex("qty")) = Qty
                .TextMatrix(RowNum, .ColIndex("RequiredQty")) = RequiredQty
                
               ' .TextMatrix(RowNum, .ColIndex("UnitFactor")) = IIf(IsNull(RsTemp("UnitFactor").value), 0, RsTemp("UnitFactor").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitName").value), "", RsTemp("UnitName").value)
                Else
                .TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(RsTemp("UnitNamee").value), "", RsTemp("UnitNamee").value)
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupNamee").value), "", RsTemp("GroupNamee").value)
                 .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreNamee").value), "", RsTemp("StoreNamee").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemNamee").value), "", RsTemp("ItemNamee").value)
                End If
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), "", RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                
               ' .TextMatrix(RowNum, .ColIndex("Requst")) = IIf(IsNull(RsTemp("Qty").value), 0, RsTemp("Qty").value)
               '   .TextMatrix(RowNum, .ColIndex("qty")) = IIf(IsNull(RsTemp("QNty").value), 0, RsTemp("QNty").value)
'If val(.TextMatrix(RowNum, .ColIndex("UnitFactor"))) <> 0 Then
'   .TextMatrix(RowNum, .ColIndex("qty")) = Round(val(.TextMatrix(RowNum, .ColIndex("qty"))) / val(.TextMatrix(RowNum, .ColIndex("UnitFactor"))), 2)
'End If
                .TextMatrix(RowNum, .ColIndex("StoreID")) = IIf(IsNull(RsTemp("StoreID").value), 0, RsTemp("StoreID").value)
                .TextMatrix(RowNum, .ColIndex("StoreName")) = IIf(IsNull(RsTemp("StoreName").value), "", RsTemp("StoreName").value)
'                .TextMatrix(RowNum, .ColIndex("RequiredQty")) = val(.TextMatrix(RowNum, .ColIndex("Requst"))) - val(.TextMatrix(RowNum, .ColIndex("Qty")))
                
                .TextMatrix(RowNum, .ColIndex("GroupName")) = IIf(IsNull(RsTemp("GroupName").value), "", RsTemp("GroupName").value)
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
               End If
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    Else
        Me.lbl(1).Caption = ""
    End If

    Exit Sub
hErr:
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
  
'  MySQL = "SELECT     TOP 100 PERCENT dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
'  MySQL = MySQL & "                     dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode AS GropFullcode,"
'  MySQL = MySQL & "                    dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
'  MySQL = MySQL & "                    dbo.GeQtyofStore(dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID) AS QNty, dbo.TblSettsRequestLimitDet.UnitFactor,"
'  MySQL = MySQL & "                    dbo.TblStore.storename , dbo.TblStore.StoreNamee, dbo.TblStore.Code"
'  MySQL = MySQL & "   FROM         dbo.Groups RIGHT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblItems RIGHT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblUnites RIGHT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.Transaction_Details RIGHT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblSettsRequestLimitDet LEFT OUTER JOIN"
'  MySQL = MySQL & "                    dbo.TblStore ON dbo.TblSettsRequestLimitDet.StoreID = dbo.TblStore.StoreID ON dbo.Transaction_Details.Item_ID = dbo.TblSettsRequestLimitDet.ItemID ON"
'  MySQL = MySQL & "                    dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
'  MySQL = MySQL & "                    dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID"
'  MySQL = MySQL & "  Where (dbo.TblSettsRequestLimitDet.typ = 0)"
MySQL = " SELECT     Xb.Qty, Xb.ConsuRateLowQty, Xb.GroupName, Xb.GroupNamee, Xb.ItemName, Xb.ItemNamee, Xb.Fullcode, Xb.UnitName, Xb.UnitNamee, Xb.GropFullcode,"
MySQL = MySQL & "                      Xb.StoreId , Xb.ItemID, Xb.ConsuRate, Xb.UnitFactor, Xb.storename, Xb.storenamee, Xb.code, BX.QNty"
MySQL = MySQL & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
MySQL = MySQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
MySQL = MySQL & "                                              dbo.Groups.Fullcode AS GropFullcode, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID,"
MySQL = MySQL & "                                              dbo.TblSettsRequestLimitDet.ConsuRate, dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
MySQL = MySQL & "                                              dbo.TblStore.code"
MySQL = MySQL & "                        FROM         dbo.TblStore INNER JOIN"
MySQL = MySQL & "                                              dbo.TblUnites INNER JOIN"
MySQL = MySQL & "                                              dbo.TblItems INNER JOIN"
MySQL = MySQL & "                                              dbo.Groups INNER JOIN"
MySQL = MySQL & "                                              dbo.TblSettsRequestLimitDet INNER JOIN"
MySQL = MySQL & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID ON"
MySQL = MySQL & "                                              dbo.Groups.GroupID = dbo.TblSettsRequestLimitDet.GroupID ON dbo.TblItems.ItemID = dbo.TblSettsRequestLimitDet.ItemID ON"
MySQL = MySQL & "                                              dbo.TblUnites.UnitID = dbo.TblSettsRequestLimitDet.UnitID ON dbo.TblStore.StoreID = dbo.TblSettsRequestLimitDet.StoreID"
MySQL = MySQL & "                        Where (dbo.TblSettsRequestLimitDet.Typ = 0)"
MySQL = MySQL & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.ConsuRateLowQty, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
MySQL = MySQL & "                                              dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Groups.Fullcode,"
MySQL = MySQL & "                                              dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.ConsuRate,"
MySQL = MySQL & "                                              dbo.TblSettsRequestLimitDet.UnitFactor, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code) Xb INNER JOIN"
MySQL = MySQL & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
MySQL = MySQL & "                             FROM         dbo.Transactions INNER JOIN"
MySQL = MySQL & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
MySQL = MySQL & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
MySQL = MySQL & "                             GROUP BY Item_ID, StoreID)  BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID and BX.Item_ID in (select ItemID from TblSettsRequestLimitDet)"
MySQL = MySQL & " where 1=1 "
'    If SystemOptions.usertype = UserAdminAll Then
'  If val(DcbStore.BoundText) <> 0 Then
'        MySQL = MySQL & " and   dbo.TblSettsRequestLimitDet.StoreId=" & val(DcbStore.BoundText) & ""
' End If
' Else
'         MySQL = MySQL & " and   dbo.TblSettsRequestLimitDet.StoreId=" & val(DcbStore.BoundText) & ""
' End If
'    If val(Me.DcbBranch.BoundText) <> 0 Then
'   MySQL = MySQL & " and   dbo.TblSettsRequestLimitDet.StoreId in(select StoreID"
'   MySQL = MySQL & " From dbo.TblStore"
'   MySQL = MySQL & " WHERE     (BranchId = " & val(DcbBranch.BoundText) & ") OR"
'   MySQL = MySQL & "  (BranchId IS NULL) )"
'End If
   Dim SotreID As String
   Dim My_SQL As String
   Dim i As Integer
    SotreID = "0"
    For i = 0 To Me.ListBranchSelected.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect.ItemData(i)
    Next i
    If SotreID <> "0" Then
     My_SQL = My_SQL & " and   Xb.StoreID in (" & SotreID & ")"
     End If
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\" & "RequestItems.rpt"
        Else
            StrFileName = App.path & "\REPORTS\" & "RequestItems.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

End Function

Private Sub Opt_Click(Index As Integer)

    If Opt(0).value = True Then
    FillGrid
    FillGrid2
       ' LoadTableData
    ElseIf Opt(1).value = True Then
    FillGrid
    FillGrid2
      '  LoadTreeGrid
    End If

    Exit Sub
End Sub

