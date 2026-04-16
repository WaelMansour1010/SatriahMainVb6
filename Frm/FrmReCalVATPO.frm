VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReCalVATPO 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10950
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   12615
   Icon            =   "FrmReCalVATPO.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   12615
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
   Begin C1SizerLibCtl.C1Elastic Main 
      Height          =   10950
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   12615
      _cx             =   22251
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
         Height          =   8175
         Left            =   0
         TabIndex        =   43
         Top             =   1440
         Width           =   12615
         _cx             =   22251
         _cy             =   14420
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
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   14871017
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "»Ì«‰«  «·ﬁÌ„… «·„÷«›…|»Ì«‰«  «·„»Ì⁄« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7800
            Left            =   45
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   45
            Width           =   12525
            _cx             =   22093
            _cy             =   13758
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   1200
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   -30
               Width           =   12465
               _cx             =   21987
               _cy             =   2117
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
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   1005
                  Left            =   5160
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   46
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   5865
               End
               Begin ImpulseButton.ISButton ShowBtn 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   47
                  Top             =   840
                  Width           =   1770
                  _ExtentX        =   3122
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmReCalVATPO.frx":6852
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin MSComCtl2.DTPicker FrmDate 
                  Height          =   345
                  Left            =   2160
                  TabIndex        =   48
                  Top             =   135
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   345
                  Left            =   2160
                  TabIndex        =   49
                  Top             =   660
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   106
                  Top             =   120
                  Width           =   1770
                  _ExtentX        =   3122
                  _ExtentY        =   503
                  Caption         =   "„—«Ã⁄Â «·ﬁÌÊœ"
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
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   107
                  Top             =   480
                  Width           =   1770
                  _ExtentX        =   3122
                  _ExtentY        =   503
                  Caption         =   "„—«Ã⁄Â «·›Ê« Ì—"
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
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   315
                  Index           =   1
                  Left            =   4050
                  TabIndex        =   52
                  Top             =   660
                  Width           =   900
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   315
                  Index           =   0
                  Left            =   4050
                  TabIndex        =   51
                  Top             =   135
                  Width           =   900
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   585
                  Index           =   21
                  Left            =   11175
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   390
                  Width           =   1170
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   2820
               Left            =   0
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1245
               Width           =   12525
               _cx             =   22093
               _cy             =   4974
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
               Begin VSFlex8UCtl.VSFlexGrid FG 
                  Height          =   1950
                  Left            =   120
                  TabIndex        =   54
                  Top             =   405
                  Width           =   12285
                  _cx             =   21669
                  _cy             =   3440
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
                  Rows            =   1
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReCalVATPO.frx":D0B4
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   0
                  Left            =   10755
                  TabIndex        =   55
                  Top             =   2430
                  Visible         =   0   'False
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":D183
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   330
                  Index           =   10
                  Left            =   240
                  TabIndex        =   58
                  Top             =   2400
                  Width           =   3570
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì ﬁ „"
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Index           =   9
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   2400
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   3
                  Left            =   5640
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1875
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   11
               Left            =   0
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   7170
               Width           =   12495
               _cx             =   22040
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
                  Caption         =   "Õ–› ﬁÌœ  Ò ﬁ „"
                  Height          =   480
                  Left            =   7005
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   120
                  Width           =   2790
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "≈‰‘«¡ ﬁÌœ Ò ﬁ „"
                  Height          =   480
                  Left            =   10005
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   120
                  Width           =   2430
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   8295
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   2070
               End
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   480
                  Left            =   2550
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   135
                  Width           =   3240
               End
               Begin VB.CommandButton Command9 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   480
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   135
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   405
                  Index           =   35
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   255
                  Width           =   1155
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   2970
               Left            =   0
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   4080
               Width           =   12525
               _cx             =   22093
               _cy             =   5239
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   2100
                  Left            =   120
                  TabIndex        =   67
                  Top             =   405
                  Width           =   12285
                  _cx             =   21669
                  _cy             =   3704
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
                  Rows            =   1
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReCalVATPO.frx":D71D
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   1
                  Left            =   15765
                  TabIndex        =   68
                  Top             =   5700
                  Visible         =   0   'False
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   397
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":D7EC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   2
                  Left            =   10815
                  TabIndex        =   69
                  Top             =   2550
                  Visible         =   0   'False
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":DD86
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   330
                  Index           =   12
                  Left            =   240
                  TabIndex        =   72
                  Top             =   2550
                  Width           =   3570
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã‰«·Ì ﬁ „"
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Index           =   11
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   2550
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·ﬁÌ„… «·„÷«›… ·„—œÊœ«  «·„»Ì⁄« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Index           =   6
                  Left            =   5280
                  TabIndex        =   70
                  Top             =   0
                  Width           =   2595
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   7800
            Left            =   13260
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   45
            Width           =   12525
            _cx             =   22093
            _cy             =   13758
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   1200
               Left            =   13425
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   2730
               Visible         =   0   'False
               Width           =   12465
               _cx             =   21987
               _cy             =   2117
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
               Begin VB.TextBox TxtRemarks2 
                  Alignment       =   1  'Right Justify
                  Height          =   1005
                  Left            =   300
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   75
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   360
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   1005
                  Left            =   0
                  TabIndex        =   76
                  Top             =   135
                  Width           =   90
                  _ExtentX        =   159
                  _ExtentY        =   1773
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
                  ButtonImage     =   "FrmReCalVATPO.frx":E320
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  LowerToggledContent=   0   'False
               End
               Begin MSComCtl2.DTPicker FrmDate2 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   77
                  Top             =   135
                  Width           =   90
                  _ExtentX        =   159
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker ToDate2 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   78
                  Top             =   660
                  Width           =   90
                  _ExtentX        =   159
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   585
                  Index           =   15
                  Left            =   690
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   390
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   315
                  Index           =   14
                  Left            =   240
                  TabIndex        =   80
                  Top             =   135
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   315
                  Index           =   13
                  Left            =   240
                  TabIndex        =   79
                  Top             =   660
                  Width           =   60
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   4020
               Left            =   0
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   45
               Width           =   12525
               _cx             =   22093
               _cy             =   7091
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
                  Height          =   3150
                  Left            =   0
                  TabIndex        =   83
                  Top             =   315
                  Width           =   12465
                  _cx             =   21987
                  _cy             =   5556
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
                  Rows            =   1
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReCalVATPO.frx":14B82
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   300
                  Index           =   3
                  Left            =   11235
                  TabIndex        =   84
                  Top             =   3540
                  Visible         =   0   'False
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":14CEA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì ﬁ»· ﬁ „"
                  ForeColor       =   &H000000FF&
                  Height          =   360
                  Index           =   24
                  Left            =   9645
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   3615
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   360
                  Index           =   23
                  Left            =   8235
                  TabIndex        =   102
                  Top             =   3615
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„»Ì⁄« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   435
                  Index           =   18
                  Left            =   2010
                  TabIndex        =   87
                  Top             =   0
                  Width           =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì »⁄œ ﬁ „"
                  Height          =   360
                  Index           =   17
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   3615
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   360
                  Index           =   16
                  Left            =   0
                  TabIndex        =   85
                  Top             =   3630
                  Width           =   1980
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   0
               Left            =   0
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   7170
               Width           =   12495
               _cx             =   22040
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
               Begin VB.CommandButton Command4 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   360
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   240
                  Width           =   1440
               End
               Begin VB.TextBox TxtNoteSerial2 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   480
                  Left            =   1500
                  Locked          =   -1  'True
                  TabIndex        =   92
                  Top             =   135
                  Width           =   2400
               End
               Begin VB.TextBox TxtNoteID2 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   10965
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "≈‰‘«¡ ﬁÌœ «·„»Ì⁄« "
                  Height          =   480
                  Left            =   8325
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   120
                  Width           =   2160
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Õ–› ﬁÌœ «·„»Ì⁄« "
                  Height          =   480
                  Left            =   5340
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   135
                  Width           =   2715
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   405
                  Index           =   0
                  Left            =   4050
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   255
                  Width           =   1080
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic12 
               Height          =   2970
               Left            =   0
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   4080
               Width           =   12525
               _cx             =   22093
               _cy             =   5239
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   225
                  Index           =   4
                  Left            =   960
                  TabIndex        =   96
                  Top             =   5820
                  Visible         =   0   'False
                  Width           =   120
                  _ExtentX        =   212
                  _ExtentY        =   397
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":15284
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   5
                  Left            =   11205
                  TabIndex        =   97
                  Top             =   2610
                  Visible         =   0   'False
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "FrmReCalVATPO.frx":1581E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
                  Height          =   2130
                  Left            =   0
                  TabIndex        =   101
                  Top             =   360
                  Width           =   12465
                  _cx             =   21987
                  _cy             =   3757
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
                  Rows            =   1
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReCalVATPO.frx":15DB8
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì ﬁ»· ﬁ „"
                  ForeColor       =   &H000000FF&
                  Height          =   360
                  Index           =   26
                  Left            =   9315
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   2565
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   360
                  Index           =   25
                  Left            =   8025
                  TabIndex        =   104
                  Top             =   2535
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—œÊœ«  «·„»Ì⁄« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Index           =   22
                  Left            =   2490
                  TabIndex        =   100
                  Top             =   0
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì »⁄œ ﬁ „"
                  Height          =   330
                  Index           =   20
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   2610
                  Width           =   2190
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   330
                  Index           =   19
                  Left            =   240
                  TabIndex        =   98
                  Top             =   2580
                  Width           =   1170
               End
            End
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   0
         Width           =   12555
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   31
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmReCalVATPO.frx":15F29
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   32
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmReCalVATPO.frx":162C3
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   33
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmReCalVATPO.frx":1665D
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   34
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmReCalVATPO.frx":169F7
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Õ ”«» «·ﬁÌ„… «·„÷«›… ·‰ﬁ«ÿ «·»Ì⁄"
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
            Height          =   495
            Index           =   2
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   120
            Width           =   5520
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   660
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   10320
         Width           =   12525
         _cx             =   22093
         _cy             =   1164
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   375
            Left            =   10965
            TabIndex        =   14
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
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
            ButtonImage     =   "FrmReCalVATPO.frx":16D91
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   375
            Left            =   7680
            TabIndex        =   15
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   1365
            _ExtentX        =   2408
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
            ButtonImage     =   "FrmReCalVATPO.frx":1D5F3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   375
            Left            =   9390
            TabIndex        =   16
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   1245
            _ExtentX        =   2196
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
            ButtonImage     =   "FrmReCalVATPO.frx":1D98D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   375
            Left            =   5985
            TabIndex        =   17
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   1485
            _ExtentX        =   2619
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
            ButtonImage     =   "FrmReCalVATPO.frx":241EF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   375
            Left            =   4290
            TabIndex        =   18
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
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
            ButtonImage     =   "FrmReCalVATPO.frx":24589
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   375
            Left            =   -90
            TabIndex        =   19
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   1245
            _ExtentX        =   2196
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
            ButtonImage     =   "FrmReCalVATPO.frx":24B23
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   450
            Left            =   3060
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   794
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmReCalVATPO.frx":24EBD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   375
            Left            =   1470
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "FrmReCalVATPO.frx":2B71F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   660
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   9600
         Width           =   12510
         _cx             =   22066
         _cy             =   1164
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   450
            Left            =   360
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   105
            Width           =   4065
            _cx             =   7170
            _cy             =   794
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   255
               Index           =   0
               Left            =   2955
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   255
               Index           =   1
               Left            =   1050
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   2130
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   135
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4530
            TabIndex        =   28
            Top             =   105
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   270
            Index           =   8
            Left            =   10860
            TabIndex        =   29
            Top             =   105
            Width           =   1470
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   780
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   12570
         _cx             =   22172
         _cy             =   1376
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9195
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   255
            Width           =   1905
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   330
            Left            =   6270
            TabIndex        =   38
            Top             =   255
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   582
            _Version        =   393216
            Format          =   93585409
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmReCalVATPO.frx":2BAB9
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   255
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   300
            Index           =   2
            Left            =   8055
            TabIndex        =   42
            Top             =   255
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   300
            Index           =   4
            Left            =   11460
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   255
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   300
            Index           =   7
            Left            =   5190
            TabIndex        =   40
            Top             =   255
            Width           =   1155
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmReCalVATPO.frx":2BACE
      Left            =   18360
      List            =   "FrmReCalVATPO.frx":2BADE
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   5
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   18360
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   18480
      Top             =   3720
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
            Picture         =   "FrmReCalVATPO.frx":2BAF7
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2BE91
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2C22B
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2C5C5
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2C95F
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2CCF9
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2D093
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReCalVATPO.frx":2D62D
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmReCalVATPO.frx":2D9C7
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmReCalVATPO.frx":34229
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ButtonImage     =   "FrmReCalVATPO.frx":3AA8B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»Ì⁄« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   900
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
      Left            =   18360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmReCalVATPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim RevenueAccount As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    Dim i As Integer
    Dim SumSal As Double
    Dim SumRet As Double
    SumSal = 0
    SumRet = 0
    IntCounter = 0
    With fg
        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("TotalVAT"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                SumSal = SumSal + val(.TextMatrix(i, .ColIndex("TotalVAT")))
            End If
        Next i
    End With
        IntCounter = 0
    With VSFlexGrid1
        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("TotalVAT"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                SumRet = SumRet + val(.TextMatrix(i, .ColIndex("TotalVAT")))
            End If
        Next i
    End With
    lbl(10).Caption = Round(SumSal, 2)
    lbl(12).Caption = Round(SumRet, 2)
End Sub
Private Sub ReLineGrid2()
    Dim IntCounter As Integer
    Dim i As Integer
    Dim SumSal As Double
    Dim SumRet As Double
    SumSal = 0
    SumRet = 0
    IntCounter = 0
    With VSFlexGrid2
        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                SumSal = SumSal + val(.TextMatrix(i, .ColIndex("Total")))
            End If
        Next i
    End With
        IntCounter = 0
    With VSFlexGrid3
        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                SumRet = SumRet + val(.TextMatrix(i, .ColIndex("Total")))
            End If
        Next i
    End With
    lbl(16).Caption = Round(SumSal, 2)
    lbl(19).Caption = Round(SumRet, 2)
    
   lbl(23).Caption = lbl(16).Caption - lbl(10).Caption
 lbl(25).Caption = lbl(19).Caption - lbl(12).Caption
     

End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 2
RemoveGridRow2
Case 3
RemoveGridRow3
Case 5
RemoveGridRow4
End Select
End If
End Sub

Private Sub Command1_Click()
If Me.TxtModFlg.Text = "R" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID2.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID2.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update TblReCalVATPO set NoteID2=null ,NoteSerial2=null where ID=" & val(TxtSerial1.Text) & " "
        RsSavRec.Requery
         FindRec val(TxtSerial1.Text)
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
End Sub

Private Sub Command2_Click()
If Me.TxtModFlg.Text = "R" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update TblReCalVATPO set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1.Text) & " "
        RsSavRec.Requery
         FindRec val(TxtSerial1.Text)
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
End Sub

Private Sub Command3_Click()
If TxtNoteSerial2.Text = "" Then
createVoucher2
FindRec val(TxtSerial1.Text)
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
        Else
            MsgBox "Done"
        End If
End If
End Sub

Private Sub Command4_Click()
ShowGL_cc Me.TxtNoteSerial2.Text, , 200
End Sub

Private Sub Command5_Click()
If TxtNoteSerial.Text = "" Then
createVoucher
FindRec val(TxtSerial1.Text)
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
        Else
            MsgBox "Done"
        End If
End If
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub



    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    FrmDate.value = Date
    ToDate.value = Date
    conection = "select * from TblReCalVATPO order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    BtnLast_Click

    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        'ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
ErrTrap:
End Sub
Public Sub FiLLRec()

  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblReCalVATPODet Where RecalVATID =" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("FrmDate").value = FrmDate.value
    RsSavRec.Fields("ToDate").value = ToDate.value
    RsSavRec.Fields("SalesTotal").value = val(lbl(10).Caption)
    RsSavRec.Fields("RetSalesTotal").value = val(lbl(12).Caption)
    RsSavRec.Fields("Total").value = val(lbl(16).Caption)
    RsSavRec.Fields("Total1").value = val(lbl(19).Caption)
    RsSavRec.Fields("Remarks2").value = TxtRemarks2.Text
    RsSavRec.Fields("FrmDate2").value = FrmDate2.value
    RsSavRec.Fields("ToDate2").value = ToDate2.value
    RsSavRec.update
  
''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblReCalVATPODet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With fg
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("TotalVAT"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("RecalVATID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("TypeTrans").value = 0
                 RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                 RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                 RsDevsub("TotalVAT").value = IIf((.TextMatrix(i, .ColIndex("TotalVAT"))) = "", Null, val((.TextMatrix(i, .ColIndex("TotalVAT")))))
                 RsDevsub.update
      End If
     Next i
    End With
    ''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblReCalVATPODet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("TotalVAT"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("RecalVATID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("TypeTrans").value = 1
                 RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                 RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                 RsDevsub("TotalVAT").value = IIf((.TextMatrix(i, .ColIndex("TotalVAT"))) = "", Null, val((.TextMatrix(i, .ColIndex("TotalVAT")))))
                 RsDevsub.update
      End If
     Next i
    End With
        ''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblReCalVATPODet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid2
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("RecalVATID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("TypeTrans").value = 2
                 RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                 RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                 RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val((.TextMatrix(i, .ColIndex("Total")))))
                 RsDevsub("Credit").value = IIf((.TextMatrix(i, .ColIndex("Credit"))) = "", Null, val((.TextMatrix(i, .ColIndex("Credit")))))
                 RsDevsub("Cash").value = IIf((.TextMatrix(i, .ColIndex("Cash"))) = "", Null, val((.TextMatrix(i, .ColIndex("Cash")))))
                 RsDevsub("BoxID").value = IIf((.TextMatrix(i, .ColIndex("BoxID"))) = "", Null, val((.TextMatrix(i, .ColIndex("BoxID")))))
                 RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val((.TextMatrix(i, .ColIndex("EmpID")))))
                 RsDevsub.update
      End If
     Next i
    End With
            ''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblReCalVATPODet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid3
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("RecalVATID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("TypeTrans").value = 3
                 RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                 RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
                 RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val((.TextMatrix(i, .ColIndex("Total")))))
                 RsDevsub("Credit").value = IIf((.TextMatrix(i, .ColIndex("Credit"))) = "", Null, val((.TextMatrix(i, .ColIndex("Credit")))))
                 RsDevsub("Cash").value = IIf((.TextMatrix(i, .ColIndex("Cash"))) = "", Null, val((.TextMatrix(i, .ColIndex("Cash")))))
                 RsDevsub("BoxID").value = IIf((.TextMatrix(i, .ColIndex("BoxID"))) = "", Null, val((.TextMatrix(i, .ColIndex("BoxID")))))
                 RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val((.TextMatrix(i, .ColIndex("EmpID")))))
                 RsDevsub.update
      End If
     Next i
    End With
    
UpdateFlg 1
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
    FrmDate.value = IIf(IsNull(RsSavRec.Fields("FrmDate").value), Date, RsSavRec.Fields("FrmDate").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    lbl(10).Caption = IIf(IsNull(RsSavRec.Fields("SalesTotal").value), 0, RsSavRec.Fields("SalesTotal").value)
    lbl(12).Caption = IIf(IsNull(RsSavRec.Fields("RetSalesTotal").value), 0, RsSavRec.Fields("RetSalesTotal").value)
    
    FrmDate2.value = IIf(IsNull(RsSavRec.Fields("FrmDate2").value), Date, RsSavRec.Fields("FrmDate2").value)
    ToDate2.value = IIf(IsNull(RsSavRec.Fields("ToDate2").value), Date, RsSavRec.Fields("ToDate2").value)
    lbl(16).Caption = IIf(IsNull(RsSavRec.Fields("Total").value), 0, RsSavRec.Fields("Total").value)
    lbl(19).Caption = IIf(IsNull(RsSavRec.Fields("Total1").value), 0, RsSavRec.Fields("Total1").value)
    TxtRemarks2.Text = IIf(IsNull(RsSavRec.Fields("Remarks2").value), "", RsSavRec.Fields("Remarks2").value)
    
     Me.TxtNoteID2.Text = IIf(IsNull(RsSavRec.Fields("NoteID2").value), "", RsSavRec.Fields("NoteID2").value)
     Me.TxtNoteSerial2.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial2").value), "", RsSavRec.Fields("NoteSerial2").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
ErrTrap:
End Sub
Function createVoucher2()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«»  «·„»Ì⁄«   ·‰ﬁ«ÿ «·»Ì⁄" & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String


tablename = "TblReCalVATPO"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0
notytype = 9092
Notevalue = val(lbl(16).Caption) + val(lbl(19).Caption)
BranchID = val(Dcbranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
        
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, , , "NoteID2", "NoteSerial2"   ', recordDateH.value"
                                              TxtNoteID2.Text = NoteID
                                                     TxtNoteSerial2.Text = NoteSerial

CREATE_VOUCHER_GE2 val(TxtNoteID2.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
     End If
End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·ﬁÌ„… «·„÷«›… ·‰ﬁ«ÿ «·»Ì⁄" & TxtSerial1.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblReCalVATPO"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0
notytype = 9087
Notevalue = val(lbl(10).Caption) + val(lbl(12).Caption)
BranchID = val(Dcbranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
        
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
     End If
End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» «·ﬁÌ„… «·„÷«›… ·‰ﬁ«ÿ «·»Ì⁄" & TxtSerial1.Text
    notes_id = general_noteid
    my_branch = val(Dcbranch.BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
With fg
For i = 1 To .Rows - 1
   Notevalue = val(.TextMatrix(i, .ColIndex("TotalVAT")))
   my_branch = val(.TextMatrix(i, .ColIndex("BranchID")))
            If Notevalue > 0 Then
                                    
                             StrAccountCodeDebt = get_account_code_branch(2, my_branch)
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
  
                           GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 21
                            line_no = line_no + 1
                                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ··„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
 Next i
End With
With VSFlexGrid1
For i = 1 To .Rows - 1
   Notevalue = val(.TextMatrix(i, .ColIndex("TotalVAT")))
   my_branch = val(.TextMatrix(i, .ColIndex("BranchID")))
            If Notevalue > 0 Then
                            GetValueAddedAccount XPDtbTrans.value, StrAccountCodeDebt, , 1, 9
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«» «·ﬁÌ„… «·„÷«›… ·„—œÊœ«  «·„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
  
                           StrAccountCodeCridet = get_account_code_branch(3, my_branch)
                            line_no = line_no + 1
                                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  „—œÊœ«  «·„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
            
 Next i
End With
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function
Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
    Dim BankID As Long
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«»  «·„»Ì⁄«  ·‰ﬁ«ÿ «·»Ì⁄" & TxtSerial1.Text
    notes_id = general_noteid
    my_branch = val(Dcbranch.BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    
    Dim VatTotals As Double
    VatTotals = val(lbl(10).Caption) - val(lbl(12).Caption)
    
With VSFlexGrid2
For i = 1 To .Rows - 1
   my_branch = val(.TextMatrix(i, .ColIndex("BranchID")))
   Notevalue = val(.TextMatrix(i, .ColIndex("Cash")))
   StrAccountCodeDebt = GetMyAccountCode("TblBoxesData", "BoxID", val(.TextMatrix(i, .ColIndex("BoxID"))))
               If Notevalue > 0 And StrAccountCodeDebt <> "" Then
                              
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·Œ“Ì‰Â  ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
            Notevalue = val(.TextMatrix(i, .ColIndex("Credit")))
            BankID = GetPaymentBank()
            StrAccountCodeDebt = GetMyAccountCode("BanksData", "BankiD", BankID)
                        If Notevalue > 0 And StrAccountCodeDebt <> "" Then
                       
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·»‰ﬂ  ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
            Notevalue = val(.TextMatrix(i, .ColIndex("Total")))
            If Notevalue > 0 Then
                             StrAccountCodeDebt = get_account_code_branch(2, my_branch)
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 1, Msg & "    Õ”«»  «·„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
            
            
 
            
            
            
 Next i
End With



With VSFlexGrid3
For i = 1 To .Rows - 1
   my_branch = val(.TextMatrix(i, .ColIndex("BranchID")))
   Notevalue = val(.TextMatrix(i, .ColIndex("Total")))
   StrAccountCodeCridet = GetMyAccountCode("TblBoxesData", "BoxID", val(.TextMatrix(i, .ColIndex("BoxID"))))
   
   
              If Notevalue > 0 Then
                             StrAccountCodeDebt = get_account_code_branch(3, my_branch)
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  „—œÊœ«  «·„»Ì⁄«   ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
            
            
               If Notevalue > 0 And StrAccountCodeDebt <> "" Then
                              
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  «·Œ“Ì‰Â  ", val(notes_id), , , , .TextMatrix(i, .ColIndex("RecordDate")), user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchID")))) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
         End If
   
        
 
            
 
            
            
            
 Next i
End With

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function
  Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " SELECT     dbo.TblReCalVATPO.ID, dbo.TblReCalVATPO.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblReCalVATPO.RecordDate, "
    MySQL = MySQL & "                  dbo.TblReCalVATPO.Remarks, dbo.TblReCalVATPO.FrmDate, dbo.TblReCalVATPO.ToDate, dbo.TblReCalVATPO.SalesTotal, dbo.TblReCalVATPO.RetSalesTotal,"
    MySQL = MySQL & "                  dbo.TblReCalVATPODet.RecordDate AS RecordDateDet, dbo.TblReCalVATPODet.BranchID AS BranchIDDet, TblBranchesData_1.branch_name AS branch_nameDet,"
    MySQL = MySQL & "                  TblBranchesData_1.branch_namee AS branch_nameeDet, dbo.TblReCalVATPODet.TotalVAT, dbo.TblReCalVATPODet.TypeTrans"
    MySQL = MySQL & "    FROM         dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblReCalVATPODet ON TblBranchesData_1.branch_id = dbo.TblReCalVATPODet.BranchID RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblReCalVATPO ON dbo.TblReCalVATPODet.RecalVATID = dbo.TblReCalVATPO.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData ON dbo.TblReCalVATPO.BranchID = dbo.TblBranchesData.branch_id"
    MySQL = MySQL & " Where (dbo.TblReCalVATPO.ID = " & val(TxtSerial1.Text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepReCalVATPO.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepReCalVATPO.rpt"
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

    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
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
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
If val(lbl(12).Caption) > 0 Then
            Account_Code_dynamic = get_account_code_branch(3, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    Exit Sub
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „—œÊœ«  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Exit Sub
                    End If
                End If
                
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 9) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ·„—œÊœ«  «·„»Ì⁄« "
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
If val(lbl(10).Caption) > 0 Then
            Account_Code_dynamic = get_account_code_branch(2, my_branch)
        If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical

                    Exit Sub
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Exit Sub
                    End If
                End If
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 21) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…«·„»Ì⁄« "
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblReCalVATPO", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    fg.Clear flexClearScrollable, flexClearEverything
    fg.Rows = 1
   sql = " SELECT     dbo.TblReCalVATPODet.ID, dbo.TblReCalVATPODet.RecalVATID, dbo.TblReCalVATPODet.RecordDate, dbo.TblReCalVATPODet.TotalVAT, "
   sql = sql & "                    dbo.TblReCalVATPODet.TypeTrans , dbo.TblReCalVATPODet.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
   sql = sql & "      FROM         dbo.TblReCalVATPODet  INNER JOIN "
   sql = sql & "                    dbo.TblBranchesData ON dbo.TblReCalVATPODet.BranchID = dbo.TblBranchesData.branch_id"
   sql = sql & "    Where (dbo.TblReCalVATPODet.TypeTrans = 0) And (dbo.TblReCalVATPODet.RecalVATID = " & val(TxtSerial1.Text) & ")"
   
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With fg
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("TotalVAT")) = IIf(IsNull(Rs1("TotalVAT").value), "", Rs1("TotalVAT").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
            End If
                   Rs1.MoveNext
             Next i
        End With
    ''///////////////
        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
   sql = " SELECT     dbo.TblReCalVATPODet.ID, dbo.TblReCalVATPODet.RecalVATID, dbo.TblReCalVATPODet.RecordDate, dbo.TblReCalVATPODet.TotalVAT, "
   sql = sql & "                    dbo.TblReCalVATPODet.TypeTrans , dbo.TblReCalVATPODet.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
   sql = sql & "      FROM         dbo.TblReCalVATPODet LEFT OUTER JOIN"
   sql = sql & "                    dbo.TblBranchesData ON dbo.TblReCalVATPODet.BranchID = dbo.TblBranchesData.branch_id"
   sql = sql & "    Where (dbo.TblReCalVATPODet.TypeTrans = 1) And (dbo.TblReCalVATPODet.RecalVATID = " & val(TxtSerial1.Text) & ")"
   
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid1
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("TotalVAT")) = IIf(IsNull(Rs1("TotalVAT").value), "", Rs1("TotalVAT").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
            End If
                   Rs1.MoveNext
             Next i
        End With
   ''///////////////
        VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 1
   'sql = "SELECT        dbo.TblReCalVATPODet.ID, dbo.TblReCalVATPODet.RecalVATID, dbo.TblReCalVATPODet.RecordDate, dbo.TblReCalVATPODet.TypeTrans, dbo.TblReCalVATPODet.BranchID, dbo.TblBranchesData.branch_name, "
   'sql = sql & "                       dbo.TblBranchesData.branch_namee, dbo.TblReCalVATPODet.Credit, dbo.TblReCalVATPODet.Cash, dbo.TblReCalVATPODet.Total, dbo.TblReCalVATPODet.BoxID, dbo.TblBoxesData.BoxName,"
   'sql = sql & "                        dbo.TblBoxesData.BoxNameE , dbo.TblReCalVATPODet.EmpID, dbo.cachierData.Name, dbo.cachierData.NameE"
   'sql = sql & "      FROM            dbo.TblReCalVATPODet INNER JOIN "
   'sql = sql & "                        dbo.cachierData ON dbo.TblReCalVATPODet.EmpID = dbo.cachierData.EmpID LEFT OUTER JOIN"
   'sql = sql & "                        dbo.TblBoxesData ON dbo.TblReCalVATPODet.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
   'sql = sql & "                        dbo.TblBranchesData ON dbo.TblReCalVATPODet.BranchID = dbo.TblBranchesData.branch_id"
   'sql = sql & "     Where (dbo.TblReCalVATPODet.TypeTrans = 2) And (dbo.TblReCalVATPODet.RecalVATID = " & val(TxtSerial1.Text) & ")"
   
 'moustafa
 sql = "SELECT     dbo.TblReCalVATPODet.ID, dbo.TblReCalVATPODet.RecalVATID, dbo.TblReCalVATPODet.RecordDate, dbo.TblReCalVATPODet.TypeTrans,"
 sql = sql & "       dbo.TblReCalVATPODet.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblReCalVATPODet.Credit,"
 sql = sql & "         dbo.TblReCalVATPODet.Cash, dbo.TblReCalVATPODet.Total, dbo.TblReCalVATPODet.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE,"
 sql = sql & "  dbo.TblReCalVATPODet.EmpID , dbo.cachierData.Name"
 sql = sql & " FROM         dbo.cachierData INNER JOIN"
 sql = sql & "                    dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID RIGHT OUTER JOIN"
 sql = sql & "       dbo.TblReCalVATPODet ON dbo.TblBoxesData.BoxID = dbo.TblReCalVATPODet.BoxID LEFT OUTER JOIN"
 sql = sql & "       dbo.TblBranchesData ON dbo.TblReCalVATPODet.BranchID = dbo.TblBranchesData.branch_id"
 sql = sql & "  Where (dbo.TblReCalVATPODet.TypeTrans = 2) And (dbo.TblReCalVATPODet.RecalVATID = " & val(TxtSerial1.Text) & ")"
   
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid2
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("Credit")) = IIf(IsNull(Rs1("Credit").value), "", Rs1("Credit").value)
                   .TextMatrix(i, .ColIndex("Cash")) = IIf(IsNull(Rs1("Cash").value), "", Rs1("Cash").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), "", Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(Rs1("BoxID").value), "", Rs1("BoxID").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
               '   .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxName").value), "", Rs1("BoxName").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Namee").value), "", Rs1("Namee").value)
               '    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxNameE").value), "", Rs1("BoxNameE").value)
            End If
                   Rs1.MoveNext
             Next i
        End With
        
      ''///////////////
     VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.Rows = 1
   sql = "SELECT        dbo.TblReCalVATPODet.ID, dbo.TblReCalVATPODet.RecalVATID, dbo.TblReCalVATPODet.RecordDate, dbo.TblReCalVATPODet.TypeTrans, dbo.TblReCalVATPODet.BranchID, dbo.TblBranchesData.branch_name, "
   sql = sql & "                       dbo.TblBranchesData.branch_namee, dbo.TblReCalVATPODet.Credit, dbo.TblReCalVATPODet.Cash, dbo.TblReCalVATPODet.Total, dbo.TblReCalVATPODet.BoxID, dbo.TblBoxesData.BoxName,"
   sql = sql & "                        dbo.TblBoxesData.BoxNameE , dbo.TblReCalVATPODet.EmpID, dbo.cachierData.Name, dbo.cachierData.NameE"
   sql = sql & "      FROM            dbo.TblReCalVATPODet LEFT OUTER JOIN"
   sql = sql & "                        dbo.cachierData ON dbo.TblReCalVATPODet.EmpID = dbo.cachierData.EmpID LEFT OUTER JOIN"
   sql = sql & "                        dbo.TblBoxesData ON dbo.TblReCalVATPODet.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
   sql = sql & "                        dbo.TblBranchesData ON dbo.TblReCalVATPODet.BranchID = dbo.TblBranchesData.branch_id"
   sql = sql & "     Where (dbo.TblReCalVATPODet.TypeTrans = 3) And (dbo.TblReCalVATPODet.RecalVATID = " & val(TxtSerial1.Text) & ")"
   
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid3
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("Credit")) = IIf(IsNull(Rs1("Credit").value), "", Rs1("Credit").value)
                   .TextMatrix(i, .ColIndex("Cash")) = IIf(IsNull(Rs1("Cash").value), "", Rs1("Cash").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), "", Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(Rs1("BoxID").value), "", Rs1("BoxID").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                 '  .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxName").value), "", Rs1("BoxName").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Namee").value), "", Rs1("Namee").value)
                  ' .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(Rs1("BoxNameE").value), "", Rs1("BoxNameE").value)
            End If
                   Rs1.MoveNext
             Next i
        End With
        
           lbl(23).Caption = lbl(16).Caption - lbl(10).Caption
    lbl(25).Caption = lbl(19).Caption - lbl(12).Caption
     

        Exit Sub
ErrTrap:
    End Sub
    Private Sub RemoveGridRow()
    With Me.fg
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
    Private Sub RemoveGridRow3()
    With Me.VSFlexGrid2
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid2
End Sub
    Private Sub RemoveGridRow4()
    With Me.VSFlexGrid3
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid2
End Sub
    Private Sub RemoveGridRow2()
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
FillGrid2
ReLineGrid2
End If
End Sub

Private Sub ISButton3_Click()
Dim sql As String
sql = "delete       dbo.Notes Where(NoteType = 64)"
 If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.Notes.NoteDate >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.Notes.NoteDate <=" & SQLDate(ToDate.value, True) & ""
End If

Cn.Execute sql
MsgBox " „ „—«Ã⁄Â «·ﬁÌÊœ"
End Sub

Private Sub ISButton4_Click()
'salim intialize
 Dim sql As String
 
sql = "update Transactions   set Transactions.Transaction_NetValue=isnull (Transactions.VAT,0)+ QryTransactionsTotal.TransNet"
sql = sql & " FROM         dbo.Transactions RIGHT OUTER JOIN"
sql = sql & "                       dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 9) or (dbo.Transactions.Transaction_Type = 22) or (dbo.Transactions.Transaction_Type = 21)or (dbo.Transactions.Transaction_Type = 5)  "
If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If

Cn.Execute sql
MsgBox " „ „—«Ã⁄Â «·›Ê« Ì—"
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
Sub UpdateFlg(Optional TypTrans As Integer)
Dim sql As String
If TypTrans = 0 Then
sql = "update transactions set FlgReCalVATPo=Null  "
Else
sql = "update transactions set FlgReCalVATPo=1  "
End If
sql = sql & "  Where (IsNull(dbo.transactions.POSBillType, 0) <> 0) And (dbo.transactions.Transaction_Type = 21)"
If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If
Cn.Execute sql
If TypTrans = 0 Then
sql = "update transactions set FlgReCalVATPo=Null  "
Else
sql = "update transactions set FlgReCalVATPo=1  "
End If
sql = sql & "  Where (IsNull(dbo.transactions.POSBillType, 0) <> 0) And (dbo.transactions.Transaction_Type = 9)"
If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If
Cn.Execute sql
End Sub
Sub FillGrid()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
  fg.Clear flexClearScrollable, flexClearEverything
  fg.Rows = 1
Set rs2 = New ADODB.Recordset
If SystemOptions.PriceWithVAT = True Then
sql = " SELECT     SUM(dbo.Transactions.Transaction_NetValue / 1.05 * 5 / 100 ) AS SumVAT"
Else
sql = " SELECT     SUM(dbo.Transactions.VAT) AS SumVAT"
End If
sql = sql & " , dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name , "
sql = sql & "                       dbo.TblBranchesData.branch_namee"
sql = sql & "  FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (dbo.transactions.POSBillType=1  or dbo.transactions.POSBillType=2  or dbo.transactions.POSBillType=3  or dbo.transactions.POSBillType=4 ) And (dbo.transactions.Transaction_Type = 21)"
 

'If Me.TxtModFlg.Text = "N" Then
'sql = sql & " AND (ISNULL(dbo.Transactions.FlgReCalVATPo, 0) = 0)"
'End If
'If Me.TxtModFlg.Text = "E" Then
'sql = sql & " AND ( (ISNULL(dbo.Transactions.FlgReCalVATPo, 0) = 0) or dbo.Transactions.FlgReCalVATPo=" & val(TxtSerial1.Text) & ")"
'End If


If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If
sql = sql & " GROUP BY dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
If SystemOptions.PriceWithVAT = True Then
sql = sql & " Having (SUM(dbo.Transactions.Transaction_NetValue / 1.05 * 5 / 100 ) > 0)"
Else
sql = sql & " Having (SUM(dbo.transactions.Vat) > 0)"
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With fg
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex(("Ser"))) = i
.TextMatrix(i, .ColIndex(("BranchID"))) = IIf(IsNull(rs2("BranchId").value), "", rs2("BranchId").value)
.TextMatrix(i, .ColIndex(("TotalVAT"))) = IIf(IsNull(rs2("SumVAT").value), 0, rs2("SumVAT").value)
.TextMatrix(i, .ColIndex(("RecordDate"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
Else
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
End If
rs2.MoveNext
Next i
End If
End With
''/////////////////////////////
  VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
  VSFlexGrid1.Rows = 1
Set rs2 = New ADODB.Recordset
If SystemOptions.PriceWithVAT = True Then
sql = " SELECT     SUM(dbo.Transactions.Transaction_NetValue / 1.05 * 5 / 100 ) AS SumVAT"
Else
sql = " SELECT     SUM(dbo.Transactions.VAT) AS SumVAT"
End If
sql = sql & "  , dbo.transactions.Transaction_Date , dbo.transactions.BranchID, dbo.TblBranchesData.branch_name, "
sql = sql & "                       dbo.TblBranchesData.branch_namee"
sql = sql & "  FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (IsNull(dbo.transactions.POSBillType, 0) <> 0) And (dbo.transactions.Transaction_Type = 9)"
If Me.TxtModFlg.Text = "N" Then
sql = sql & " AND (ISNULL(dbo.Transactions.FlgReCalVATPo, 0) = 0)"
End If
If Me.TxtModFlg.Text = "E" Then
sql = sql & " AND ( (ISNULL(dbo.Transactions.FlgReCalVATPo, 0) = 0) or dbo.Transactions.FlgReCalVATPo=" & val(TxtSerial1.Text) & ")"
End If
If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If
sql = sql & " GROUP BY dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
If SystemOptions.PriceWithVAT = True Then
sql = sql & " Having (SUM(dbo.transactions.Vat) > 0)"
Else
sql = sql & " Having (SUM(dbo.Transactions.Transaction_NetValue / 1.05 * 5 / 100 ) > 0)"
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With VSFlexGrid1
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex(("Ser"))) = i
.TextMatrix(i, .ColIndex(("BranchID"))) = IIf(IsNull(rs2("BranchId").value), "", rs2("BranchId").value)
.TextMatrix(i, .ColIndex(("TotalVAT"))) = IIf(IsNull(rs2("SumVAT").value), 0, rs2("SumVAT").value)
.TextMatrix(i, .ColIndex(("RecordDate"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
Else
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
End If
rs2.MoveNext
Next i
End If
End With
End Sub
Sub FillGrid2()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

  VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
  VSFlexGrid2.Rows = 1




'moustafa

Set rs2 = New ADODB.Recordset
sql = " SELECT     sum(dbo.Transactions.Transaction_NetValue) as totals ,   dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Emp_ID, "
sql = sql & "                         dbo.cachierData.Name , dbo.cachierData.NameE, dbo.cachierData.boxId, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE"
sql = sql & " FROM            dbo.Transactions INNER JOIN"
sql = sql & "                         dbo.cachierData INNER JOIN"
sql = sql & "                         dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID ON dbo.Transactions.Emp_ID = dbo.cachierData.EmpID LEFT OUTER JOIN"
sql = sql & "                         dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (IsNull(dbo.transactions.POSBillType, 0) <> 0) And (dbo.transactions.Transaction_Type = 21) and (isnull(dbo.Transactions.Emp_ID,0)<>0)"

If Not IsNull(FrmDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If

sql = sql & " GROUP BY dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Emp_ID,"
sql = sql & "                         dbo.cachierData.Name , dbo.cachierData.NameE, dbo.cachierData.boxId, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

With VSFlexGrid2
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex(("Ser"))) = i
.TextMatrix(i, .ColIndex(("BoxID"))) = IIf(IsNull(rs2("BoxID").value), 0, rs2("BoxID").value)
.TextMatrix(i, .ColIndex(("EmpID"))) = IIf(IsNull(rs2("Emp_ID").value), 0, rs2("Emp_ID").value)
.TextMatrix(i, .ColIndex(("BranchID"))) = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
.TextMatrix(i, .ColIndex(("RecordDate"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)
'.TextMatrix(i, .ColIndex(("Cash"))) = GetValue(0, 21, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
'.TextMatrix(i, .ColIndex(("Credit"))) = GetValue(1, 21, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
'.TextMatrix(i, .ColIndex(("Total"))) = val(.TextMatrix(i, .ColIndex(("Credit")))) + val(.TextMatrix(i, .ColIndex(("Cash"))))


'*****************BY AHMED SALIM*********************
.TextMatrix(i, .ColIndex(("Total"))) = IIf(IsNull(rs2("totals").value), 0, rs2("totals").value)
.TextMatrix(i, .ColIndex(("Credit"))) = GetValue(1, 21, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
.TextMatrix(i, .ColIndex(("Cash"))) = IIf(IsNull(rs2("totals").value), 0, rs2("totals").value) - val(.TextMatrix(i, .ColIndex(("Credit"))))
If 1 = 2 Then
.TextMatrix(i, .ColIndex(("Cash"))) = GetValue(0, 21, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
.TextMatrix(i, .ColIndex(("Total"))) = val(.TextMatrix(i, .ColIndex(("Cash")))) + val(.TextMatrix(i, .ColIndex(("Cash"))))
End If

If val(.TextMatrix(i, .ColIndex(("Cash")))) < 0 Then
.TextMatrix(i, .ColIndex(("Cash"))) = IIf(IsNull(rs2("totals").value), 0, rs2("totals").value)
.TextMatrix(i, .ColIndex(("Credit"))) = 0
End If
 '*****************BY AHMED SALIM*********************




If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex(("name"))) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
Else
.TextMatrix(i, .ColIndex(("name"))) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
End If

rs2.MoveNext
Next i
End If
End With
''/////////////////////////////
  VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
  VSFlexGrid3.Rows = 1
Set rs2 = New ADODB.Recordset
sql = " SELECT        dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Emp_ID, "
sql = sql & "                         dbo.cachierData.Name , dbo.cachierData.NameE, dbo.cachierData.boxId, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE"
sql = sql & " FROM            dbo.Transactions INNER JOIN"
sql = sql & "                         dbo.cachierData INNER JOIN"
sql = sql & "                         dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID ON dbo.Transactions.Emp_ID = dbo.cachierData.EmpID LEFT OUTER JOIN"
sql = sql & "                         dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  Where (IsNull(dbo.transactions.POSBillType, 0) <> 0) And (dbo.transactions.Transaction_Type = 9) and (isnull(dbo.Transactions.Emp_ID,0)<>0)  "
If Not IsNull(FrmDate2.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
End If
If Not IsNull(ToDate2.value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate.value, True) & ""
End If
sql = sql & " GROUP BY dbo.Transactions.Transaction_Date, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Emp_ID,"
sql = sql & "                         dbo.cachierData.Name , dbo.cachierData.NameE, dbo.cachierData.boxId, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With VSFlexGrid3
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex(("Ser"))) = i
.TextMatrix(i, .ColIndex(("BoxID"))) = IIf(IsNull(rs2("BoxID").value), 0, rs2("BoxID").value)
.TextMatrix(i, .ColIndex(("EmpID"))) = IIf(IsNull(rs2("Emp_ID").value), 0, rs2("Emp_ID").value)
.TextMatrix(i, .ColIndex(("BranchID"))) = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
.TextMatrix(i, .ColIndex(("RecordDate"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)
.TextMatrix(i, .ColIndex(("Cash"))) = GetValue(0, 9, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
.TextMatrix(i, .ColIndex(("Credit"))) = GetValue(1, 9, val(.TextMatrix(i, .ColIndex(("EmpID")))), .TextMatrix(i, .ColIndex(("RecordDate"))))
.TextMatrix(i, .ColIndex(("Total"))) = val(.TextMatrix(i, .ColIndex(("Credit")))) + val(.TextMatrix(i, .ColIndex(("Cash"))))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex(("name"))) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
Else
.TextMatrix(i, .ColIndex(("name"))) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
.TextMatrix(i, .ColIndex(("branch_name"))) = IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
End If
rs2.MoveNext
Next i
End If
End With
End Sub
Function GetValue(Optional PaymentID As Integer = 0, Optional Transaction_Type As Integer, Optional Emp_id As Double, Optional RecDate As Date) As Double
Dim sql As String
Dim My_SQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


sql = " SELECT        SUM(ISNULL(dbo.TblTransactionPayments.value, dbo.Transactions.Transaction_NetValue) ) AS Value"
'sql = " SELECT        SUM(ISNULL(dbo.TblTransactionPayments.value, dbo.Transactions.Transaction_NetValue)-isnull(dbo.Transactions.Vat,0)) AS Value"
sql = sql & " FROM            dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                         dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID"
sql = sql & " WHERE         (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transactions.Emp_ID = " & Emp_id & ") AND (dbo.Transactions.Transaction_Date = " & SQLDate(RecDate, True) & ") and (IsNull(dbo.transactions.POSBillType, 0) <> 0)"


GoTo salim

 'salimhere
   'new salim
 My_SQL = "SELECT     TOP 100 PERCENT SUM("
My_SQL = My_SQL & " ISNULL(dbo.TblTransactionPayments.[value] *"

My_SQL = My_SQL & "  isnull( dbo.TblTransactionPayments.Effect,"
My_SQL = My_SQL & "  case"

My_SQL = My_SQL & "  when Transaction_Type=21 then 1"
My_SQL = My_SQL & "  else  -1"
My_SQL = My_SQL & "  End"

My_SQL = My_SQL & "  )"

My_SQL = My_SQL & "  , dbo.Transactions.Transaction_NetValue)) AS TotalValue,"
My_SQL = My_SQL & "   dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.branch_no, dbo.TblPaymentType.MaxValue, dbo.TblPaymentType.TypTran, dbo.Transactions.Emp_ID,"
My_SQL = My_SQL & "                       dbo.transactions.POSBillType,dbo.TblPaymentType.Accountsus"
My_SQL = My_SQL & " FROM         dbo.TblTransactionPayments RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.Transactions ON dbo.TblTransactionPayments.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Date = CONVERT(DATETIME, '2019-11-10 00:00:00', 102))"
' My_SQL = My_SQL & " WHERE 1=1 "
' My_SQL = My_SQL & "  AND  (Transaction_Date >='" & SQLDate(FromDate) & "'"
' My_SQL = My_SQL & "  AND   Transaction_Date <='" & SQLDate(ToDate) & "')"
 
My_SQL = My_SQL & " WHERE         (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transactions.Emp_ID = " & Emp_id & ") AND (dbo.Transactions.Transaction_Date = " & SQLDate(RecDate, True) & ") and (IsNull(dbo.transactions.POSBillType, 0) <> 0)"

'My_SQL = My_SQL & " AND  ( isnull(dbo.TblTransactionPayments.PaymentID,0)=" & PaymentID & ")"

If PaymentID = 0 Then
My_SQL = My_SQL & " and (isnull(dbo.TblTransactionPayments.PaymentID,0) = 0  )"
Else
My_SQL = My_SQL & " and   (isnull(dbo.TblTransactionPayments.PaymentID,0) <> 0  )"
End If

'My_SQL = My_SQL & " and      (dbo.Transactions.POSBillType = 1 OR"
'My_SQL = My_SQL & "                       dbo.Transactions.POSBillType = 4) AND (dbo.Transactions.Transaction_Type = 21 OR"
'My_SQL = My_SQL & "                       dbo.Transactions.Transaction_Type = 9) AND (dbo.Transactions.Emp_ID = " & Emp_id & ")"
'If PaymentID = 0 Then
'My_SQL = My_SQL & " and (isnull(dbo.TblTransactionPayments.PaymentID,0) = 0  )"
'Else
'My_SQL = My_SQL & " and   (isnull(dbo.TblTransactionPayments.PaymentID,0) <> 0  )"
'End If
'
My_SQL = My_SQL & " GROUP BY isnull(dbo.TblTransactionPayments.PaymentID,0), dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.branch_no, dbo.TblPaymentType.MaxValue, dbo.TblPaymentType.TypTran, dbo.Transactions.Emp_ID,"
My_SQL = My_SQL & "                       dbo.transactions.POSBillType,dbo.TblPaymentType.Accountsus"

My_SQL = My_SQL & "  ORDER BY isnull(dbo.TblTransactionPayments.PaymentID,0)"


sql = My_SQL

salim:
Dim i As Integer
Dim sumAll As Double
sumAll = 0
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
For i = 1 To rs2.RecordCount
sumAll = sumAll + IIf(IsNull(rs2("TotalValue").value), 0, rs2("TotalValue").value)
rs2.MoveNext
Next i
GetValue = sumAll

Else
GetValue = 0
End If
End Function
Private Sub ShowBtn_Click()
If Me.TxtModFlg.Text <> "R" Then
FillGrid
ReLineGrid
FillGrid2
ReLineGrid2
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
    If TxtNoteSerial.Text <> "" Or TxtNoteSerial2.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
      Dim StrSQL As String
      UpdateFlg 0
      StrSQL = "delete from   TblReCalVATPODet where RecalVATID =" & val(TxtSerial1.Text) & ""
      Cn.Execute StrSQL
            RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
            RsSavRec.delete
            fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 1
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 1
            VSFlexGrid1.Rows = 1
            lbl(10).Caption = 0
            lbl(12).Caption = 0
            lbl(16).Caption = 0
            lbl(19).Caption = 0
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           'Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
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
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    XPDtbTrans.Enabled = True
        Command5.Enabled = False
        Command3.Enabled = False
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        
    ElseIf TxtModFlg.Text = "R" Then
    Command5.Enabled = True
    Command3.Enabled = True
     XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
   Command5.Enabled = False
   Command3.Enabled = False
    XPDtbTrans.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtNoteSerial.Text <> "" Or TxtNoteSerial2.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
       
        Me.DCboUserName.BoundText = user_id
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
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
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    TxtModFlg.Text = "N"
    fg.Clear flexClearScrollable, flexClearEverything
    fg.Rows = 2
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.Rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    XPDtbTrans.value = Date
    lbl(10).Caption = 0
    lbl(12).Caption = 0
    lbl(16).Caption = 0
    lbl(19).Caption = 0
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub


'Information for camand
'++++++++++++++++++++++++++++++++++++++
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
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub ChangeLang()
On Error GoTo ErrTrap
       Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
   ''''''''''''''''''''////
       Me.Caption = "Recalculate The Cost"
      Label1(2).Caption = Me.Caption
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(21).Caption = "Remarks"
      Cmd(0).Caption = "Delete"
      Cmd(1).Caption = "Delete All"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
   


ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblReCalVATPO"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end

