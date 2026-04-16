VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmJobSeekersSelect 
   Caption         =   "«·»ÕÀ ›Ï „·›«  «·—«€»Ì‰ ›Ï «·⁄„·"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   Icon            =   "FrmJobSeekersSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   10590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8460
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10590
      _cx             =   18680
      _cy             =   14923
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
      Version         =   800
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmJobSeekersSelect.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   10
         Left            =   5295
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   7980
         Width           =   5265
         _cx             =   9287
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
         Appearance      =   5
         MousePointer    =   0
         Version         =   800
         BackColor       =   14871017
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
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√ŸÂ— «· ·„ÌÕ"
            Height          =   300
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   10
            Left            =   1290
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   120
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰ «∆Ã «·»ÕÀ:-"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   9
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   120
            Width           =   1155
         End
      End
      Begin ImpulseButton.ISButton CmdExit 
         Height          =   450
         Left            =   30
         TabIndex        =   36
         Top             =   7980
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   794
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
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7935
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10530
         _cx             =   18574
         _cy             =   13996
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
         Version         =   800
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "⁄Ê«„· «·»ÕÀ|‰ «∆Ã «·»ÕÀ"
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
         Picture(0)      =   "FrmJobSeekersSelect.frx":040F
         Picture(1)      =   "FrmJobSeekersSelect.frx":07A9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7530
            Index           =   1
            Left            =   11145
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   15
            Width           =   10500
            _cx             =   18521
            _cy             =   13282
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
            Version         =   800
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
            Begin VSFlex8UCtl.VSFlexGrid FgRes 
               Height          =   7410
               Left            =   30
               TabIndex        =   19
               Top             =   30
               Width           =   10470
               _cx             =   18468
               _cy             =   13070
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
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmJobSeekersSelect.frx":0B43
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7530
            Index           =   0
            Left            =   15
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   10500
            _cx             =   18521
            _cy             =   13282
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
            Version         =   800
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
               Height          =   1485
               Index           =   7
               Left            =   90
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   3900
               Width           =   5040
               _cx             =   8890
               _cy             =   2619
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   1
               MousePointer    =   0
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "«· ⁄·Ì„"
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
               Style           =   1
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
               Begin MSDataListLib.DataCombo DcboEducation 
                  Height          =   315
                  Left            =   630
                  TabIndex        =   50
                  Top             =   330
                  Width           =   3045
                  _ExtentX        =   5371
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ƒÂ· «· ⁄·Ì„Ï"
                  Height          =   285
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   780
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ƒÂ· «· ⁄·Ì„Ï"
                  Height          =   285
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   360
                  Width           =   1125
               End
            End
            Begin ImpulseButton.ISButton CmdSearch 
               Height          =   420
               Left            =   60
               TabIndex        =   34
               Top             =   7020
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   741
               ButtonPositionImage=   1
               Caption         =   "»œ¡ ⁄·„Ì… «·»ÕÀ"
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
               ButtonImage     =   "FrmJobSeekersSelect.frx":0DDD
               ColorButton     =   14871017
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4095
               Index           =   3
               Left            =   5160
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   3390
               Width           =   5280
               _cx             =   9313
               _cy             =   7223
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   4
               MousePointer    =   0
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "«·„Â‰… Ê«·Œ»—«  «·⁄„·Ì…"
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
               Style           =   1
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
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   3
                  Left            =   2580
                  MaxLength       =   2
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   2
                  Left            =   3480
                  MaxLength       =   2
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VSFlex8UCtl.VSFlexGrid FgJobs 
                  Height          =   2085
                  Left            =   180
                  TabIndex        =   37
                  Top             =   1530
                  Visible         =   0   'False
                  Width           =   4995
                  _cx             =   8811
                  _cy             =   3678
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
                  Rows            =   1
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmJobSeekersSelect.frx":1177
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
               Begin MSDataListLib.DataCombo DcboJobCat 
                  Height          =   315
                  Left            =   600
                  TabIndex        =   32
                  Top             =   375
                  Width           =   4005
                  _ExtentX        =   7064
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboJob 
                  Height          =   315
                  Left            =   600
                  TabIndex        =   33
                  Top             =   720
                  Width           =   4005
                  _ExtentX        =   7064
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdToggle 
                  Height          =   405
                  Index           =   2
                  Left            =   1020
                  TabIndex        =   40
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   795
                  _ExtentX        =   1402
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmJobSeekersSelect.frx":1203
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdToggle 
                  Height          =   405
                  Index           =   3
                  Left            =   240
                  TabIndex        =   41
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmJobSeekersSelect.frx":159D
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcboJobFav 
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   53
                  Top             =   3690
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√›÷·Ì… „ﬂ«‰ «·⁄„·"
                  Height          =   345
                  Index           =   60
                  Left            =   3690
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   3720
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄«„"
                  Height          =   270
                  Index           =   8
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   1170
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   210
                  Index           =   7
                  Left            =   3210
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   1170
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   240
                  Index           =   6
                  Left            =   4020
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·Œ»—…"
                  Height          =   315
                  Index           =   5
                  Left            =   4410
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   1170
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· Œ’’"
                  Height          =   345
                  Index           =   4
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   720
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ã«·"
                  Height          =   315
                  Index           =   3
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   390
                  Width           =   645
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3345
               Index           =   2
               Left            =   5160
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   30
               Width           =   5280
               _cx             =   9313
               _cy             =   5900
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   4
               MousePointer    =   0
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "«·»Ì«‰«  «·‘Œ’Ì…"
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
                  Height          =   405
                  Index           =   9
                  Left            =   660
                  TabIndex        =   57
                  TabStop         =   0   'False
                  Top             =   2880
                  Width           =   2850
                  _cx             =   5027
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
                  Appearance      =   5
                  MousePointer    =   0
                  Version         =   800
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
                  Begin VB.OptionButton OptHind 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "€Ì— „ﬁ»Ê·…"
                     Height          =   360
                     Index           =   0
                     Left            =   1620
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   30
                     Value           =   -1  'True
                     Width           =   1050
                  End
                  Begin VB.OptionButton OptHind 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„ﬁ»Ê·…"
                     Height          =   240
                     Index           =   1
                     Left            =   330
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   90
                     Width           =   1050
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   405
                  Index           =   8
                  Left            =   660
                  TabIndex        =   54
                  TabStop         =   0   'False
                  Top             =   2460
                  Width           =   2850
                  _cx             =   5027
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
                  Appearance      =   5
                  MousePointer    =   0
                  Version         =   800
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
                  Begin VB.OptionButton OptPass 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·«»œ „‰ ÊÃÊœÂ"
                     Height          =   315
                     Index           =   0
                     Left            =   1380
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   30
                     Value           =   -1  'True
                     Width           =   1290
                  End
                  Begin VB.OptionButton OptPass 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·«ÌÂ„ ÊÃÊœÂ"
                     Height          =   315
                     Index           =   1
                     Left            =   60
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   30
                     Width           =   1290
                  End
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ»Ê· Õ«·«  «·√⁄«ﬁ…"
                  Height          =   450
                  Index           =   7
                  Left            =   3900
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   2835
                  Width           =   1260
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÃÊ«“ «·”›—"
                  Height          =   285
                  Index           =   6
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   2535
                  Width           =   1620
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Êﬁ› „‰ «· Ã‰Ìœ"
                  Height          =   285
                  Index           =   5
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   2145
                  Value           =   1  'Checked
                  Width           =   1620
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ«·… «·√Ã „«⁄Ì…"
                  Height          =   285
                  Index           =   4
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   1800
                  Width           =   1440
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·œÌ«‰…"
                  Height          =   285
                  Index           =   3
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   1410
                  Width           =   960
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ã‰”Ì…"
                  Height          =   300
                  Index           =   2
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   1065
                  Width           =   960
               End
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   1
                  Left            =   1680
                  MaxLength       =   2
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   645
                  Width           =   570
               End
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   2640
                  MaxLength       =   2
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   645
                  Width           =   570
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„—"
                  Height          =   210
                  Index           =   1
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   675
                  Width           =   960
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê⁄"
                  Height          =   300
                  Index           =   0
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   270
                  Width           =   960
               End
               Begin MSDataListLib.DataCombo DcboSexType 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   5
                  Top             =   270
                  Width           =   2580
                  _ExtentX        =   4551
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboNationality 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   14
                  Top             =   1065
                  Width           =   2580
                  _ExtentX        =   4551
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboRegs 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   16
                  Top             =   1410
                  Width           =   2580
                  _ExtentX        =   4551
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboSocStates 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   17
                  Top             =   1770
                  Width           =   2580
                  _ExtentX        =   4551
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DboMStates 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   20
                  Top             =   2115
                  Width           =   2580
                  _ExtentX        =   4551
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄«„"
                  Height          =   270
                  Index           =   2
                  Left            =   1380
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   705
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   210
                  Index           =   1
                  Left            =   2340
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   705
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   240
                  Index           =   0
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   675
                  Width           =   270
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1965
               Index           =   4
               Left            =   60
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   30
               Width           =   5070
               _cx             =   8943
               _cy             =   3466
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   4
               MousePointer    =   0
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "«··€« "
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
               Begin VSFlex8UCtl.VSFlexGrid FgLang 
                  Height          =   1545
                  Left            =   90
                  TabIndex        =   27
                  Top             =   300
                  Width           =   4830
                  _cx             =   8520
                  _cy             =   2725
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
                  Rows            =   1
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmJobSeekersSelect.frx":1937
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1440
               Index           =   5
               Left            =   60
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   5370
               Width           =   5070
               _cx             =   8943
               _cy             =   2540
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
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "„ ‰Ê⁄"
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
               Style           =   1
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
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ’— «·»ÕÀ ›Ï «·⁄«ÿ·Ì‰ ⁄‰ «·⁄„·( «·Ì‰ ·« ÊÃœ ·Â„ ÊŸ«∆› Õ«·Ì«)"
                  Height          =   390
                  Index           =   9
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   825
                  Width           =   4815
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì›÷· «·€Ì— „— »ÿÌ‰ »ÊŸ«∆› ›Ï «·Êﬁ  «·Õ«·Ï"
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   510
                  Width           =   4815
               End
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì›÷· «·Õ«’·Ì‰ ⁄·Ï œÊ—« "
                  Height          =   210
                  Index           =   11
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   240
                  Width           =   2235
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1860
               Index           =   6
               Left            =   90
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   2010
               Width           =   5040
               _cx             =   8890
               _cy             =   3281
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   4
               MousePointer    =   0
               Version         =   800
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "„Â«—«  «·Õ«”» «·√·Ï"
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
                  Height          =   1335
                  Left            =   90
                  TabIndex        =   31
                  Top             =   345
                  Width           =   4830
                  _cx             =   8520
                  _cy             =   2355
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
                  Rows            =   1
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmJobSeekersSelect.frx":1A02
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
            End
         End
      End
   End
End
Attribute VB_Name = "FrmJobSeekersSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private FrmComment As FrmFlexQtyTip
Dim m_QtyToolTipShown As Boolean
Dim cDcboSearch(27) As clsDCboSearch
Private Sub Chk_Click(Index As Integer)
Select Case Index
    Case 0
        Me.DcboSexType.Enabled = CBool(Me.Chk(Index).Value)
    Case 1
       Me.lbl(0).Enabled = CBool(Me.Chk(Index).Value)
       Me.lbl(1).Enabled = CBool(Me.Chk(Index).Value)
       Me.lbl(2).Enabled = CBool(Me.Chk(Index).Value)
       Me.Txt(0).Enabled = CBool(Me.Chk(Index).Value)
       Me.Txt(1).Enabled = CBool(Me.Chk(Index).Value)
    Case 2
        Me.DcboNationality.Enabled = CBool(Me.Chk(Index).Value)
    Case 3
        Me.DcboRegs.Enabled = CBool(Me.Chk(Index).Value)
    Case 4
        Me.DcboSocStates.Enabled = CBool(Me.Chk(Index).Value)
    Case 5
        Me.DboMStates.Enabled = CBool(Me.Chk(Index).Value)
    Case 6
        Me.OptPass(0).Enabled = CBool(Me.Chk(Index).Value)
        Me.OptPass(1).Enabled = CBool(Me.Chk(Index).Value)
    Case 7
        Me.OptHind(0).Enabled = CBool(Me.Chk(Index).Value)
        Me.OptHind(1).Enabled = CBool(Me.Chk(Index).Value)
End Select
End Sub
Private Sub ChkShow_Click()
If ChkShow.Value = vbUnchecked Then
   HideComment
End If
End Sub

Private Sub CmdSearch_Click()
Dim StrSQL As String
Dim StrWhere As String
Dim Rs As ADODB.Recordset
Dim BolBegin As Boolean
StrSQL = "Select * From QrySeekersSearch"
StrWhere = ""
BolBegin = False
If Chk(0).Value = vbChecked Then
    If Me.DcboSexType.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.SexID=" & Me.DcboSexType.BoundText & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.SexID=" & Me.DcboSexType.BoundText & ""
            BolBegin = True
        End If
    End If
End If
If Chk(1).Value = vbChecked Then
    If Val(Me.Txt(0).Text) > 0 Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.SeekerAge >=" & Val(Me.Txt(0).Text) & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.SeekerAge >=" & Val(Me.Txt(0).Text) & ""
            BolBegin = True
        End If
    End If
    If Val(Me.Txt(1).Text) > 0 Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.SeekerAge <=" & Val(Me.Txt(1).Text) & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.SeekerAge <=" & Val(Me.Txt(1).Text) & ""
            BolBegin = True
        End If
    End If
End If
If Chk(2).Value = vbChecked Then
    If Me.DcboNationality.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.NtionalityID=" & Me.DcboNationality.BoundText & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.NtionalityID=" & Me.DcboNationality.BoundText & ""
            BolBegin = True
        End If
    End If
End If
If Chk(3).Value = vbChecked Then
    If Me.DcboRegs.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.RegID=" & Me.DcboRegs.BoundText & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.RegID=" & Me.DcboRegs.BoundText & ""
            BolBegin = True
        End If
    End If
End If
If Chk(4).Value = vbChecked Then
    If Me.DcboSocStates.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.SocID=" & Me.DcboSocStates.BoundText & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.SocID=" & Me.DcboSocStates.BoundText & ""
            BolBegin = True
        End If
    End If
End If
If Chk(5).Value = vbChecked Then
    If Me.DboMStates.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.MStateID=" & Me.DboMStates.BoundText & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.MStateID=" & Me.DboMStates.BoundText & ""
            BolBegin = True
        End If
    End If
End If
If Chk(6).Value = vbChecked Then
    If Me.OptPass(0).Value = True Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.IsPassport=" & 1 & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.IsPassport=" & 1 & ""
            BolBegin = True
        End If
    End If
End If
If Chk(7).Value = vbChecked Then
    If Me.OptHind(0).Value = True Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and QrySeekersSearch.IsHindrance=" & 0 & ""
        Else
            StrWhere = StrWhere + " Where QrySeekersSearch.IsHindrance=" & 0 & ""
            BolBegin = True
        End If
    End If
End If
If Me.DcboJobCat.BoundText <> "" Then
    If BolBegin = True Then
        StrWhere = StrWhere + " and QrySeekersSearch.JobCatID=" & Me.DcboJobCat.BoundText & ""
    Else
        StrWhere = StrWhere + " Where QrySeekersSearch.JobCatID=" & Me.DcboJobCat.BoundText & ""
        BolBegin = True
    End If
End If

If Me.DcboJob.BoundText <> "" Then
    If BolBegin = True Then
        StrWhere = StrWhere + " and QrySeekersSearch.JobFieldID=" & Me.DcboJob.BoundText & ""
    Else
        StrWhere = StrWhere + " Where QrySeekersSearch.JobFieldID=" & Me.DcboJob.BoundText & ""
        BolBegin = True
    End If
End If

Set Rs = New ADODB.Recordset
StrSQL = StrSQL + StrWhere
Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
With Me.FgRes
    .Clear flexClearScrollable, flexClearEverything
    .Rows = .FixedRows
    If Not (Rs.BOF Or Rs.EOF) Then
        Me.lbl(10).Caption = Rs.RecordCount
        .Rows = .FixedRows + Rs.RecordCount
        For I = 1 To .Rows - 1
            .TextMatrix(I, .ColIndex("JobSeekerID")) = IIf(IsNull(Rs("JobSeekerID").Value), "", Rs("JobSeekerID").Value)
            .TextMatrix(I, .ColIndex("CreationDate")) = IIf(IsNull(Rs("CreationDate").Value), "", Rs("CreationDate").Value)
            .TextMatrix(I, .ColIndex("JobSeekerCode")) = IIf(IsNull(Rs("JobSeekerCode").Value), "", Rs("JobSeekerCode").Value)
            .TextMatrix(I, .ColIndex("JobSeekerName")) = IIf(IsNull(Rs("JobSeekerName").Value), "", Rs("JobSeekerName").Value)
            .TextMatrix(I, .ColIndex("SexType")) = IIf(IsNull(Rs("SexType").Value), "", Rs("SexType").Value)
            .TextMatrix(I, .ColIndex("BrithDate")) = IIf(IsNull(Rs("BrithDate").Value), "", Rs("BrithDate").Value)
            .TextMatrix(I, .ColIndex("SeekerAge")) = IIf(IsNull(Rs("SeekerAge").Value), "", Rs("SeekerAge").Value)
            .TextMatrix(I, .ColIndex("CountryName")) = IIf(IsNull(Rs("CountryName").Value), "", Rs("CountryName").Value)
            .TextMatrix(I, .ColIndex("RegType")) = IIf(IsNull(Rs("RegType").Value), "", Rs("RegType").Value)
            .TextMatrix(I, .ColIndex("SocType")) = IIf(IsNull(Rs("SocType").Value), "", Rs("SocType").Value)
            '
            .TextMatrix(I, .ColIndex("ChilNo")) = IIf(IsNull(Rs("ChilNo").Value), "", Rs("ChilNo").Value)
            .TextMatrix(I, .ColIndex("MStateType")) = IIf(IsNull(Rs("MStateType").Value), "", Rs("MStateType").Value)
            .TextMatrix(I, .ColIndex("IsPassport")) = IIf(IsNull(Rs("IsPassport").Value), "", Rs("IsPassport").Value)
            .TextMatrix(I, .ColIndex("IsHindrance")) = IIf(IsNull(Rs("IsHindrance").Value), "", Rs("IsHindrance").Value)
            .TextMatrix(I, .ColIndex("HindranceName")) = IIf(IsNull(Rs("HindranceName").Value), "", Rs("HindranceName").Value)
            Rs.MoveNext
        Next I
    Else
        Me.lbl(10).Caption = 0
    End If
    .AutoSize 0, .Cols - 1, False
End With
End Sub

Private Sub DcboJobCat_Change()
LoadJobFields Me.DcboJobCat, Me.DcboJob, cDcboSearch(6)
End Sub

Private Sub DcboJobCat_Click(Area As Integer)
LoadJobFields Me.DcboJobCat, Me.DcboJob, cDcboSearch(6)
End Sub

Private Sub FgRes_Click()
Dim LngJobSeekerID As Long
Static LngOldCol As Long
Static LngOldRow As Long

If FgRes.Col = -1 Then
    'HideComment
    Exit Sub
End If
If FgRes.Row = -1 Then
    'HideComment
    Exit Sub
End If
If (LngOldRow = FgRes.Row And m_QtyToolTipShown = True) Then
    Exit Sub
Else
    'HideComment
    LngOldCol = FgRes.Col
    LngOldRow = FgRes.Row
End If



With Me.FgRes
    LngJobSeekerID = Val(.TextMatrix(.Row, .ColIndex("JobSeekerID")))
    If LngJobSeekerID <> 0 Then
        If Me.ChkShow.Value = vbChecked Then
            ShowSeekerTip LngJobSeekerID, .Row, .Col, True, True
        End If
    End If
End With
End Sub

Private Sub FgRes_DblClick()
With Me.FgRes
    If .Row = -1 Then Exit Sub
    If .Col = -1 Then Exit Sub
    If Val(.TextMatrix(.Row, .ColIndex("JobSeekerID"))) = 0 Then
        Exit Sub
    End If
    If .Col = .ColIndex("JobSeekerName") Then
        Load FrmJobSeekersData
        FrmJobSeekersData.Retrive Val(.TextMatrix(.Row, .ColIndex("JobSeekerID")))
        FrmJobSeekersData.Show
    End If
End With
End Sub

Private Sub FgRes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim LngJobSeekerID As Long
'Static LngMouseRow As Long
'Static LngMouseCol As Long
'
'If LngMouseRow = FgRes.MouseRow Then Exit Sub
'If LngMouseCol = FgRes.MouseCol Then Exit Sub
'
'If LngMouseCol = -1 Then
'    HideComment
'    Exit Sub
'End If
'If LngMouseRow = -1 Then
'    HideComment
'    Exit Sub
'End If
'
'With Me.FgRes
'    LngJobSeekerID = Val(.TextMatrix(.Row, .ColIndex("JobSeekerID")))
'    If LngJobSeekerID <> 0 Then
'        If Me.ChkShow.Value = vbChecked Then
'            ShowSeekerTip LngJobSeekerID, LngMouseRow, LngMouseRow, True, True
'        End If
'    End If
'End With
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim Dcombos As ClsDataCombos
Dim GrdBack As ClsBackGroundPic

Set Dcombos = New ClsDataCombos
Dcombos.GetSexTypes Me.DcboSexType
Set cDcboSearch(0) = New clsDCboSearch
Set cDcboSearch(0).Client = Me.DcboSexType

Dcombos.GetNationality Me.DcboNationality
Set cDcboSearch(1) = New clsDCboSearch
Set cDcboSearch(1).Client = Me.DcboNationality

Dcombos.GetRegs Me.DcboRegs
Set cDcboSearch(2) = New clsDCboSearch
Set cDcboSearch(2).Client = Me.DcboRegs

Dcombos.GetMiltStates Me.DboMStates
Set cDcboSearch(3) = New clsDCboSearch
Set cDcboSearch(3).Client = Me.DboMStates

Dcombos.GetSocStates Me.DcboSocStates
Set cDcboSearch(4) = New clsDCboSearch
Set cDcboSearch(4).Client = Me.DcboSocStates

Dcombos.GetJobCats Me.DcboJobCat
Set cDcboSearch(5) = New clsDCboSearch
Set cDcboSearch(5).Client = Me.DcboJobCat

Dcombos.GetJobFields Me.DcboJob
Set cDcboSearch(6) = New clsDCboSearch
Set cDcboSearch(6).Client = Me.DcboJob

For I = 0 To 7
    Me.Chk(I).Value = vbUnchecked
    Chk_Click I
Next I
For I = 0 To Me.FgRes.Cols - 1
    Me.FgRes.ColAlignment(I) = flexAlignRightCenter
    Me.FgRes.FixedAlignment(I) = flexAlignRightCenter
Next I
With Me.FgRes
    .AutoSize 0, .Cols - 1, False
End With
Set GrdBack = New ClsBackGroundPic
Set Me.FgRes.WallPaper = GrdBack.Picture
Set FgJobs.WallPaper = GrdBack.Picture

Me.Width = 10710
Me.Height = 8970
Me.ChkShow.Value = Val(GetSetting(SystemOptions.SysRegsAppPath & "\" & "UsersSetting", "User_Name", "ShowTip", Me.ChkShow.Value))
Resize_Form Me

End Sub

Private Sub ShowSeekerTip(LngSeekerID As Long, _
    LngShowRow As Long, LngShowCol As Long, _
    Optional bSetFocus As Boolean = False, _
    Optional BolShowInGrid As Boolean = True, _
    Optional Lnghwnd As Long)
    
'------------------------------

'------------------------------

Dim uPoint As POINTAPI
Static lNoteRow As Long, lNoteCol As Long, r As Long, C As Long
Dim lLeft  As Single
Dim LTop As Single
Dim StrSQL As String
Dim Rs As ADODB.Recordset
Dim Acmd As ADODB.Command
Dim XParItemID As ADODB.Parameter
Dim XParTransType As ADODB.Parameter
Dim I As Integer
Dim LngItemID As Long

Dim xRect As RECT

'HideComment

With Me.FgRes
    If BolShowInGrid = True Then
        If LngShowRow <= -1 Then Exit Sub
        If LngShowCol <= -1 Then Exit Sub
        If lNoteRow = LngShowRow And lNoteCol = LngShowCol Then
            'If bSetFocus = False Then Exit Sub
        End If
        ClientToScreen FgRes.hwnd, uPoint
        lLeft = (uPoint.X * Screen.TwipsPerPixelX) + .ColPos(LngShowCol) + .ColWidth(LngShowCol)
        LTop = (uPoint.Y * Screen.TwipsPerPixelY) + .RowPos(LngShowRow) + .RowHeight(LngShowRow) + 50
        'lLeft = lLeft + 100
    Else
        ClientToScreen Lnghwnd, uPoint
        GetWindowRect Lnghwnd, xRect
        lLeft = (uPoint.X * Screen.TwipsPerPixelX) ' + xRect.right '+ .ColPos(LngShowCol) + .ColWidth(LngShowCol)
        LTop = (uPoint.Y * Screen.TwipsPerPixelY) + xRect.bottom   '+ .RowPos(LngShowRow) + .RowHeight(LngShowRow) + 50
    End If
    If m_QtyToolTipShown = False Then
        Set FrmComment = New FrmFlexQtyTip
    End If
    With FrmComment
        .LoadSeekerData LngSeekerID
        If lLeft + .Width > Screen.Width Then
            lLeft = lLeft - .Width
        End If
        If LTop + .Height > Screen.Height Then
            LTop = LTop - .Height
        End If
        .left = lLeft
        .top = LTop
        If bSetFocus = False Then
            ShowWindow .hwnd, SW_SHOWNA
            SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

        Else
            .Show
            .SetFocus
            .TabMain.SetFocus
            SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

        End If
        '.ZOrder 0
        m_QtyToolTipShown = True
    End With
    m_QtyToolTipShown = True
    lNoteRow = LngShowRow
    lNoteCol = LngShowCol
End With
End Sub

Private Sub Form_Resize()
HideComment
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Integer
HideComment
SaveSetting SystemOptions.SysRegsAppPath & "\" & "UsersSetting", "User_Name", "ShowTip", Me.ChkShow.Value

For I = LBound(cDcboSearch) To UBound(cDcboSearch)
    Set cDcboSearch(I) = Nothing
Next I
End Sub

Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Txt(Index).Text, 1)
End Sub
Private Sub HideComment()
If Not FrmComment Is Nothing Then
    Unload FrmComment
    Set FrmComment = Nothing
    m_QtyToolTipShown = False
End If
End Sub
