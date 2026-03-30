VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form frmequipment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «·„⁄œ«  / «·√·« "
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   Icon            =   "frmequipment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10095
      _cx             =   17806
      _cy             =   14314
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6855
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   9975
         _cx             =   17595
         _cy             =   12091
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
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…|«·”⁄…|«·„Ê«’›« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6480
            Left            =   10620
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   45
            Width           =   9885
            _cx             =   17436
            _cy             =   11430
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   2415
               Left            =   -120
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   0
               Width           =   10005
               _cx             =   17648
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
               Begin VB.TextBox TxtItemCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7035
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   2040
                  Width           =   990
               End
               Begin VB.TextBox TxtCapacity 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1455
                  TabIndex        =   78
                  TabStop         =   0   'False
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   750
               End
               Begin VB.ListBox ListGroupSelected 
                  Height          =   1620
                  ItemData        =   "frmequipment.frx":058A
                  Left            =   165
                  List            =   "frmequipment.frx":0591
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   360
                  Width           =   4290
               End
               Begin VB.ListBox ListGroupAll 
                  Height          =   1620
                  ItemData        =   "frmequipment.frx":05A8
                  Left            =   5325
                  List            =   "frmequipment.frx":05AF
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   360
                  Width           =   4530
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·«’‰«›"
                  ForeColor       =   &H00FF0000&
                  Height          =   210
                  Index           =   0
                  Left            =   8340
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1365
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
                  ForeColor       =   &H000000FF&
                  Height          =   210
                  Index           =   2
                  Left            =   7920
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   2040
                  Width           =   2025
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ã„Ê⁄«  „Õœœ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   210
                  Index           =   1
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   120
                  Width           =   3765
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   390
                  Left            =   240
                  TabIndex        =   73
                  Top             =   1920
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«÷«›…"
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
                  ButtonImage     =   "frmequipment.frx":05C1
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcbItem 
                  Bindings        =   "frmequipment.frx":095B
                  Height          =   315
                  Left            =   1110
                  TabIndex        =   81
                  Top             =   2040
                  Width           =   5820
                  _ExtentX        =   10266
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄…"
                  Height          =   285
                  Index           =   1
                  Left            =   2190
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   945
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
                  Height          =   255
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   600
                  Width           =   660
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
                  Height          =   375
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   840
                  Width           =   660
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
                  Height          =   375
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1560
                  Width           =   660
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
                  Height          =   375
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1200
                  Width           =   660
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid FgItem 
               Height          =   3615
               Left            =   0
               TabIndex        =   82
               Top             =   2400
               Width           =   9825
               _cx             =   17330
               _cy             =   6376
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmequipment.frx":0970
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   3
               Left            =   8655
               TabIndex        =   83
               Top             =   6000
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ— "
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
               ButtonImage     =   "frmequipment.frx":0A79
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   4
               Left            =   7200
               TabIndex        =   84
               Top             =   6000
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
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
               ButtonImage     =   "frmequipment.frx":1013
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6480
            Left            =   45
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   45
            Width           =   9885
            _cx             =   17436
            _cy             =   11430
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
            Begin VB.Frame Frm2 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   3165
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   3360
               Width           =   9825
               Begin VB.TextBox TxtCapacites 
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
                  Left            =   360
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   2760
                  Width           =   1080
               End
               Begin VB.CheckBox ChKLockeq 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ìﬁ«› „⁄œ…"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   120
                  Width           =   2415
               End
               Begin VB.TextBox TxtEqupNameE 
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
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   840
                  Width           =   3975
               End
               Begin VB.TextBox TxtEqupName 
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
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   480
                  Width           =   3975
               End
               Begin VB.TextBox Txthelper 
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
                  TabIndex        =   45
                  Top             =   2760
                  Width           =   1560
               End
               Begin VB.TextBox TxtEmployer 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2760
                  Width           =   1920
               End
               Begin VB.TextBox TxtStopPercentage 
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
                  TabIndex        =   43
                  Top             =   1920
                  Width           =   1440
               End
               Begin VB.TextBox TxtStopvalue 
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
                  TabIndex        =   42
                  Top             =   1920
                  Width           =   1560
               End
               Begin VB.ComboBox CboInterval 
                  Height          =   315
                  ItemData        =   "frmequipment.frx":15AD
                  Left            =   6480
                  List            =   "frmequipment.frx":15BA
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Text            =   "ÌÊ„Ì"
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.TextBox TxtRent 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   2400
                  Width           =   840
               End
               Begin VB.TextBox txtHourdipp 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   2040
                  Width           =   1920
               End
               Begin VB.TextBox TxtHourCount 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   1680
                  Width           =   1920
               End
               Begin VB.TextBox TXTNotes 
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
                  Left            =   120
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   37
                  Top             =   2280
                  Width           =   3960
               End
               Begin VB.TextBox TXTUsedElectricPriceH 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   1920
               End
               Begin VB.TextBox TXTUsedPowerPriceH 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   960
                  Width           =   1920
               End
               Begin VB.TextBox TxtCode 
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
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.TextBox TxtVacName 
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
                  Left            =   5400
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„⁄œÂ «Ê «·«·…"
                  Top             =   -435
                  Visible         =   0   'False
                  Width           =   3960
               End
               Begin VB.TextBox TxtSerial 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   -75
                  Visible         =   0   'False
                  Width           =   1065
               End
               Begin VB.ComboBox CmbType 
                  BackColor       =   &H80000018&
                  Height          =   315
                  ItemData        =   "frmequipment.frx":15D0
                  Left            =   2280
                  List            =   "frmequipment.frx":15E0
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   3990
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin MSDataListLib.DataCombo DcFixedAssets 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   47
                  Top             =   480
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCEmp1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   48
                  Top             =   1200
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCEmp2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1560
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Rd 
                  Height          =   255
                  Index           =   0
                  Left            =   2520
                  TabIndex        =   50
                  Top             =   120
                  Width           =   1575
                  _Version        =   786432
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "—»ÿ „‰ „·› «·«’Ê·"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Rd 
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   51
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÌœÊÌ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÃ"
                  Height          =   285
                  Index           =   0
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   2760
                  Width           =   390
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”⁄… «·„⁄œ…"
                  Height          =   285
                  Index           =   15
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   2760
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„⁄œÂ «‰Ã·Ì“Ì"
                  Height          =   285
                  Index           =   14
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   840
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”»…  Õ„· «·„”«⁄œ"
                  Height          =   285
                  Index           =   15
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   2760
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”»…  Õ„· «·„‘€·"
                  Height          =   285
                  Index           =   14
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   2760
                  Width           =   2130
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„”«⁄œ"
                  Height          =   285
                  Index           =   12
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1560
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„‘€·"
                  Height          =   285
                  Index           =   11
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   1200
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰”»… «·«Ìﬁ«›"
                  Height          =   285
                  Index           =   10
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1920
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·«Ìﬁ«›"
                  Height          =   285
                  Index           =   9
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   2040
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «· √ÃÌ—"
                  Height          =   315
                  Index           =   8
                  Left            =   7680
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   2400
                  Width           =   2070
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·«Â·«ﬂ ›Ì «·”«⁄Â"
                  Height          =   315
                  Index           =   7
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   2310
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œ«œ «·”«⁄« "
                  Height          =   315
                  Index           =   6
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   1680
                  Width           =   2310
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê’› „Œ ’—"
                  Height          =   285
                  Index           =   5
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   2400
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «” Â·«ﬂ «·ﬂÂ—»«¡›Ì «·”«⁄Â"
                  Height          =   195
                  Index           =   4
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   1320
                  Width           =   2430
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «” Â·«ﬂ «·ÊﬁÊœ ›Ì «·”«⁄Â"
                  Height          =   315
                  Index           =   1
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   960
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„⁄œÂ"
                  Height          =   285
                  Index           =   0
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   480
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·„⁄œÂ"
                  Height          =   195
                  Index           =   3
                  Left            =   7905
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   630
                  Width           =   1710
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   3315
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   9825
               _cx             =   17330
               _cy             =   5847
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmequipment.frx":15F9
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   6480
            Left            =   10920
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   45
            Width           =   9885
            _cx             =   17436
            _cy             =   11430
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   615
               Left            =   -120
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   0
               Width           =   10005
               _cx             =   17648
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
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ê«’›« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   330
                  Index           =   16
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   45
                  Width           =   2430
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   5415
               Left            =   0
               TabIndex        =   93
               Top             =   600
               Width           =   9825
               _cx             =   17330
               _cy             =   9551
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
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmequipment.frx":170C
               ScrollTrack     =   0   'False
               ScrollBars      =   0
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   0
               Left            =   8655
               TabIndex        =   94
               Top             =   6000
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ— "
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
               ButtonImage     =   "frmequipment.frx":1781
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   1
               Left            =   7200
               TabIndex        =   95
               Top             =   6000
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
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
               ButtonImage     =   "frmequipment.frx":1D1B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   -60
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   10065
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   5
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   15
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
               TabIndex        =   6
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
            TabIndex        =   3
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
            TabIndex        =   2
            Top             =   510
            Visible         =   0   'False
            Width           =   945
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
                  Picture         =   "frmequipment.frx":22B5
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":264F
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":29E9
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":2D83
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":311D
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":34B7
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":3851
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmequipment.frx":3DEB
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   90
            TabIndex        =   7
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
            ButtonImage     =   "frmequipment.frx":4185
            ColorButton     =   14871017
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   555
            TabIndex        =   8
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
            ButtonImage     =   "frmequipment.frx":451F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1155
            TabIndex        =   9
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
            ButtonImage     =   "frmequipment.frx":48B9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1620
            TabIndex        =   10
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
            ButtonImage     =   "frmequipment.frx":4C53
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·„⁄œ«  / «·√·« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   90
            Width           =   3360
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1020
         Index           =   0
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7200
         Width           =   9825
         _cx             =   17330
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
         AutoSizeChildren=   0
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   6360
            TabIndex        =   13
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":4FED
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   4830
            TabIndex        =   14
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":5387
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   5595
            TabIndex        =   15
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":5721
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   4065
            TabIndex        =   16
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":5ABB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   3180
            TabIndex        =   17
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":5E55
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   7320
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   1290
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "frmequipment.frx":63EF
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Index           =   0
            Left            =   10605
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   1305
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "frmequipment.frx":6789
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   8325
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1350
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmequipment.frx":6B23
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   1425
            TabIndex        =   21
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "frmequipment.frx":6EBD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   330
            Left            =   2280
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   555
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "frmequipment.frx":7257
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   225
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2505
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   225
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmequipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Sub Retrivetitems()
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  bool = True

  With Me.FgItem
   j = .Rows
 If XPOptShowType(2).value = True Or XPOptShowType(0).value = True Then
Set Rs1 = New ADODB.Recordset
     sql = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, "
     sql = sql & "                  dbo.TblItems.fullcode , dbo.TblItemsUnits.unitid, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee , dbo.TblItemsUnits.UnitFactor"
     sql = sql & "      FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
     sql = sql & "                  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
     sql = sql & "                  dbo.TblItems ON dbo.TblItemsUnits.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
     sql = sql & "                  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
If XPOptShowType(2).value = True Then
     sql = sql & "  Where (dbo.TblItems.ItemID =" & val(Me.DcbItem.BoundText) & ")"
End If
     Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
   If Rs1.RecordCount > 0 Then
 .Rows = .Rows + Rs1.RecordCount
 Rs1.MoveFirst
        For i = j To .Rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
               .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
               .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
               .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
         If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
         Else
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
         End If
           .TextMatrix(i, .ColIndex("Capacity")) = val(Me.TxtCapacity.Text)
           Rs1.MoveNext
        Next i
   
  End If
  End If
      If XPOptShowType(1).value = True Then
          Dim GROUPIDS As String
          For k = 1 To ListGroupSelected.ListCount
          Set Rs1 = New ADODB.Recordset
        GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))
        If Len(GROUPIDS) > 2 Then GROUPIDS = mId(GROUPIDS, 2, Len(GROUPIDS))
        Debug.Print GROUPIDS
        If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
           sql = " SELECT dbo.TblItems.Fullcode,     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
           sql = sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee ,dbo.TblItemsUnits.UnitFactor"
           sql = sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
           sql = sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
           sql = sql & "            dbo.TblUnites RIGHT OUTER JOIN"
           sql = sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
           sql = sql & "  where dbo.TblItems.GroupID IN ( " & GROUPIDS & ")"
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
 j = .Rows
.Rows = .Rows + Rs1.RecordCount

        For i = j To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            End If
            .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
            .TextMatrix(i, .ColIndex("Capacity")) = val(Me.TxtCapacity.Text)
          Rs1.MoveNext
        Next i

    End If
       
       
         Next k
  End If
  
   End With
   
   
         DcbItem.Text = ""
        txtCode.Text = ""
   
End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow2
Case 1
RemoveGridAllRow2
Case 3
RemoveGridRow
Case 4
RemoveGridAllRow
End Select
End If
End Sub

Private Sub FgItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FgItem
If .ColKey(Col) <> "Capacity" Then
Cancel = True
End If
End With
End Sub

Private Sub Label8_Click()
Dim GROUPIDS, sql As String
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.XPOptShowType(1).value = True Then
 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)

            End If
            End If
End Sub
Private Sub Label7_Click()
    Dim i As Integer
    If Me.XPOptShowType(1).value = True Then
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End If
End Sub
Private Sub Label5_Click()
   If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub Label6_Click()
ListGroupSelected.Clear
End Sub
Private Sub DcbItem_Change()
DcbItem_Click (0)
End Sub




Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap
    'If DoPremis(Do_Delete, Me.name, True) = False Then
    '    Exit Sub
    'End If

    'If TxtVac_ID.text <> "" Then
    '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
    '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    Else
    MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    End If

    If MSGType = vbYes Then
      StrSQL = "Delete From FixedAssets Where CarsDataID =" & val(Me.TxtVac_ID.Text) & "  And FlgCarNotFixed = 2"
              Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblAssestes Where CarsDataID =" & val(Me.TxtVac_ID.Text) & "  And FlgCarNotFixed = 2"
              Cn.Execute StrSQL, , adExecuteNoRecords
         Cn.Execute "Delete from TblEquiCapacity where EquipID=" & val(Me.TxtVac_ID.Text) & " "
         Cn.Execute "Delete from TblEquipmentsDes where EquipID=" & val(Me.TxtVac_ID.Text) & " "
        RsSavRec.find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
        RsSavRec.delete
        CuurentLogdata ("D")
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
     FgItem.Clear flexClearScrollable, flexClearEverything
     FgItem.Rows = 1
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid1.Rows = 1
        '------------------------------ Move Next ---------------------------.
        FillGridWithData
        BtnNext_Click
    End If

    'End If
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

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    Dim Msg As String
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
    VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
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

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"
    RemoveGridAllRow
    ListGroupSelected.Clear
    FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.Rows = 1
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
Rd(0).value = True
Rd_Click (0)
    My_SQL = "TblEquipments"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
   ' CmbType.ListIndex = 0
    txtCode.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    TxtModFlg = "R"

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

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
        FindRec val(TxtVac_ID.Text)
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
Sub SaveFixed()
Dim sql As String
Dim StrSQL As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim FexdID As Double

If Me.TxtModFlg.Text = "N" Then
sql = "Select * from FixedAssets where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
FexdID = CStr(new_id("FixedAssets", "id", "", True))
Rs5.AddNew
Else
FexdID = val(DcFixedAssets.BoundText)
sql = "Select * from FixedAssets where id=" & FexdID & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
End If
Rs5("id").value = FexdID
Rs5("CarsDataID").value = val(TxtVac_ID.Text)
Rs5("FlgCarNotFixed").value = 2
Rs5("ISEQUP").value = 1
Rs5("HaveDepreciation").value = 1
Rs5("PurchasePrice").value = 1
'Rs5("branch_no").value = val(dcBranch.BoundText)
Rs5("NameE").value = TxtEqupNameE.Text
Rs5("Name").value = TxtEqupName.Text
Rs5("code").value = txtCode.Text
Rs5.update
sql = "Update TblEquipments set fixedAssetid=" & FexdID & "  where id =" & val(TxtVac_ID.Text) & ""
Cn.Execute sql
SaveAssest FexdID
End Sub
Sub SaveAssest(Optional FexdID As Double = 0)
Dim sql As String
Dim StrSQL As String
Dim Msg As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
If Me.TxtModFlg.Text = "N" Then
sql = "Select * from TblAssestes where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Rs5.AddNew
Else
sql = "Select * from TblAssestes where AsFixedID=" & FexdID & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs5.RecordCount <= 0 Then
Rs5.AddNew
End If
End If
Rs5("CarsDataID").value = val(TxtVac_ID.Text)
Rs5("FlgCarNotFixed").value = 2
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ „·› «·„⁄œ« "
Else
Msg = "From Equipment File"
End If
Rs5("AsFixedID").value = FexdID
Rs5("AsDes").value = Msg
Rs5("AsName").value = TxtEqupName.Text
Rs5("AsCode").value = val(txtCode.Text)
Rs5.update
End Sub
Private Sub btnSave_Click()
   On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

   ' For Each CtrlTxt In Me.Controls
'
'        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
'            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
'                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
'                CtrlTxt.SetFocus
'                Exit Sub
'            End If
'        End If
'
'    Next
If Rd(0).value = False And Rd(1).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ Õ«·… «·—»ÿ "
Else
MsgBox "Please Select Type Link"
End If
Exit Sub
End If
    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblEquipments", "name", Trim(TxtVacName.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
End Sub





Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FixedAssetsSearch.RetrunType = 2
        FixedAssetsSearch.show vbModal
  
    End If
End Sub
Private Sub RemoveGridAllRow2()
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.Rows = 1
End Sub

Private Sub RemoveGridAllRow()
 FgItem.Clear flexClearScrollable, flexClearEverything
            FgItem.Rows = 1
'    ReLineGrid
End Sub
Private Sub DcbItem_Click(Area As Integer)
Me.TxtItemCode.Text = GetItemCode(val(Me.DcbItem.BoundText))
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
Retrivetitems
End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As New ClsDataCombos
    FillMylist
    ScreenNameArabic = "»Ì«‰«  «·„⁄œ«  Ê «·«·«  "
    ScreenNameEnglish = "Equipments Data "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
     Dcombos.GetFixedAssets Me.DcFixedAssets, True
     Dcombos.GetEmployees Me.DCEmp1
     Dcombos.GetEmployees Me.DCEmp2
     Dcombos.GetItemsNamesupdate Me.DcbItem
    My_SQL = "TblEquipments"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
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

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ   " & txtCode.Text & CHR(13) & " «·«”„ " & TxtVacName.Text & CHR(13) & "ﬁÌ„… «” Â·«ﬂ «·ÊﬁÊœ ›Ì «·”«⁄Â  " & TXTUsedPowerPriceH & CHR(13) & " ﬁÌ„… «” Â·«ﬂ «·ﬂÂ—»«¡›Ì «·”«⁄Â  " & TXTUsedElectricPriceH & CHR(13) & " „·«ÕŸ«    " & TxtNotes
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Code   " & txtCode.Text & CHR(13) & " Name " & TxtVacName.Text & CHR(13) & " Used Power Per Hour   " & TXTUsedPowerPriceH & CHR(13) & " Used Electric Per Hour  " & TXTUsedElectricPriceH & CHR(13) & " Remarks   " & TxtNotes
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D", "", ""
    End If
    
End Function

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
C1Tab1.Caption = " Date|Capacity|Specifications"
    Me.Caption = "Equipment Data"
    Label1(2).Caption = Me.Caption
    Cmd(0).Caption = "Delete"
    Cmd(1).Caption = "Delete All"
    Label1(16).Caption = "Specifications"
    lbl(14).Caption = "Arrying operator %"
    lbl(15).Caption = "Carrying Assistant %"
Rd(0).Caption = "Link FA"
Rd(1).Caption = "Manual"
XPOptShowType(0).RightToLeft = False
XPOptShowType(1).RightToLeft = False
XPOptShowType(2).RightToLeft = False
XPOptShowType(0).Caption = "All Items"
XPOptShowType(1).Caption = "Groups"
XPOptShowType(2).Caption = "Select Item"
ISButton2.Caption = "Add"
lbl(1).Caption = "Capacity"
Label1(15).Caption = "Capacity"
Cmd(3).Caption = "Delete"
Cmd(4).Caption = "Delete All"
lbl(0).Caption = "Kg"
With FgItem
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
.TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
.TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
.TextMatrix(0, .ColIndex("Capacity")) = "Capacity"
End With

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Equ Name"
        .TextMatrix(0, .ColIndex("UsedPowerPriceH")) = "Used Power Price H"
        .TextMatrix(0, .ColIndex("UsedElectricPriceH")) = "Used Electric Price H"
    End With
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("ShortDes")) = "Short"
        .TextMatrix(0, .ColIndex("Descrip")) = "Description"
    End With

Label1(11).Caption = "Driver"
Label1(14).Caption = "Name English"
Label1(12).Caption = "Helper"
Label1(9).Caption = "Stop Cost"
Label1(10).Caption = "Stop %"
 ChKLockeq.Caption = "Lock"
 
With CboInterval
.Clear
.AddItem "Day"
.AddItem "Month"
.AddItem "Year"

End With



Label1(6).Caption = "Hours Counter"
Label1(7).Caption = "Dip Value"
Label1(8).Caption = "Rent Value"

    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name"
    Label1(1).Caption = "Used Power Price H"
    Label1(4).Caption = "Used Electric Price H"
    Label1(5).Caption = "Notes"

    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
ISButton1.Caption = "Search"
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblEquipments", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    TxtVac_ID.Text = IIf(StrRecID <> "", StrRecID, 0)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "E" Then
    Cn.Execute "Delete from TblEquiCapacity where EquipID=" & val(Me.TxtVac_ID.Text) & " "
    Cn.Execute "Delete from TblEquipmentsDes where EquipID=" & val(Me.TxtVac_ID.Text) & " "
    End If
If Rd(0).value = True Then
txtCode.Text = getFixedAsstName(val(DcFixedAssets.BoundText), "Code")
If SystemOptions.UserInterface = ArabicInterface Then
TxtEqupNameE.Text = getFixedAsstName(val(DcFixedAssets.BoundText), "NameE")
TxtVacName.Text = DcFixedAssets.Text
Else
TxtVacName.Text = getFixedAsstName(val(DcFixedAssets.BoundText), "Name")
TxtEqupNameE.Text = DcFixedAssets.Text
End If
Else
TxtVacName.Text = TxtEqupName.Text
End If

    RsSavRec.Fields("nameE").value = IIf(TxtEqupNameE.Text <> "", Trim(TxtEqupNameE.Text), Null)
    RsSavRec.Fields("name").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("Code").value = IIf(Me.txtCode.Text <> "", Trim(txtCode.Text), Null)
    RsSavRec.Fields("UsedPowerPriceH").value = IIf(IsNumeric(Me.TXTUsedPowerPriceH.Text), val(TXTUsedPowerPriceH.Text), 0)
    RsSavRec.Fields("UsedElectricPriceH").value = IIf(IsNumeric(Me.TXTUsedElectricPriceH.Text), val(TXTUsedElectricPriceH.Text), 0)
    RsSavRec.Fields("HourCount").value = IIf(IsNumeric(Me.TxtHourCount.Text), val(TxtHourCount.Text), 0)
    RsSavRec.Fields("Hourdipp").value = IIf(IsNumeric(Me.txtHourdipp.Text), val(txtHourdipp.Text), 0)
    RsSavRec.Fields("Notes").value = IIf(Me.TxtNotes.Text <> "", Trim(TxtNotes.Text), Null)
    RsSavRec.Fields("fixedAssetid").value = IIf(DcFixedAssets.BoundText <> "", val(DcFixedAssets.BoundText), Null)
   
   RsSavRec.Fields("empID1").value = IIf(DCEmp1.BoundText <> "", val(DCEmp1.BoundText), Null)
   RsSavRec.Fields("empID2").value = IIf(DCEmp2.BoundText <> "", val(DCEmp2.BoundText), Null)
   RsSavRec.Fields("helper").value = IIf(Me.Txthelper.Text <> "", val(Txthelper.Text), Null)
   RsSavRec.Fields("Employer").value = IIf(Me.TxtEmployer.Text <> "", val(TxtEmployer.Text), Null)
   If Rd(1).value = True Then
   RsSavRec.Fields("TypeEqup").value = 1
   Else
   RsSavRec.Fields("TypeEqup").value = 0
   End If
   RsSavRec.Fields("EqupName").value = IIf(Me.TxtEqupName.Text <> "", Trim(TxtEqupName.Text), Null)
 
If ChKLockeq.value = vbChecked Then
        RsSavRec.Fields("ChKLockeq").value = 1
Else
       RsSavRec.Fields("ChKLockeq").value = 0
End If
 
   RsSavRec.Fields("Stopvalue").value = IIf(IsNumeric(Me.TxtStopvalue.Text), val(TxtStopvalue.Text), 0)
   RsSavRec.Fields("StopPercentage").value = IIf(IsNumeric(Me.TxtStopPercentage.Text), val(TxtStopPercentage.Text), 0)
   
   RsSavRec.Fields("Rent").value = IIf(IsNumeric(Me.TxtRent.Text), val(TxtRent.Text), 0)
   RsSavRec.Fields("Interval").value = val(CboInterval.ListIndex)
   RsSavRec.Fields("Capacites").value = val(TxtCapacites.Text)
   RsSavRec.update
  ''////////////////
    Dim RsDevsub As ADODB.Recordset
    Dim StrSQL As String
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEquiCapacity Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.FgItem
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("EquipID").value = val(Me.TxtVac_ID.Text)
                RsDevsub("GroupID").value = IIf((.TextMatrix(i, .ColIndex("GroupID"))) = "", 0, val(.TextMatrix(i, .ColIndex("GroupID"))))
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", 0, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("Capacity").value = IIf((.TextMatrix(i, .ColIndex("Capacity"))) = "", 0, val(.TextMatrix(i, .ColIndex("Capacity"))))
       RsDevsub.update
      End If
     Next i
    End With
    '''''/////////
        Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEquipmentsDes Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If (.TextMatrix(i, .ColIndex("Descrip"))) <> "" Then
       RsDevsub.AddNew
                RsDevsub("EquipID").value = val(Me.TxtVac_ID.Text)
                RsDevsub("Descrip").value = IIf((.TextMatrix(i, .ColIndex("Descrip"))) = "", "", (.TextMatrix(i, .ColIndex("Descrip"))))
                RsDevsub("ShortDes").value = IIf((.TextMatrix(i, .ColIndex("ShortDes"))) = "", "", (.TextMatrix(i, .ColIndex("ShortDes"))))
       RsDevsub.update
      End If
     Next i
    End With
    
     If Rd(1).value = True Then
 SaveFixed
 End If
    CuurentLogdata
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGridWithData
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtEqupNameE.Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    Me.txtCode.Text = IIf(IsNull(RsSavRec.Fields("code").value), "", RsSavRec.Fields("code").value)
    Me.TXTUsedPowerPriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("UsedPowerPriceH").value), 0, RsSavRec.Fields("UsedPowerPriceH").value)
    Me.TXTUsedElectricPriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("UsedElectricPriceH").value), 0, RsSavRec.Fields("UsedElectricPriceH").value)
    Me.TxtHourCount.Text = IIf(Not IsNumeric(RsSavRec.Fields("HourCount").value), 0, RsSavRec.Fields("HourCount").value)
    Me.txtHourdipp.Text = IIf(Not IsNumeric(RsSavRec.Fields("Hourdipp").value), 0, RsSavRec.Fields("Hourdipp").value)
    Me.TxtEqupName.Text = IIf(IsNull(RsSavRec.Fields("EqupName").value), "", RsSavRec.Fields("EqupName").value)
    Me.TxtNotes.Text = IIf(IsNull(RsSavRec.Fields("Notes").value), "", RsSavRec.Fields("Notes").value)
    DcFixedAssets.BoundText = IIf(IsNull(RsSavRec.Fields("fixedAssetid").value), "", RsSavRec.Fields("fixedAssetid").value)
    TxtCapacites.Text = IIf(IsNull(RsSavRec.Fields("Capacites").value), 0, RsSavRec.Fields("Capacites").value)
    DCEmp1.BoundText = IIf(IsNull(RsSavRec.Fields("empID1").value), "", RsSavRec.Fields("empID1").value)
    DCEmp2.BoundText = IIf(IsNull(RsSavRec.Fields("empID2").value), "", RsSavRec.Fields("empID2").value)
    Me.Txthelper.Text = IIf(IsNull(RsSavRec.Fields("helper").value), "", RsSavRec.Fields("helper").value)
    Me.TxtEmployer.Text = IIf(IsNull(RsSavRec.Fields("Employer").value), "", RsSavRec.Fields("Employer").value)



    Me.TxtStopvalue.Text = IIf(Not IsNumeric(RsSavRec.Fields("Stopvalue").value), 0, RsSavRec.Fields("Stopvalue").value)
Me.TxtStopPercentage.Text = IIf(Not IsNumeric(RsSavRec.Fields("StopPercentage").value), 0, RsSavRec.Fields("StopPercentage").value)
If IsNull(RsSavRec.Fields("TypeEqup").value) Then
    Rd(0).value = True
Else
           If (RsSavRec.Fields("TypeEqup").value) = 1 Then
                Rd(1).value = True
            Else
                Rd(0).value = True
            End If
 
End If

If IsNull(RsSavRec.Fields("ChKLockeq").value) Then
    ChKLockeq.value = vbUnchecked
Else
           If (RsSavRec.Fields("ChKLockeq").value) = True Then
                ChKLockeq.value = vbChecked
            Else
                ChKLockeq.value = vbUnchecked
            End If
 
End If
Me.TxtRent.Text = IIf(Not IsNumeric(RsSavRec.Fields("Rent").value), 0, RsSavRec.Fields("Rent").value)
Me.CboInterval.ListIndex = IIf(Not IsNumeric(RsSavRec.Fields("Interval").value), -1, RsSavRec.Fields("Interval").value)



    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
FullGridData
    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Private Sub ISButton1_Click()
Load FrmSearchEqupment
FrmSearchEqupment.show vbModal
End Sub

Private Sub RemoveGridRow2()
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.FgItem
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub


 Sub FullGridData()
  On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
     Me.FgItem.Clear flexClearScrollable, flexClearEverything
    FgItem.Rows = 1
sql = " SELECT     dbo.TblEquiCapacity.EquipID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblEquiCapacity.ItemID, dbo.TblEquiCapacity.Capacity, "
sql = sql & "                       dbo.TblEquiCapacity.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.Fullcode AS GroupFullcode, dbo.Groups.GroupNamee"
sql = sql & "  FROM         dbo.TblEquiCapacity LEFT OUTER JOIN"
sql = sql & "                       dbo.Groups ON dbo.TblEquiCapacity.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.TblEquiCapacity.ItemID = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblEquiCapacity.EquipID = " & val(TxtVac_ID.Text) & ")"
Set Rs1 = New ADODB.Recordset
  Dim i As Integer
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With Me.FgItem
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), 0, Rs1("GroupID").value)
                   .TextMatrix(i, .ColIndex("Capacity")) = IIf(IsNull(Rs1("Capacity").value), 0, Rs1("Capacity").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                   .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
                   Else
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                   .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
    ''////////
VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.Rows = 1
sql = " SELECT *  from TblEquipmentsDes"
sql = sql & "  Where (EquipID = " & val(TxtVac_ID.Text) & ")"
Set Rs1 = New ADODB.Recordset
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     With Me.VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ShortDes")) = IIf(IsNull(Rs1("ShortDes").value), "", Rs1("ShortDes").value)
                   .TextMatrix(i, .ColIndex("Descrip")) = IIf(IsNull(Rs1("Descrip").value), "", Rs1("Descrip").value)
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub

Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
  

  sql = " SELECT * from  Groups where GroupID>1"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function

Private Sub Rd_Click(Index As Integer)
If Rd(1).value = True Then
DcFixedAssets.Visible = False
TxtEqupName.Visible = True
ElseIf Rd(0).value = True Then
DcFixedAssets.Visible = True
TxtEqupName.Visible = False
End If
End Sub





Private Sub TxtCapacites_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtCapacites.Text, 0)
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.Text = "" Then
            Me.DcbItem.BoundText = ""
        Else
            Me.DcbItem.BoundText = GetItemID(Trim$(Me.TxtItemCode.Text))
        End If
    End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "id=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
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

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
   '     BtnUpdate.Enabled = False
        ' btnNext.Enabled = False
        ' btnPrevious.Enabled = False
        ' btnFirst.Enabled = False
        ' btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

       ' BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.Text = "E" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
    '    BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()
 
    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblEquipments order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
           
                .TextMatrix(i, .ColIndex("UsedPowerPriceH")) = IIf(IsNull(rs.Fields("UsedPowerPriceH").value), "", rs.Fields("UsedPowerPriceH").value)
            
                .TextMatrix(i, .ColIndex("UsedElectricPriceH")) = IIf(IsNull(rs.Fields("UsedElectricPriceH").value), "", rs.Fields("UsedElectricPriceH").value)
            
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs.Fields("Code").value), "", rs.Fields("Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
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

Private Function CheckDelCountry(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid1
If (Row = .Rows - 1) And .TextMatrix(Row, .ColIndex("Descrip")) <> "" Then
.Rows = .Rows + 1
End If
End With
End Sub

