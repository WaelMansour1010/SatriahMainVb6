VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmVisa 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… »Ì«‰«  «· √‘Ì—…"
   ClientHeight    =   9915
   ClientLeft      =   4665
   ClientTop       =   2025
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmVisa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12375
      _cx             =   21828
      _cy             =   17489
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8850
         Left            =   30
         TabIndex        =   1
         Top             =   -105
         Width           =   17325
         _cx             =   30559
         _cy             =   15610
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
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8430
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   17235
            _cx             =   30401
            _cy             =   14870
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   915
               Index           =   5
               Left            =   0
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   0
               Width           =   12315
               _cx             =   21722
               _cy             =   1614
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
               Picture         =   "FrmVisa.frx":038A
               Caption         =   "‘«‘… »Ì«‰«  «· √‘Ì—«   "
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
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   30
                  Top             =   90
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
                  ButtonImage     =   "FrmVisa.frx":1064
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
                  Left            =   630
                  TabIndex        =   31
                  Top             =   90
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
                  ButtonImage     =   "FrmVisa.frx":13FE
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
                  Left            =   2220
                  TabIndex        =   32
                  Top             =   90
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
                  ButtonImage     =   "FrmVisa.frx":1798
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
                  Left            =   1155
                  TabIndex        =   33
                  Top             =   90
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
                  ButtonImage     =   "FrmVisa.frx":1B32
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
               Height          =   8400
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   16545
               _cx             =   29184
               _cy             =   14817
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
               Caption         =   "„œ… «·’·«ÕÌ…"
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
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  ’«Õ» «·⁄„·"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   1245
                  Index           =   6
                  Left            =   2970
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   2940
                  Width           =   9075
                  Begin VB.TextBox txtKafelID 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   6150
                     MaxLength       =   30
                     TabIndex        =   57
                     Top             =   330
                     Width           =   1845
                  End
                  Begin VB.TextBox txtkafeltel 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   6150
                     MaxLength       =   30
                     TabIndex        =   56
                     Top             =   750
                     Width           =   1845
                  End
                  Begin VB.TextBox txtkafeladd 
                     Alignment       =   1  'Right Justify
                     Height          =   465
                     Left            =   2970
                     MaxLength       =   150
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   55
                     Top             =   690
                     Width           =   2565
                  End
                  Begin MSDataListLib.DataCombo DcbKafelName 
                     Height          =   315
                     Left            =   2910
                     TabIndex        =   58
                     Top             =   330
                     Width           =   2565
                     _ExtentX        =   4524
                     _ExtentY        =   582
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄‰Ê«‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   14
                     Left            =   5550
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   720
                     Width           =   585
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«”„"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   13
                     Left            =   5550
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   330
                     Width           =   585
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· ·Ì›Ê‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   12
                     Left            =   8070
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   750
                     Width           =   705
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   11
                     Left            =   8100
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   300
                     Width           =   585
                  End
               End
               Begin VB.ComboBox DcbPeriodsID 
                  Height          =   330
                  ItemData        =   "FrmVisa.frx":1ECC
                  Left            =   4050
                  List            =   "FrmVisa.frx":1ED9
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   1680
                  Width           =   1200
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
                  Height          =   375
                  Left            =   5250
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   1680
                  Width           =   1140
               End
               Begin VB.TextBox TxtOrder 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Left            =   4050
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   1185
                  Width           =   2340
               End
               Begin VB.TextBox TxtVisa 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   1185
                  Width           =   2340
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   390
                  Left            =   6150
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   510
               End
               Begin VB.TextBox xptxtid 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Left            =   8580
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   1185
                  Width           =   2220
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Index           =   0
                  Left            =   -4290
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   11790
                  Width           =   2370
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   6345
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   2310
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2490
                  Left            =   0
                  TabIndex        =   7
                  Top             =   5865
                  Width           =   12225
                  _cx             =   21564
                  _cy             =   4392
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   27
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVisa.frx":1EEC
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
               Begin MSComCtl2.DTPicker StarDate 
                  Height          =   375
                  Left            =   8580
                  TabIndex        =   12
                  Top             =   1680
                  Width           =   2220
                  _ExtentX        =   3916
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   184745985
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal StarDateH 
                  Height          =   360
                  Left            =   8580
                  TabIndex        =   39
                  Top             =   2145
                  Width           =   2220
                  _ExtentX        =   3916
                  _ExtentY        =   635
               End
               Begin MSComCtl2.DTPicker EndDate 
                  Height          =   375
                  Left            =   135
                  TabIndex        =   40
                  Top             =   1680
                  Width           =   2340
                  _ExtentX        =   4128
                  _ExtentY        =   661
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   184745985
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal EndDateH 
                  Height          =   360
                  Left            =   135
                  TabIndex        =   41
                  Top             =   2145
                  Width           =   2340
                  _ExtentX        =   4128
                  _ExtentY        =   635
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid2 
                  Height          =   1560
                  Left            =   0
                  TabIndex        =   46
                  Top             =   4215
                  Width           =   12225
                  _cx             =   21564
                  _cy             =   2752
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   28
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVisa.frx":22C5
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
               Begin MSDataListLib.DataCombo cmbOffice 
                  Height          =   330
                  Left            =   8280
                  TabIndex        =   50
                  Top             =   2520
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   582
                  _Version        =   393216
                  ListField       =   "6"
                  BoundColumn     =   ""
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker txtArriveDate 
                  Height          =   375
                  Left            =   4950
                  TabIndex        =   51
                  Top             =   2490
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   184745985
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal txtArriveDateH 
                  Height          =   375
                  Left            =   2685
                  TabIndex        =   52
                  Top             =   2490
                  Width           =   2235
                  _ExtentX        =   3942
                  _ExtentY        =   661
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Ê’Ê·"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   10
                  Left            =   5970
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   2520
                  Width           =   1950
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ﬂ » «·„›Ê÷"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   9
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   2520
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄œœ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   6
                  Left            =   6150
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   2145
                  Width           =   1950
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   5
                  Left            =   3795
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   2145
                  Width           =   1950
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«‰ Â«¡"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   4
                  Left            =   1845
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   1965
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·’·«ÕÌ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   3
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   1680
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ÿ·»"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   2
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   1185
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· √‘Ì—…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   0
                  Left            =   1845
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   1185
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒÂ«"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   8
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   1965
                  Width           =   1905
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   7
                  Left            =   10110
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1185
                  Width           =   1920
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   510
                  Left            =   14985
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   1185
                  Width           =   930
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   9105
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   105
               Width           =   1245
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1140
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   8775
         Width           =   12375
         _cx             =   21828
         _cy             =   2011
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   375
            Left            =   12915
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   105
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVisa.frx":26C4
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   450
            Left            =   13875
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   255
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   794
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
            ButtonImage     =   "FrmVisa.frx":2A5E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   330
            Left            =   15180
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   165
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
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
            ButtonImage     =   "FrmVisa.frx":2DF8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   630
            Index           =   0
            Left            =   9045
            TabIndex        =   19
            Top             =   435
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   1
            Left            =   8070
            TabIndex        =   20
            Top             =   435
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   2
            Left            =   7140
            TabIndex        =   21
            Top             =   435
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   3
            Left            =   6120
            TabIndex        =   22
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   4
            Left            =   4995
            TabIndex        =   23
            Top             =   435
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   6
            Left            =   1845
            TabIndex        =   24
            Top             =   435
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   5
            Left            =   3960
            TabIndex        =   25
            Top             =   435
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1111
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   420
            Left            =   9915
            TabIndex        =   28
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
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
            MICON           =   "FrmVisa.frx":3192
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   630
            Index           =   7
            Left            =   3000
            TabIndex        =   45
            Top             =   435
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1111
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
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   255
            Width           =   1890
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   270
            Width           =   1605
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "FrmVisa.frx":31AE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Dim sql As String
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long







'Private Sub ChkDetails_Click()
'    FillGridWithData
'End Sub

'Private Sub ALLButton1_Click()
'    FrmShowCol1.show
'End Sub

'Function check_previous_dev(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from notes where salary=" & year & Month
'
'    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs.RecordCount = 0 Then
'        check_previous_dev = False
'    Else
'        check_previous_dev = True
'    End If
'
'End Function

'Function check_previous_dev1(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
'
'    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs.RecordCount = 0 Then
'        check_previous_dev1 = False
'    Else
'        check_previous_dev1 = True
'    End If
'
'End Function
'
'Function Create_dev()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
''    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'    Dim notes_serial As String
'    Dim notes_id As String
'
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            GoTo ErrTrap
'
'        End If
'    End If
'
'    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "
'
'    Dim StrSQL As String
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From Notes where NoteType=66 order by NoteID"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    notes_id = CStr(new_id("Notes", "NoteID", "", True))
'    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
'
'    rs.AddNew
'    rs("NoteID").value = notes_id
'    rs("NoteSerial").value = notes_serial '
''    rs("Note_Value").value = Null
 '   rs("Remark").value = Msg
'
''    rs("NoteType").value = 66
 '   rs("NoteDate").value = Date
 '   rs("UserID").value = user_id
 '   rs.update
 '
 '   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
 ''
  '  Dim line_no As Integer
  '  line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
'            Else
'                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
'            StrAccountCode = Employee_account
'
'            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 2
'
'        Next i
'
'    End With
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
'
'End Function
'
'Function Create_dev1()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
'    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            GoTo ErrTrap
'
'        End If
'    End If
'
'    'StrAccountCode = Account_Code_dynamic
'
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'    Dim line_no As Integer
'    line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
''            Else
 '               Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
'            StrAccountCode = Employee_account
'
'            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 2
'
'        Next i
'
'    End With
'
'    Set rs = New ADODB.Recordset
'    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'    rs.AddNew
'
'    rs("voucher_id").value = LngDevID
'
'    rs.update
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
'
'End Function

'Private Sub ALLButton2_Click()
'    'Dcemp.text = ""
'
'    dcproject.text = ""
'    FillGridWithData
'
'    DoEvents
'    Create_dev
'    CmdOk_Click
'End Sub



'Private Sub CboPayMentType_Click()
'    CboPayMentType_Change
'End Sub

'Private Sub CboYear_Click()
'    CmdOk_Click
'End Sub

'Private Sub Check1_Click()
'Exit Sub
'    If Check1.value = vbChecked Then
'        get_all_employee
'    Else
'
''        With Me.Grid
 '           .Rows = 2
 '           .Clear flexClearScrollable
 '       End With
'
'    End If
''
'End Sub

'Private Sub CmbMonth_Click()
'    CmdOk_Click
    'FillGridWithData
'End Sub

'Private Sub CmdExit_Click()
'    Unload Me
'End Sub



'Private Sub CmdPrint_Click()
'    On Error Resume Next
'    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
'    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
'    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

'    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
'End Sub



Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.TxtOrder.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  «œŒ«· —ﬁ„ «·ÿ·»..!!"
            Else
                Msg = "Request No. is a must"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtOrder.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        If Trim(Me.TxtVisa.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  «œŒ«· —ﬁ„ «· √‘Ì—…..!!"
            Else
                Msg = "Visa No. is a must"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtVisa.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TbVisaDeti where VisaID=" & val(Me.xptxtid.Text)
   
    End If
    
    rs("ID").value = xptxtid.Text
   
    rs("StarDate").value = StarDate.value
    rs("StarDateH").value = StarDateH.value
    
    rs("ArriveDate").value = txtArriveDate.value
    rs("ArriveDateH").value = txtArriveDateH.value
    
    rs("EndDate").value = EndDate.value
    rs("EndDateH").value = EndDateH.value
    rs("OrderNo").value = IIf(Me.TxtOrder.Text = "", "", Me.TxtOrder.Text)
    rs("VisaNo").value = IIf(Me.TxtVisa.Text = "", "", Me.TxtVisa.Text)
    rs("Priod").value = IIf(Me.TxtPeriods.Text = "", Null, Me.TxtPeriods.Text)
    rs("DMYPriod").value = IIf(val(Me.TxtVisa.Text) = -1, -Null, val(Me.DcbPeriodsID.ListIndex))
    rs("OfficeId").value = val(Me.cmbOffice.BoundText)
    
    rs("KafelID").value = IIf(txtKafelID.Text = "", Null, Trim(txtKafelID.Text))
    rs("KafelName").value = IIf(Me.DcbKafelName.Text = "", Null, Trim(DcbKafelName.Text))
    
    rs("kafeltel").value = IIf(txtkafeltel.Text = "", Null, Trim(txtkafeltel.Text))
    rs("kafeladd").value = IIf(txtkafeladd.Text = "", Null, Trim(txtkafeladd.Text))
    
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TbVisaDeti", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.Grid
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Notional")) <> "" Then
                RsDev.AddNew
                RsDev("Type").value = 0
                RsDev("VisaID").value = Me.xptxtid.Text
                RsDev("EmpID").value = val(.TextMatrix(i, .ColIndex("Emp_id")))
                RsDev("HododNo").value = .TextMatrix(i, .ColIndex("HododNo"))
                RsDev("JobID").value = val(.TextMatrix(i, .ColIndex("JobID")))
                RsDev("NotionalID").value = val(.TextMatrix(i, .ColIndex("NotionalID")))
                RsDev("CityID").value = val(.TextMatrix(i, .ColIndex("CityID")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("OfficeID").value = val(.TextMatrix(i, .ColIndex("OfficeID")))
                RsDev("remarks").value = Trim(.TextMatrix(i, .ColIndex("remarks")))
                 
                RsDev.update
             If val(.TextMatrix(i, .ColIndex("Emp_id"))) <> 0 Then
             StrSQL = "Update TblEmployee set VisaNo='" & TxtVisa.Text & "',NationlID=" & val(.TextMatrix(i, .ColIndex("NotionalID"))) & ","
             StrSQL = StrSQL & " JobTypeID1=" & val(.TextMatrix(i, .ColIndex("JobID"))) & ", hdodno='" & .TextMatrix(i, .ColIndex("HododNo")) & "'   ,"
             If Trim(txtkafeladd) <> "" Then
                StrSQL = StrSQL & " kafeladd='" & Trim(txtkafeladd) & "',"
             End If
             If Trim(txtkafeltel) <> "" Then
                StrSQL = StrSQL & " kafeltel='" & Trim(txtkafeltel) & "',"
             End If
             
             If Trim(DcbKafelName.Text) <> "" Then
                StrSQL = StrSQL & " KafelName='" & Trim(DcbKafelName.Text) & "',"
             End If
             
             
             If Trim(txtKafelID.Text) <> "" Then
                StrSQL = StrSQL & " KafelID='" & Trim(txtKafelID.Text) & "',"
             End If
             
             
             StrSQL = StrSQL & " OfficeID=" & val(.TextMatrix(i, .ColIndex("OfficeID"))) & ""
             StrSQL = StrSQL & " where Emp_ID=" & val(.TextMatrix(i, .ColIndex("Emp_id"))) & ""
             Cn.Execute StrSQL
             End If
            End If
            
            '
        Next i

    End With
    
 '''//
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TbVisaDeti", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    

    With Me.GRID2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("count")) <> "" Then
         
                RsDev.AddNew
                RsDev("Type").value = 1
                RsDev("VisaID").value = Me.xptxtid.Text
               ' RsDev("EmpID").value = val(.TextMatrix(i, .ColIndex("Emp_id")))
               ' RsDev("HododNo").value = .TextMatrix(i, .ColIndex("HododNo"))
                RsDev("JobID").value = val(.TextMatrix(i, .ColIndex("JobID")))
                RsDev("NotionalID").value = val(.TextMatrix(i, .ColIndex("NotionalID")))
                 RsDev("CityID").value = val(.TextMatrix(i, .ColIndex("CityID")))
                 RsDev("count").value = val(.TextMatrix(i, .ColIndex("count")))
                 RsDev("Place").value = .TextMatrix(i, .ColIndex("Place"))
                 RsDev("OfficeID").value = val(.TextMatrix(i, .ColIndex("OfficeID")))
                RsDev("remarks").value = Trim(.TextMatrix(i, .ColIndex("remarks")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
    Cn.CommitTrans
    BeginTrans = False
 
 sql = "SELECT DISTINCT KafelName, KafelName AS KafelNames"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (KafelName IS NULL)) "
fill_combo DcbKafelName, sql
Retrive

    Select Case Me.TxtModFlg.Text

        Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = "Record was saved successfully" & CHR(13)
                Msg = Msg + "Do you want to enter new data ?"
            End If

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Record edited and saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
            Me.xptxtid.Text = CStr(new_id("TbVisa", "ID", "", True))
       
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
             GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 2
            Grid.Enabled = True
GRID2.Enabled = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

           ' If ChKauto.value = vbChecked Then
           '     If SystemOptions.UserInterface = ArabicInterface Then
           '         MsgBox " ·« Ì„ﬂ‰  ⁄œÌ·  Œ’Ì’ «·Ì ", vbCritical
           '     Else
           '         MsgBox " Can't Delete Auto Employee Allocation ", vbCritical
           '     End If
'
'                Exit Sub
'            End If

            TxtModFlg.Text = "E"
            GRID2.Rows = GRID2.Rows + 1
            GRID2.Enabled = True

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

             Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
'wael
 'General_Search.send_form = "visa"
 'wael
           ' Load General_Search
           'wael
            


          'General_Search.show
'wael
        Case 6
            Unload Me

        Case 7
                 If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report
         End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If xptxtid.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "This Record data will be deleted" & CHR(13)
            Msg = Msg + "Are you sure you want delete this record "
        End If
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TbVisa Where id=" & val(Me.xptxtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                 Cn.Execute "delete TbVisaDeti where VisaID=" & val(Me.xptxtid)

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
                    GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 2
                    clear_all Me
                    TxtModFlg_Change
                   ' XPTxtCurrent.Caption = 0
                   ' XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This action is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
'Private Sub Dcdep_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub Dcedara_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub Dcemp_Click(Area As Integer)
'    CmdOk_Click
'End Sub

'Private Sub DCmboEmp_Click(Area As Integer)
'    FillGridWithData
'End Sub

'Function SHow_grig_col()
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    With Grid
''
 '       If rs2("s1").value = True Then
 '           .ColHidden(.ColIndex("Emp_Code")) = False
 '       Else
 '           .ColHidden(.ColIndex("Emp_Code")) = True
 '       End If
 '
 '       If rs2("s2").value = True Then
 '           .ColHidden(.ColIndex("Emp_Name")) = False
 '       Else
 '           .ColHidden(.ColIndex("Emp_Name")) = True
 '       End If
 '
 '       If rs2("s3").value = True Then
 '           .ColHidden(.ColIndex("Emp_Salary")) = False
 '       Else
 ''           .ColHidden(.ColIndex("Emp_Salary")) = True
  '      End If
  '
  '      If rs2("s4").value = True Then
  '          .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
  '      Else
  '          .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
  '      End If
  ''
   '     If rs2("s5").value = True Then
   '         .ColHidden(.ColIndex("Emp_Salary_bus")) = False
   '     Else
   ''         .ColHidden(.ColIndex("Emp_Salary_bus")) = True
    '    End If
    '
    '    If rs2("s6").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_food")) = False
    '    Else
    '        .ColHidden(.ColIndex("Emp_Salary_food")) = True
    '    End If
    '
    '    If rs2("s7").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_mob")) = False
    '    Else
    '        .ColHidden(.ColIndex("Emp_Salary_mob")) = True
    '    End If
    '
    '    If rs2("s8").value = True Then
    '        .ColHidden(.ColIndex("Emp_Salary_mang")) = False
    '    Else
    ''        .ColHidden(.ColIndex("Emp_Salary_mang")) = True
     '   End If
     '
     '   If rs2("s9").value = True Then
     '       .ColHidden(.ColIndex("Emp_Salary_others")) = False
     '   Else
     '       .ColHidden(.ColIndex("Emp_Salary_others")) = True
     '   End If
     '
     '   If rs2("s10").value = True Then
     '       .ColHidden(.ColIndex("OverTimePrice")) = False
     '   Else
     '       .ColHidden(.ColIndex("OverTimePrice")) = True
     '   End If
     ''
      '  If rs2("s11").value = True Then
      '      .ColHidden(.ColIndex("Mokafea")) = False
      '  Else
      '      .ColHidden(.ColIndex("Mokafea")) = True
      ''  End If
       '
       ' If rs2("s12").value = True Then
       '     .ColHidden(.ColIndex("SalesCom")) = False
       ' Else
       '     .ColHidden(.ColIndex("SalesCom")) = True
       ' End If
       ''
        'If rs2("s13").value = True Then
        '    .ColHidden(.ColIndex("total1")) = False
        'Else
        '    .ColHidden(.ColIndex("total1")) = True
        'End If
        '
        'If rs2("s14").value = True Then
        ''    .ColHidden(.ColIndex("TotalAdvance")) = False
        'Else
         '   .ColHidden(.ColIndex("TotalAdvance")) = True
        'End If
         '
        'if rs2("s15").value = True Then
         '   .ColHidden(.ColIndex("TotalDiscount")) = False
        'Else
         '   .ColHidden(.ColIndex("TotalDiscount")) = True
        'End If
         '
        'If rs2("s16").value = True Then
        '    .ColHidden(.ColIndex("total2")) = False
        'Else
        '    .ColHidden(.ColIndex("total2")) = True
        'End If
                 
        'If rs2("s17").value = True Then
        '    .ColHidden(.ColIndex("EmpTotalNet")) = False
        'Else
        '    .ColHidden(.ColIndex("EmpTotalNet")) = True
        'End If
                  
        'If rs2("s18").value = True Then
        '    .ColHidden(.ColIndex("sgn")) = False
        'Else
        '    .ColHidden(.ColIndex("sgn")) = True
        'End If
     '
    'End With

'End Function

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub









Private Sub DcbPeriodsID_Change()
ChaDate
End Sub

Private Sub DcbPeriodsID_Click()
ChaDate
End Sub

Private Sub ENDDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         EndDateH.value = ToHijriDate(EndDate.value)
End If
End Sub

Private Sub ENDDATEH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    EndDate.value = ToGregorianDate(EndDateH.value)
    End If
End Sub

Private Sub Form_Load()

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture

    Dim My_SQL2 As String
  '  My_SQL2 = " select id,Project_name from projects order by Project_name"
  '  fill_combo DCPROJECT1, My_SQL2
  '  My_SQL2 = " select  oprid,des from projects_des"
  '  fill_combo Dcterm1, My_SQL2
  '  My_SQL2 = " select  id,name from terms_operations"
  '  fill_combo dcopr, My_SQL2
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If


sql = "SELECT DISTINCT KafelName, KafelName AS KafelNames"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (NOT (KafelName IS NULL)) "
fill_combo DcbKafelName, sql


    If SystemOptions.UserInterface = EnglishInterface Then
        StrSQL = "SELECT id,Namee from TblOffice"
    Else
        StrSQL = "SELECT id, Name from TblOffice"
    End If
    fill_combo cmbOffice, StrSQL
      'Dim ii As Integer
      
      
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TbVisa  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
Ele(5).Caption = "Visa Data"

lbl(7).Caption = "No."
lbl(2).Caption = "Request No."
lbl(0).Caption = "Visa Number"
lbl(6).Caption = "Count"
lbl(8).Caption = "Visa Date"
lbl(3).Caption = "Expiry"
lbl(4).Caption = "Expiration date"

With GRID2
    .TextMatrix(0, .ColIndex("Ser")) = "No."
    .TextMatrix(0, .ColIndex("Job")) = "Job"
    .TextMatrix(0, .ColIndex("Notional")) = "Nationality"
    .TextMatrix(0, .ColIndex("Place")) = "Arriving Port"
    .TextMatrix(0, .ColIndex("count")) = "Count"
End With

With Grid
    .TextMatrix(0, .ColIndex("Ser")) = "No."
    .TextMatrix(0, .ColIndex("Emp_Code")) = "Employee Code"
    .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
    .TextMatrix(0, .ColIndex("HododNo")) = "Borders No."
    .TextMatrix(0, .ColIndex("Notional")) = "Nationality"
    .TextMatrix(0, .ColIndex("Job")) = "Job"
    .TextMatrix(0, .ColIndex("price")) = "Amount"
End With

CmdRemove.Caption = "Delete Row"

Cmd(0).Caption = "New"
Cmd(1).Caption = "Edit"
Cmd(2).Caption = "Save"
Cmd(3).Caption = "Undo"
Cmd(4).Caption = "Delete"
Cmd(5).Caption = "Search"
Cmd(7).Caption = "Print"
Cmd(6).Caption = "Exit"







 End Sub
'Public Sub get_all_employee()
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
'    Dim J As Integer
'
'    Dim sql As String
'    Dim i As Integer
'
'    sql = "Select * from emp_all_details "
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then Exit Sub
'
'    With Grid
'
'        .Rows = 2
'        .Clear flexClearScrollable
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + 1
'            Rs3.MoveFirst
'
'            For i = 1 To Rs3.RecordCount
'                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
'                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
'                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
'                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
'                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
'
'                Rs3.MoveNext
'            Next i
'
'            .AutoSize 0, .Cols - 1, False
'        End If
'
'    End With
'
'    Rs3.Close
'
'End Sub
''Public Sub get_all_employee()
''    Dim Rs3 As ADODB.Recordset
''    Set Rs3 = New ADODB.Recordset
''    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
'    Dim J As Integer
''
 '   Dim sql As String
 '   Dim i As Integer
'
''    sql = "Select * from emp_all_details "
 '
 '   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '
 '   If Rs3.RecordCount = 0 Then Exit Sub
 ''
  '  With Grid
'
'        .Rows = 2
''        .Clear flexClearScrollable
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + 1
''            Rs3.MoveFirst
 '
 '           For i = 1 To Rs3.RecordCount
 '               .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
 ''
  '              .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
  '              .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
  '              .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
  '              .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
  '              .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
  '
  '              Rs3.MoveNext
  '          Next i
 '
 '           .AutoSize 0, .Cols - 1, False
 '       End If
'
'    End With
 
'    Rs3.Close

'End Sub

'Public Sub FillGridWithData()
'    Exit Sub
'
''    Dim i As Integer
 '   Dim rs As ADODB.Recordset
 '   Dim rs2 As ADODB.Recordset
 '   Dim LstDay As Date
 '   Dim FrstDay As Date
 '   Dim StrTxt As String
 '   Dim My_SQL As String
 ''   Dim StrWhere As String
  '  Dim StrGrp As String
  '  Dim IntMonth As Integer
  '  Dim IntYear As Integer
  ''  Dim Msg As String
'
'    On Error GoTo ErrTrap
'
'    Set rs = New ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'
'    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '   With Me.Grid
 '       .Rows = 2
 '       .Clear flexClearScrollable
'
'        If rs.RecordCount > 0 Then
'            .Rows = rs.RecordCount + 1
'            rs.MoveFirst
'
'            For i = 1 To .Rows - 1
'
'                .TextMatrix(i, .ColIndex("Ser")) = i
'                ',DepartmentID,project_id
''
 '               .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
 '
 '               .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
 '
 '               .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
 ''
  '              .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
  '
  '              .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
  '
  '              .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
  '
  '              .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
  ''
   '             .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
   '
   '             '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
   '              "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
   '
   '             '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
   '             '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
   '
   '             .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
   '
   '             rs.MoveNext
   '
   '         Next
'
'            rs.Close
'        End If
'
'        .Rows = .Rows + 1
'        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
'        .IsSubtotal(.Rows - 1) = True
'        Dim SngTotal As Single
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
'        net_value = SngTotal
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
'        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
'        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
'        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
'        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
'        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
'
'        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
'        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
'        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
'        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
'        .AutoSize 0, .Cols - 1, False
'    End With
''
'ErrTrap:
'End Sub
 
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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
     
    
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs2 As ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
           Case "Emp_Name"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_id")) = StrAccountCode
                If val(.TextMatrix(Row, .ColIndex("Emp_id"))) <> 0 Then
                Set rs2 = New ADODB.Recordset
                StrSQL = "Select Fullcode from TblEmployee where Emp_ID=" & val(.TextMatrix(Row, .ColIndex("Emp_id"))) & ""
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
                Else
                .TextMatrix(Row, .ColIndex("Emp_Code")) = ""
                End If
                End If
               Case "Emp_Code"
                If (.TextMatrix(Row, .ColIndex("Emp_Code"))) <> "" Then
                Set rs2 = New ADODB.Recordset
                StrSQL = "Select * from TblEmployee where Fullcode='" & (.TextMatrix(Row, .ColIndex("Emp_Code"))) & "'"
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs2("Emp_ID").value), 0, rs2("Emp_ID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
                Else
                .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("Emp_id")) = 0
                End If
                End If
                
            Case "Notional"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("NotionalID"), False, True)
                .TextMatrix(Row, .ColIndex("NotionalID")) = StrAccountCode
         Case "Office"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("OfficeID"), False, True)
                .TextMatrix(Row, .ColIndex("OfficeID")) = StrAccountCode
            Case "Job"

                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("JobID"), False, True)
                .TextMatrix(Row, .ColIndex("JobID")) = StrAccountCode
             Case "City"

                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CityID"), False, True)
                .TextMatrix(Row, .ColIndex("CityID")) = StrAccountCode
                
    End Select
   
       ' If Row = .Rows - 1 Then
    '
    '        .Rows = .Rows + 1
    '    End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid1()


    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    lbl(5).Caption = 0
'
    With Me.GRID2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("count")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                lbl(5).Caption = val(lbl(5).Caption) + val(.TextMatrix(i, .ColIndex("count")))

            End If

        Next i
   
    End With
    End Sub
    Private Sub Fill()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    Dim j As Integer
    If val(lbl(5).Caption) = 0 Then
    Exit Sub
    End If
    If Me.TxtModFlg.Text = "N" Then
      Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
      End If
Dim sumx As Long
sumx = 1
    IntCounter = 0
    
     With Me.GRID2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("count")) <> "" Then
              sumx = sumx + val(.TextMatrix(i, .ColIndex("count")))
      Grid.Rows = sumx
            For j = 1 To val(.TextMatrix(i, .ColIndex("count")))
             IntCounter = IntCounter + 1
        If val(Grid.TextMatrix(IntCounter, Grid.ColIndex("Emp_id"))) = 0 Then
            Grid.TextMatrix(IntCounter, Grid.ColIndex("CityID")) = val(GRID2.TextMatrix(i, GRID2.ColIndex("CityID")))
              Grid.TextMatrix(IntCounter, Grid.ColIndex("City")) = GRID2.TextMatrix(i, GRID2.ColIndex("City"))
               Grid.TextMatrix(IntCounter, Grid.ColIndex("Notional")) = GRID2.TextMatrix(i, GRID2.ColIndex("Notional"))
                Grid.TextMatrix(IntCounter, Grid.ColIndex("NotionalID")) = GRID2.TextMatrix(i, GRID2.ColIndex("NotionalID"))
                Grid.TextMatrix(IntCounter, Grid.ColIndex("JobID")) = GRID2.TextMatrix(i, GRID2.ColIndex("JobID"))
                Grid.TextMatrix(IntCounter, Grid.ColIndex("Job")) = GRID2.TextMatrix(i, GRID2.ColIndex("Job"))
                Grid.TextMatrix(IntCounter, Grid.ColIndex("OfficeID")) = GRID2.TextMatrix(i, GRID2.ColIndex("OfficeID"))
                Grid.TextMatrix(IntCounter, Grid.ColIndex("Office")) = GRID2.TextMatrix(i, GRID2.ColIndex("Office"))
                
               End If
            Next
               
                
  
            End If

        Next i
   
    End With
ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    IntCounter = 0
     With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Notional")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        Select Case .ColKey(Col)
            Case "price"
                .ComboList = ""
            Case "Emp_Code"
                .ComboList = ""

            Case "Job"
                'Cancel = True
        
            Case "count"
                .ComboList = ""
        
            Case "HododNo"
                .ComboList = ""
 Case "Price"
                .ComboList = ""
            
        End Select

    End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Me.Grid

        Select Case .ColKey(Col)
            Case "Emp_Name"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     Emp_ID , Emp_Name from  TblEmployee "
                  Else
                     StrSQL = " SELECT     Emp_ID ,  Emp_Namee from  TblEmployee "
                 End If
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = .BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
    Case "Office"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT id,Name from TblOffice"
                  Else
                     StrSQL = " SELECT id,Namee from TblOffice"
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "nameee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

            
Case "Notional"
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name from  Nationality where( name <> '”⁄ÊœÌ' and name <> '”⁄ÊœÏ')"
                  Else
                     StrSQL = " SELECT     id, upper(namee)as nameee from  Nationality  where namee <>'SAUDI'"
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "nameee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

 
           Case "Job"
        
 If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     JobTypeID, JobTypeName from  TblEmpJobsTypes "
                  Else
                     StrSQL = " SELECT     JobTypeID, JobTypeNamee from  TblEmpJobsTypes "
                End If
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "JobTypeName", "JobTypeID")
                Else
                    StrComboList = .BuildComboList(rs, "JobTypeNamee", "JobTypeID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

         Case "City"

                    StrSQL = " SELECT     GovernmentID, GovernmentName from  TblCountriesGovernments "
               
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = .BuildComboList(rs, "GovernmentName", "GovernmentID")
                        If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
        End Select

    End With

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

MySQL = " SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID,TblOffice.Name as OfficeName,TblOffice.Namee as OfficeNamee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
MySQL = MySQL & "                      dbo.TbVisaDeti.NotionalID,TbVisaDeti.OfficeId, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
MySQL = MySQL & "                      dbo.TbVisa.OrderNo, dbo.TbVisa.VisaNo, dbo.TbVisa.Priod, dbo.TbVisa.DMYPriod, dbo.TbVisa.StarDate, dbo.TbVisa.StarDateH, dbo.TbVisa.EndDate,"
MySQL = MySQL & "                      dbo.TbVisa.EndDateH, dbo.TbVisa.ID AS IDM, dbo.TbVisaDeti.Place, dbo.TbVisaDeti.Type, dbo.TbVisaDeti.[count], dbo.TbVisaDeti.Price"
MySQL = MySQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TbVisa LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TbVisaDeti ON dbo.TbVisa.ID = dbo.TbVisaDeti.VisaID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TbVisaDeti.CityID = dbo.TblCountriesGovernments.GovernmentID ON"
MySQL = MySQL & "                      dbo.Nationality.id = dbo.TbVisaDeti.NotionalID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TbVisaDeti.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "                       LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblOffice ON dbo.TbVisaDeti.OfficeId = dbo.TblOffice.ID"
MySQL = MySQL & " Where (dbo.TbVisa.id = " & val(xptxtid.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVisaData.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVisaDataE.rpt"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
    
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
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
      GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 2
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
  If Lngid <> 0 Then
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    Me.xptxtid.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    Me.TxtOrder.Text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
    Me.TxtVisa.Text = IIf(IsNull(rs("VisaNo").value), "", rs("VisaNo").value)
    StarDate.value = IIf(IsNull(rs("StarDate").value), Date, rs("StarDate").value)
    StarDateH.value = IIf(IsNull(rs("StarDateH").value), "", rs("StarDateH").value)
    
    txtArriveDate.value = IIf(IsNull(rs("ArriveDate").value), Date, rs("ArriveDate").value)
    txtArriveDateH.value = IIf(IsNull(rs("ArriveDateH").value), "", rs("ArriveDateH").value)


    EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    EndDateH.value = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
    Me.TxtPeriods.Text = IIf(IsNull(rs("Priod").value), 0, rs("Priod").value)
    cmbOffice.BoundText = IIf(IsNull(rs("OfficeId").value), 0, rs("OfficeId").value)
    Me.DcbPeriodsID.ListIndex = val(IIf(IsNull(rs("DMYPriod").value), -1, rs("DMYPriod").value))
    
    txtKafelID.Text = IIf(IsNull(rs("KafelID").value), "", Trim(rs("KafelID").value))
    Me.DcbKafelName.Text = IIf(IsNull(rs("KafelName").value), "", Trim(rs("KafelName").value))
    
    txtkafeltel.Text = IIf(IsNull(rs("kafeltel").value), "", Trim(rs("kafeltel").value))
    
    txtkafeladd.Text = IIf(IsNull(rs("kafeladd").value), "", Trim(rs("kafeladd").value))

'kafeladd,kafeltel,KafelName,KafelID
   ' MsgBox dcopr.BoundText


StrSQL = " SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID,TbVisaDeti.OfficeId,TbVisaDeti.remarks ,TblOffice.Name as OfficeName,TblOffice.Namee as OfficeNamee, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.[count] , dbo.TbVisaDeti.Type, dbo.TbVisaDeti.Place , dbo.TbVisaDeti.Price"
StrSQL = StrSQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCountriesGovernments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TbVisaDeti ON dbo.TblCountriesGovernments.GovernmentID = dbo.TbVisaDeti.CityID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Nationality ON dbo.TbVisaDeti.NotionalID = dbo.Nationality.id ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TbVisaDeti.JobID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "                       lEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblOffice ON dbo.TbVisaDeti.OfficeId = dbo.TblOffice.ID"
StrSQL = StrSQL & "  Where (dbo.TbVisaDeti.VisaID = " & val(xptxtid.Text) & ") And (dbo.TbVisaDeti.Type = 0)"
'StrSQL = StrSQL & " Where (dbo.TbVisaDeti.VisaID = " & val(Me.xptxtid.text) & ")"

   ' StrSQL = "select * from opr_employee_details where pk_id=" & Me.xptxtid.text
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                .TextMatrix(i, .ColIndex("OfficeId")) = IIf(IsNull(RsDev("OfficeId").value), "", RsDev("OfficeId").value)
                 .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeName").value), "", RsDev("OfficeName").value)
               
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeNamee").value), "", RsDev("OfficeNamee").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
                .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(RsDev("HododNo").value), "", RsDev("HododNo").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(RsDev("GovernmentName").value), "", RsDev("GovernmentName").value)
                .TextMatrix(i, .ColIndex("NotionalID")) = IIf(IsNull(RsDev("NotionalID").value), "", RsDev("NotionalID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
                .TextMatrix(i, .ColIndex("CityID")) = IIf(IsNull(RsDev("CityID").value), "", RsDev("CityID").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
                .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("remarks").value), "", RsDev("remarks").value)
            
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 StrSQL = " SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID ,TbVisaDeti.OfficeId ,TbVisaDeti.remarks, TblOffice.Name as OfficeName,TblOffice.Namee as OfficeNamee,dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.[count] , dbo.TbVisaDeti.Type, dbo.TbVisaDeti.Place"
StrSQL = StrSQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCountriesGovernments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TbVisaDeti ON dbo.TblCountriesGovernments.GovernmentID = dbo.TbVisaDeti.CityID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Nationality ON dbo.TbVisaDeti.NotionalID = dbo.Nationality.id ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TbVisaDeti.JobID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "                       Left OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblOffice ON dbo.TbVisaDeti.OfficeId = dbo.TblOffice.ID"
StrSQL = StrSQL & "  Where (dbo.TbVisaDeti.VisaID = " & val(xptxtid.Text) & ") And (dbo.TbVisaDeti.Type = 1)"
'StrSQL = StrSQL & " Where (dbo.TbVisaDeti.VisaID = " & val(Me.xptxtid.text) & ")"

   ' StrSQL = "select * from opr_employee_details where pk_id=" & Me.xptxtid.text
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GRID2
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeName").value), "", RsDev("OfficeName").value)
            '    .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            '
               ' .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
               ' .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeName").value), "", RsDev("OfficeName").value)
                Else
               ' .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                .TextMatrix(i, .ColIndex("Office")) = IIf(IsNull(RsDev("OfficeNamee").value), "", RsDev("OfficeNamee").value)
                End If
                .TextMatrix(i, .ColIndex("OfficeId")) = IIf(IsNull(RsDev("OfficeId").value), "", RsDev("OfficeId").value)
                .TextMatrix(i, .ColIndex("Place")) = IIf(IsNull(RsDev("Place").value), "", RsDev("Place").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(RsDev("GovernmentName").value), "", RsDev("GovernmentName").value)
                .TextMatrix(i, .ColIndex("NotionalID")) = IIf(IsNull(RsDev("NotionalID").value), "", RsDev("NotionalID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
                .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev("count").value), "", RsDev("count").value)
                .TextMatrix(i, .ColIndex("CityID")) = IIf(IsNull(RsDev("CityID").value), "", RsDev("CityID").value)
            
                 .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("remarks").value), "", RsDev("remarks").value)
                RsDev.MoveNext
            Next i
 
        End With

    End If
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 


Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
                          
     
    
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    If val(GRID2.TextMatrix(Row, GRID2.ColIndex("OfficeID"))) = 0 Then
        GRID2.TextMatrix(Row, GRID2.ColIndex("OfficeID")) = cmbOffice.BoundText
        GRID2.TextMatrix(Row, GRID2.ColIndex("Office")) = cmbOffice.Text
    End If
    With GRID2

        Select Case .ColKey(Col)
 
            Case "Notional"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("NotionalID"), False, True)
                .TextMatrix(Row, .ColIndex("NotionalID")) = StrAccountCode
        Case "Office"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("OfficeID"), False, True)
                .TextMatrix(Row, .ColIndex("OfficeID")) = StrAccountCode
        
        
            Case "Job"

                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("JobID"), False, True)
                .TextMatrix(Row, .ColIndex("JobID")) = StrAccountCode
             Case "City"

                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CityID"), False, True)
                .TextMatrix(Row, .ColIndex("CityID")) = StrAccountCode
                
    End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        
    End With
    ReLineGrid1
End Sub

Private Sub GRID2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   With GRID2

        Select Case .ColKey(Col)

            
        
            Case "count"
                .ComboList = ""
        
            Case "Place"
                .ComboList = ""

            
        End Select

    End With
End Sub

Private Sub grid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Me.GRID2

        Select Case .ColKey(Col)
 
Case "Notional"
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name from  Nationality where( name <> '”⁄ÊœÌ' and name <> '”⁄ÊœÏ')"
                  Else
                     StrSQL = " SELECT     id, upper(namee)as nameee from  Nationality  where namee <>'SAUDI'"
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "nameee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
Case "Office"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT id,Name from TblOffice"
                  Else
                     StrSQL = " SELECT id,Namee from TblOffice"
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "nameee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing


 
           Case "Job"
        
 If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     JobTypeID, JobTypeName from  TblEmpJobsTypes "
                  Else
                     StrSQL = " SELECT     JobTypeID, JobTypeNamee from  TblEmpJobsTypes "
                End If
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "JobTypeName", "JobTypeID")
                Else
                    StrComboList = .BuildComboList(rs, "JobTypeNamee", "JobTypeID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing

         Case "City"

                    StrSQL = " SELECT     GovernmentID, GovernmentName from  TblCountriesGovernments "
               
                      Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = .BuildComboList(rs, "GovernmentName", "GovernmentID")
                        If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

               .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
        End Select

    End With
End Sub

Private Sub lbl_Change(Index As Integer)
'Grid.Rows = val(lbl(5).Caption) + Grid.FixedRows
Fill
End Sub

Private Sub StarDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         StarDateH.value = ToHijriDate(StarDate.value)
End If
End Sub

Private Sub StarDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    StarDate.value = ToGregorianDate(StarDateH.value)
    End If
End Sub



Private Sub txtArriveDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
        txtArriveDateH.value = ToHijriDate(txtArriveDate.value)
End If
End Sub

Private Sub txtArriveDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    txtArriveDate.value = ToGregorianDate(txtArriveDateH.value)
    End If
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtPeriods_Change()
ChaDate
End Sub
Sub ChaDate()
Dim DateInterval As String
Dim DateNumber As Integer
If val(Me.DcbPeriodsID.ListIndex) <> -1 Then
If val(TxtPeriods.Text) <> 0 Then
If Me.TxtModFlg.Text <> "R" Then
DateNumber = val(TxtPeriods.Text)
    If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
    End If
EndDate.value = Me.StarDate.value
                EndDate.value = DateAdd(DateInterval, DateNumber, EndDate)
                EndDateH.value = ToHijriDate(EndDate)
              
     End If
     End If
 End If
 End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

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
