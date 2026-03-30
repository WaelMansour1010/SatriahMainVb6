VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmEmpSalary3 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘…  Œ’Ì’ ⁄„· «·„ÊŸ›Ì‰ ›Ì „‘—Ê⁄ „⁄Ì‰"
   ClientHeight    =   8505
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14445
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
   Icon            =   "FrmEmpSalary3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   14445
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8505
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14445
      _cx             =   25479
      _cy             =   15002
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
         Height          =   7440
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14385
         _cx             =   25374
         _cy             =   13123
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
         Caption         =   "«·»Ì«‰« |«·«⁄ „«œ« | ›«’Ì· «·«ÃÊ—"
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
            Height          =   7020
            Left            =   15330
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   45
            Width           =   14295
            _cx             =   25215
            _cy             =   12383
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   7020
               Left            =   0
               TabIndex        =   87
               Top             =   0
               Width           =   14295
               _cx             =   25215
               _cy             =   12382
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary3.frx":038A
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
               ExplorerBar     =   1
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   14295
            _cx             =   25215
            _cy             =   12383
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
               Height          =   795
               Index           =   5
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   0
               Width           =   14340
               _cx             =   25294
               _cy             =   1402
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
               Picture         =   "FrmEmpSalary3.frx":059E
               Caption         =   "‘«‘…  Œ’Ì’ ⁄„· «·„ÊŸ›Ì‰   "
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
               Begin VB.TextBox toid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   36
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
                  ButtonImage     =   "FrmEmpSalary3.frx":1278
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
                  TabIndex        =   37
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
                  ButtonImage     =   "FrmEmpSalary3.frx":1612
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
                  TabIndex        =   38
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
                  ButtonImage     =   "FrmEmpSalary3.frx":19AC
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
                  TabIndex        =   39
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
                  ButtonImage     =   "FrmEmpSalary3.frx":1D46
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
               Height          =   7170
               Index           =   1
               Left            =   -135
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   -135
               Width           =   15240
               _cx             =   26882
               _cy             =   12647
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
               Begin VB.CheckBox ChckAutoEmp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " Õ„Ì· «·„ÊŸ›Ì‰ „‰ «·„‘—Ê⁄ "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   9735
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   1755
                  Width           =   2310
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
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
                  Left            =   6045
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   1005
                  Width           =   1575
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   1230
                  Left            =   135
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   2070
                  Width           =   14340
                  _cx             =   25294
                  _cy             =   2170
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
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
                  Begin VB.TextBox Text2 
                     Alignment       =   1  'Right Justify
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
                     Left            =   12045
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   750
                     Width           =   720
                  End
                  Begin VB.OptionButton Option2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«Œ Ì«— «·„ÊŸ›Ì‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   5055
                     TabIndex        =   72
                     Top             =   750
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   6945
                     TabIndex        =   71
                     Top             =   750
                     Width           =   1485
                  End
                  Begin VB.TextBox TxtSearchCode 
                     Alignment       =   2  'Center
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
                     Left            =   4020
                     TabIndex        =   70
                     Top             =   750
                     Width           =   720
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   405
                     Index           =   20
                     Left            =   135
                     TabIndex        =   73
                     Top             =   615
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   714
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
                     ButtonImage     =   "FrmEmpSalary3.frx":20E0
                     DrawFocusRectangle=   0   'False
                  End
                  Begin XtremeSuiteControls.CheckBox SelectBranch 
                     Height          =   225
                     Left            =   12945
                     TabIndex        =   74
                     Top             =   270
                     Width           =   1125
                     _Version        =   786432
                     _ExtentX        =   1984
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "›—⁄ „Õœœ"
                     BackColor       =   -2147483633
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcpDept1 
                     Height          =   315
                     Left            =   4425
                     TabIndex        =   75
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   270
                     Width           =   2925
                     _ExtentX        =   5159
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin XtremeSuiteControls.CheckBox SelectDept 
                     Height          =   225
                     Left            =   7485
                     TabIndex        =   76
                     Top             =   270
                     Width           =   1080
                     _Version        =   786432
                     _ExtentX        =   1905
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "«œ«—… „Õœœ…"
                     BackColor       =   -2147483633
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCEmployee 
                     Height          =   315
                     Left            =   945
                     TabIndex        =   77
                     Top             =   750
                     Width           =   3030
                     _ExtentX        =   5345
                     _ExtentY        =   556
                     _Version        =   393216
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
                     Height          =   315
                     Left            =   8655
                     TabIndex        =   78
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   270
                     Width           =   4110
                     _ExtentX        =   7250
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin MSDataListLib.DataCombo DcbProject1 
                     Height          =   315
                     Left            =   8655
                     TabIndex        =   79
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   750
                     Width           =   3300
                     _ExtentX        =   5821
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin XtremeSuiteControls.CheckBox SelectProject 
                     Height          =   225
                     Left            =   12945
                     TabIndex        =   80
                     Top             =   750
                     Width           =   1125
                     _Version        =   786432
                     _ExtentX        =   1984
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "„‘—Ê⁄ „Õœœ"
                     BackColor       =   -2147483633
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbTeam 
                     Height          =   315
                     Left            =   135
                     TabIndex        =   81
                     Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                     Top             =   270
                     Width           =   2925
                     _ExtentX        =   5159
                     _ExtentY        =   556
                     _Version        =   393216
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
                  Begin XtremeSuiteControls.CheckBox SelctTeam 
                     Height          =   225
                     Left            =   3150
                     TabIndex        =   82
                     Top             =   270
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "›—Ìﬁ „Õœœ"
                     BackColor       =   -2147483633
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «Œ — «·„ÊŸ›"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Index           =   18
                     Left            =   4740
                     TabIndex        =   83
                     Top             =   1230
                     Visible         =   0   'False
                     Width           =   1260
                  End
               End
               Begin XtremeSuiteControls.RadioButton RdTypePay 
                  Height          =   285
                  Index           =   0
                  Left            =   4920
                  TabIndex        =   62
                  Top             =   2070
                  Visible         =   0   'False
                  Width           =   1575
                  _Version        =   786432
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  _StockProps     =   79
                  Caption         =   "ÿ»ﬁ« ·⁄ﬁœ «·„ÊŸ›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox txtDays 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   12270
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Text            =   "0"
                  Top             =   1755
                  Width           =   855
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   8655
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1245
                  Visible         =   0   'False
                  Width           =   990
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Text            =   "0"
                  Top             =   750
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox xptxtid 
                  Alignment       =   1  'Right Justify
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
                  Left            =   9735
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   1005
                  Width           =   3345
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«— ﬂ· «·„ÊŸ›Ì‰"
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
                  Left            =   9060
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   1005
                  Visible         =   0   'False
                  Width           =   1710
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   9555
                  Width           =   2175
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   5865
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   495
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3720
                  Left            =   135
                  TabIndex        =   7
                  Top             =   3315
                  Width           =   14250
                  _cx             =   25135
                  _cy             =   6562
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
                  Cols            =   36
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmEmpSalary3.frx":247A
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
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   315
                  Left            =   11865
                  TabIndex        =   12
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   142802945
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   135
                  TabIndex        =   13
                  Top             =   1755
                  Width           =   7485
                  _ExtentX        =   13203
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo dcproject 
                  Height          =   315
                  Left            =   135
                  TabIndex        =   16
                  Top             =   1005
                  Width           =   5820
                  _ExtentX        =   10266
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo Dcterm 
                  Height          =   315
                  Left            =   135
                  TabIndex        =   31
                  Top             =   1380
                  Width           =   7485
                  _ExtentX        =   13203
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSComCtl2.DTPicker end_date 
                  Height          =   330
                  Left            =   9735
                  TabIndex        =   42
                  Top             =   2700
                  Visible         =   0   'False
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   582
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   142802945
                  CurrentDate     =   38784
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   585
                  Index           =   3
                  Left            =   9735
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   4470
                  _cx             =   7885
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
                  Caption         =   " Õœœ «·› —…"
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
                  Begin VB.ComboBox CmbMonth 
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
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   180
                     Width           =   1485
                  End
                  Begin VB.ComboBox CboYear 
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
                     ItemData        =   "FrmEmpSalary3.frx":29D3
                     Left            =   250
                     List            =   "FrmEmpSalary3.frx":29D5
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   45
                     Top             =   180
                     Width           =   1005
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   6
                     Left            =   1425
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   270
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   3
                     Left            =   2955
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   180
                     Width           =   1020
                  End
               End
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   330
                  Left            =   11640
                  TabIndex        =   50
                  Top             =   1380
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   582
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  Format          =   142802945
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   330
                  Left            =   9735
                  TabIndex        =   53
                  Top             =   1380
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   582
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CheckBox        =   -1  'True
                  Format          =   142802945
                  CurrentDate     =   38784
               End
               Begin XtremeSuiteControls.RadioButton RdTypePay 
                  Height          =   270
                  Index           =   1
                  Left            =   3240
                  TabIndex        =   63
                  Top             =   2205
                  Visible         =   0   'False
                  Width           =   1590
                  _Version        =   786432
                  _ExtentX        =   2805
                  _ExtentY        =   476
                  _StockProps     =   79
                  Caption         =   "ÿ»ﬁ« ··„‘—Ê⁄"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   345
                  Index           =   10
                  Left            =   9015
                  TabIndex        =   89
                  Top             =   1755
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmEmpSalary3.frx":29D7
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «Õ ”«»  „⁄œ· «·ÌÊ„"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   12
                  Left            =   7575
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   2205
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«Ì«„"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   11
                  Left            =   13440
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   1755
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   10
                  Left            =   5190
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   2070
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   9
                  Left            =   7665
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   2070
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   11040
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   1380
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»‰œ"
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
                  Index           =   0
                  Left            =   7710
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   1380
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·„‘—Ê⁄"
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
                  Index           =   5
                  Left            =   7710
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   1005
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·Ì…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   4
                  Left            =   7710
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   1755
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   8
                  Left            =   12900
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   1380
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· Œ’Ì’"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   7
                  Left            =   12900
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1005
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   390
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   1005
                  Width           =   855
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
               Height          =   315
               Index           =   1
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   105
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7020
            Left            =   15030
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   45
            Width           =   14295
            _cx             =   25215
            _cy             =   12383
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   3705
               Left            =   135
               TabIndex        =   58
               Tag             =   "1"
               Top             =   240
               Width           =   11415
               _cx             =   20135
               _cy             =   6535
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary3.frx":2D71
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
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   240
               Left            =   6585
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   4155
               Width           =   4830
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   990
         Left            =   30
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   7485
         Width           =   14385
         _cx             =   25374
         _cy             =   1746
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   345
            Left            =   11880
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   105
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   609
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
            ButtonImage     =   "FrmEmpSalary3.frx":2EB4
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   315
            Left            =   12765
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
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
            ButtonImage     =   "FrmEmpSalary3.frx":324E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   300
            Left            =   13965
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   165
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   529
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
            ButtonImage     =   "FrmEmpSalary3.frx":35E8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   525
            Index           =   0
            Left            =   9720
            TabIndex        =   24
            Top             =   510
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   926
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
            Height          =   525
            Index           =   1
            Left            =   8520
            TabIndex        =   25
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   926
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
            Height          =   525
            Index           =   2
            Left            =   7680
            TabIndex        =   26
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   926
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
            Height          =   525
            Index           =   3
            Left            =   6675
            TabIndex        =   27
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   926
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
            Height          =   525
            Index           =   4
            Left            =   5640
            TabIndex        =   28
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   926
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
            Height          =   525
            Index           =   6
            Left            =   480
            TabIndex        =   29
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   926
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
            Height          =   525
            Index           =   5
            Left            =   4710
            TabIndex        =   30
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   926
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
            Height          =   525
            Index           =   7
            Left            =   3720
            TabIndex        =   49
            Top             =   510
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   926
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   8
            Left            =   2520
            TabIndex        =   56
            Top             =   495
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·Â"
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
         Begin ImpulseButton.ISButton Accredit 
            Height          =   510
            Left            =   5040
            TabIndex        =   60
            Top             =   0
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
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
            Height          =   510
            Index           =   9
            Left            =   1440
            TabIndex        =   64
            Top             =   495
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì"
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
         Begin ImpulseButton.ISButton CmdRemove 
            Height          =   285
            Left            =   8985
            TabIndex        =   90
            Top             =   90
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
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
            ButtonImage     =   "FrmEmpSalary3.frx":3982
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRemoveAll 
            Height          =   285
            Left            =   7320
            TabIndex        =   91
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmEmpSalary3.frx":3F1C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label XPTxtCount 
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
            Height          =   330
            Left            =   240
            TabIndex        =   68
            Top             =   135
            Width           =   855
         End
         Begin VB.Label XPTxtCurrent 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   67
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
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
            Left            =   3660
            TabIndex        =   66
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
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
            Left            =   1500
            TabIndex        =   65
            Top             =   135
            Width           =   1065
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
            Height          =   225
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   1740
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
            Height          =   225
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   255
            Width           =   1515
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
      ButtonImage     =   "FrmEmpSalary3.frx":44B6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalary3"
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
Dim rsDummy As ADODB.Recordset
Dim s As String
Public LongRow As Long
Public LngCol As Long
Public LngRow As Long
Dim mToDate As String
Dim NoDay As Long
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long



Private Sub ChkDetails_Click()
    FillGridWithData
End Sub
Function print_report2(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     TOP 100 PERCENT dbo.ProJectMofrdSalar.ID, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, projects_1.Project_name, projects_1.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
MySQL = MySQL & "                       dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, dbo.ProJectMofrdSalar.Total, dbo.ProJectMofrdSalar.NoDay,"
MySQL = MySQL & "                       dbo.ProJectMofrdSalar.YearID, dbo.ProJectMofrdSalar.MonthID, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
MySQL = MySQL & "                       dbo.TblEmployee.ContractID, dbo.TblEmployee.GroupID, projects_1.Salary_account, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.ProJectMofrdSalar.pk_id, dbo.opr_Employee.opr_type, dbo.opr_Employee.Years,"
MySQL = MySQL & "                       dbo.opr_Employee.Months, dbo.opr_Employee.FromDate, dbo.opr_Employee.ToDate, dbo.opr_Employee.Project_id, projects_2.Project_name AS Project_nameH,"
MySQL = MySQL & "                       projects_2.Project_nameE AS Project_nameHE, dbo.opr_Employee.PandID, dbo.projects_des.project_no, dbo.projects_des.des, dbo.opr_Employee.OpraID,"
MySQL = MySQL & "                       dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.ProJectMofrdSalar.TypeSalary, projects_2.Fullcode AS ProjectFullcodeDet,"
MySQL = MySQL & "                       projects_1.Fullcode AS ProjectFullcode, projects_2.EmpId AS MangerEmpId, TblEmployee_1.Emp_Name AS MangerEmp_Name,"
MySQL = MySQL & "                       TblEmployee_1.Fullcode AS MangerFullcode, TblEmployee_1.Emp_Namee AS MangerEmp_NameE, projects_2.EmpId1,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Name AS SuperEmp_Name, TblEmployee_2.Fullcode AS SuperFullcode, TblEmployee_2.Emp_Namee AS SuperEmp_NameE"
MySQL = MySQL & "  FROM         dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.projects projects_2 ON TblEmployee_2.Emp_ID = projects_2.EmpId1 LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee TblEmployee_1 ON projects_2.EmpId = TblEmployee_1.Emp_ID FULL OUTER JOIN"
MySQL = MySQL & "                       dbo.projects projects_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.opr_Employee LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblProcessDEF ON dbo.opr_Employee.OpraID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.projects_des ON dbo.opr_Employee.PandID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 ON"
MySQL = MySQL & "                       projects_1.id = dbo.opr_Employee.Project_id FULL OUTER JOIN"
MySQL = MySQL & "                       dbo.mofrdat RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.ProJectMofrdSalar ON dbo.mofrdat.mofrad_code = dbo.ProJectMofrdSalar.MofrdID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID ON dbo.opr_Employee.id = dbo.ProJectMofrdSalar.pk_id ON"
MySQL = MySQL & "                       projects_2.ID = dbo.ProJectMofrdSalar.ProjID"
MySQL = MySQL & "  Where (dbo.ProJectMofrdSalar.pk_id =" & XPTxtID.Text & ") And (dbo.opr_Employee.opr_type = 0) "
MySQL = MySQL & "  ORDER BY dbo.ProJectMofrdSalar.ProjID, dbo.ProJectMofrdSalar.EmpID"
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â·  —Ìœ ⁄„· Ã—Ê» ⁄·Ï „” ÊÏ «·„‘—Ê⁄"
Else
Msg = "Do you want to Group at the project level"
End If
 If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects21.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects21e.rpt"
        End If
   Else
         If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects2.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects2e.rpt"
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      Else
      Msg = "No Found Data"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
Sub RetriveProjectSalar()
Dim rs2 As ADODB.Recordset
Dim i As Integer
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     TOP 100 PERCENT dbo.ProJectMofrdSalar.ID, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
sql = sql & "                      dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
sql = sql & "                      dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, dbo.ProJectMofrdSalar.Total, dbo.ProJectMofrdSalar.NoDay,"
sql = sql & "                      dbo.ProJectMofrdSalar.YearID, dbo.ProJectMofrdSalar.MonthID, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
sql = sql & "                      dbo.TblEmployee.ContractID , dbo.TblEmployee.GroupID, dbo.Projects.Salary_account ,dbo.ProJectMofrdSalar.pk_id ,dbo.ProJectMofrdSalar.TypeSalary"
sql = sql & " FROM         dbo.ProJectMofrdSalar LEFT OUTER JOIN"
sql = sql & "                      dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & "  Where (dbo.ProJectMofrdSalar.pk_id =" & XPTxtID.Text & ")"
sql = sql & "  ORDER BY dbo.ProJectMofrdSalar.ProjID ,dbo.ProJectMofrdSalar.EmpID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.Rows = 2
.Rows = .Rows + rs2.RecordCount
For i = .FixedRows To .Rows - 2
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Salary_account")) = IIf(IsNull(rs2("Salary_account").value), "", rs2("Salary_account").value)
.TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
If Not IsNull(rs2("TypeSalary").value) Then
If rs2("TypeSalary").value = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("TypeSalary")) = "—« » „‘—Ê⁄"
Else
.TextMatrix(i, .ColIndex("TypeSalary")) = "Salary Projects"
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("TypeSalary")) = "—« » ‘—ﬂ…"
Else
.TextMatrix(i, .ColIndex("TypeSalary")) = "Salary Company"
End If
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("TypeSalary")) = "—« » ‘—ﬂ…"
Else
.TextMatrix(i, .ColIndex("TypeSalary")) = "Salary Company"
End If
End If
.TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(rs2("ProjID").value), 0, rs2("ProjID").value)
.TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(rs2("MofrdID").value), 0, rs2("MofrdID").value)
.TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(rs2("Valuee").value), 0, rs2("Valuee").value)
.TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
.TextMatrix(i, .ColIndex("NoDay")) = IIf(IsNull(rs2("NoDay").value), 0, rs2("NoDay").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs2("Project_name").value), "", rs2("Project_name").value)
.TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs2("mofrad_name").value), "", rs2("mofrad_name").value)
Else
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs2("Project_nameE").value), "", rs2("Project_nameE").value)
.TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs2("mofrad_namee").value), "", rs2("mofrad_namee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Function print_report(Optional NoteSerial As String)
     Dim rs As ADODB.Recordset
   ' Set rs = New ADODB.Recordset
   ' rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = "SELECT     dbo.opr_Employee.id, dbo.opr_Employee.Start_date, dbo.opr_Employee.Auto, dbo.opr_Employee.RecordDate, dbo.opr_Employee.Years, dbo.opr_Employee.Months, "
MySQL = MySQL & "                      dbo.opr_Employee.FromDate, dbo.opr_Employee.ToDate, dbo.opr_Employee.Posted, dbo.opr_Employee.PostedDate, dbo.opr_Employee.Approved,"
MySQL = MySQL & "                      dbo.opr_Employee.TypePay, dbo.opr_Employee.SelectAll, dbo.opr_Employee.SelectEmp, dbo.opr_Employee.SelectProj1, dbo.opr_Employee.ProjectID,"
MySQL = MySQL & "                      dbo.opr_Employee.Project_id, projects_1.Fullcode, projects_1.Project_name, projects_1.Project_nameE, dbo.opr_Employee.PandID, projects_des_1.des,"
MySQL = MySQL & "                      dbo.opr_Employee.OpraID, TblProcessDEF_1.ProcessName, TblProcessDEF_1.ProcessNameE, TblEmployee_3.Emp_Name,"
MySQL = MySQL & "                      TblEmployee_1.Fullcode AS EmpFullcode, TblEmployee_1.Emp_Namee, dbo.opr_employee_details.Emp_id, dbo.opr_employee_details.Start_date AS Start_dateDet,"
MySQL = MySQL & "                      dbo.opr_employee_details.end_date, dbo.opr_employee_details.no_of_days, dbo.opr_employee_details.Ended,"
MySQL = MySQL & "                      dbo.opr_employee_details.Project_id AS Project_idDe, projects_1.Project_name AS Project_nameDet, projects_1.Fullcode AS ProhecFullcodeDet,"
MySQL = MySQL & "                      projects_1.Project_nameE AS Project_nameDetE, dbo.opr_employee_details.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName,"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.opr_employee_details.[Count], dbo.opr_employee_details.daysalary, dbo.opr_employee_details.Total,"
MySQL = MySQL & "                      dbo.opr_employee_details.toid, dbo.opr_employee_details.[interval], dbo.opr_employee_details.FromDate AS FromDateDet,"
MySQL = MySQL & "                      dbo.opr_employee_details.ToDate AS ToDateDet, dbo.opr_employee_details.ContProjSalar, dbo.opr_employee_details.NumEkama,"
MySQL = MySQL & "                      dbo.opr_employee_details.PandID AS PandIDDet, projects_des_1.des AS desDet, dbo.opr_employee_details.OperID,"
MySQL = MySQL & "                      TblProcessDEF_1.ProcessName AS ProcessNameDet, TblProcessDEF_1.ProcessNameE AS ProcessNameDetDet, projects_1.EmpId,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name AS ManagerEmp_Name, TblEmployee_1.Fullcode AS ManagerFullcode, TblEmployee_1.Emp_Namee AS ManagerEmp_NameE,"
MySQL = MySQL & "                      projects_1.EmpId1, TblEmployee_2.Emp_Name AS SuperEmp_Name, TblEmployee_2.Fullcode AS SuperFullcode, TblEmployee_2.Emp_Namee AS SuperEmp_NameE,"
MySQL = MySQL & "                      dbo.opr_employee_details.DepartmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL & "                      dbo.opr_employee_details.BranchID , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
MySQL = MySQL & " FROM         dbo.projects_des projects_des_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.opr_employee_details ON dbo.TblBranchesData.branch_id = dbo.opr_employee_details.BranchId LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.opr_employee_details.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessDEF TblProcessDEF_1 ON dbo.opr_employee_details.OperID = TblProcessDEF_1.TblProcessDEFID ON"
MySQL = MySQL & "                      projects_des_1.oprid = dbo.opr_employee_details.PandID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.opr_employee_details.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects projects_1 LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON projects_1.EmpId1 = TblEmployee_2.Emp_ID ON TblEmployee_1.Emp_ID = projects_1.EmpId ON"
MySQL = MySQL & "                      dbo.opr_employee_details.Project_id = projects_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_3 ON dbo.opr_employee_details.Emp_id = TblEmployee_3.Emp_ID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects projects_2 ON dbo.opr_Employee.Project_id = projects_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblProcessDEF TblProcessDEF_2 ON dbo.opr_Employee.OpraID = TblProcessDEF_2.TblProcessDEFID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects_des projects_des_2 ON dbo.opr_Employee.PandID = projects_des_2.oprid"
MySQL = MySQL & "   Where (dbo.opr_Employee.opr_type = 0) And (dbo.opr_Employee.id = " & val(XPTxtID.Text) & ")"

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpSalary3.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpSalary3E.rpt"
            
       End If
           

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
NoteSerial = MySQL
    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation3
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
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

Private Sub ALLButton1_Click()

End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
       Else
       MsgBox "No Create Branch", vbCritical
     End If
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    ''create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function Create_dev1()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    'StrAccountCode = Account_Code_dynamic
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    'create_report_data

    DoEvents

    Exit Function
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  Else
   MsgBox "Sory...error douring save data", vbExclamation
End If
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    DCproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    ''CmdOk_Click
End Sub





Private Sub Accredit_Click()
    Dim sql As String
    Dim BeginTrans As Boolean
    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
    'Cn.Execute sql

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If


   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
  Retrive (val(XPTxtID.Text))



End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function

Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID,dbo.TbllevelWorker.EmpID1, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
               Dim dcjopstatus As Integer
               Dim EmpID As Integer
               EmpID = GetempidFromUserid(RSApproval("EmpID").value)
               get_employee_information EmpID, , , , , , , , , , , , , , , , , , , , , , , , , dcjopstatus
               
               If dcjopstatus = 4 Then
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID1").value), Null, Rs1("EmpID1").value)
               End If
               
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtID.Text)
                  RSApproval("NoteSerial").value = val(XPTxtID.Text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function


Private Sub CboYear_Click()
  '  'CmdOk_Click
End Sub

Private Sub ChckAutoEmp_Click()
If ChckAutoEmp.value = vbChecked Then
C1Elastic2.Enabled = False
Cmd(10).Enabled = True
Else
C1Elastic2.Enabled = True
Cmd(10).Enabled = False
End If
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .Rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub

Private Sub CmbMonth_Click()
   ' 'CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub



Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub
Function GetTypeEmployee(Optional EmpID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select TypeEmp from TblEmployee where Emp_ID=" & EmpID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetTypeEmployee = IIf(IsNull(Rs3("TypeEmp").value), 0, Rs3("TypeEmp").value) + 1
Else
GetTypeEmployee = 0
End If
End Function
Private Sub Combo1_Click()
 
End Sub

Private Sub CmdRemoveAll_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    Grid.Rows = 1
            Me.Grid.Clear flexClearScrollable, flexClearEverything
            
 
End Sub

Private Sub DcbProject1_Change()
DcbProject1_Click (0)
End Sub

Private Sub DcbProject1_Click(Area As Integer)
Dim Fullcode As String
GetCodeIDProject val(DcbProject1.BoundText), Fullcode
Text2.Text = Fullcode
End Sub

Private Sub DcbProject1_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
               FrmProjectSearch.lblSearchtype.Caption = 31
               FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
Select Case .ColKey(Col)
 Case "FromDate"
        LngRow = Row
        LngCol = Col
        FrmDateOpProject.Index = 30
       
        Load FrmDateOpProject
        FrmDateOpProject.Index = 30
        FrmDateOpProject.show vbModal
  Case "ToDate"
        LngRow = Row
        LngCol = Col
        Dim Frm As New FrmDateOpProject
        Frm.Index = 31
         Frm.XPDtbBill.CheckBox = True
        Load Frm
        Frm.Index = 31
        Frm.show vbModal
End Select
End With
End Sub

Private Sub Option2_Click()
If Me.Option2.value = False Then
TxtSearchCode.Text = ""
TxtSearchCode.Enabled = False
DCEmployee.BoundText = 0
DCEmployee.Enabled = False
Else
DCEmployee.Enabled = True
TxtSearchCode.Enabled = True
End If
End Sub

Private Sub SelctTeam_Click()
If Me.SelctTeam.value = vbChecked Then
DcbTeam.Enabled = True
Else
DcbTeam.Enabled = False
DcpDept1.BoundText = ""
End If
End Sub

Private Sub SelectBranch_Click()
If Me.SelectBranch.value = vbChecked Then
DcbBranch1.Enabled = True
Else
DcbBranch1.BoundText = ""
DcbBranch1.Enabled = False
End If
End Sub
Private Sub SelectProject_Click()
If Me.SelectProject.value = vbChecked Then
DcbProject1.Enabled = True
Else
DcbProject1.Enabled = False
DcbProject1.BoundText = ""
End If
End Sub
Private Sub SelectDept_Click()
If Me.SelectDept.value = vbChecked Then
DcpDept1.Enabled = True
Else
DcpDept1.Enabled = False
DcpDept1.BoundText = ""
End If
End Sub
Private Sub Option1_Click()
If Me.Option1.value = True Then
TxtSearchCode.Text = ""
TxtSearchCode.Enabled = False
DCEmployee.BoundText = 0
DCEmployee.Enabled = False
Else
DCEmployee.Enabled = True
TxtSearchCode.Enabled = True
End If
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.DCproject.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„‘—Ê⁄..!!"
        Else
            Msg = "Must Select year Project    ..!!"
        End If
        
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            dcproject.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
        End If
 End If
'    If val(Me.CboYear.ListIndex) = -1 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«— «·”‰…..!!"
'        Else
'            Msg = "Must Select year    ..!!"
'        End If
'
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        CboYear.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
'
'    If val(Me.CmbMonth.ListIndex) = -1 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» ≈Œ Ì«— «·‘Â—..!!"
'        Else
'            Msg = "Must Select Month    ..!!"
'        End If
'
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        CmbMonth.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
  '  End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
            Me.XPTxtID.Text = CStr(new_id("opr_Employee", "ID", "", True))
            
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete opr_employee_details where pk_id=" & val(Me.XPTxtID.Text)
        Cn.Execute "delete ProJectMofrdSalar where pk_id=" & val(Me.XPTxtID.Text)
    End If
    
    rs("ID").value = XPTxtID.Text
   
    rs("Start_date").value = XPDtbTrans.value
    rs("Project_id").value = IIf(Me.DCproject.BoundText = "", Null, Me.DCproject.BoundText)
    rs("opr_type").value = IIf(Me.txtType.Text = "", 0, Me.txtType.Text)
    rs("EmpID1").value = val(Me.DCEmployee.BoundText)
    rs("BrnchID1").value = val(DcbBranch1.BoundText)
    rs("DeptID1").value = val(Me.DcpDept1.BoundText)
    rs("TemID1").value = val(Me.DcbTeam.BoundText)
    rs("ProjectID").value = val(Me.DcbProject1.BoundText)
    If ChckAutoEmp.value = vbChecked Then
    rs("AutoEmp").value = 1
    Else
    rs("AutoEmp").value = 0
    End If
    
    If SelectBranch.value = vbChecked Then
    rs("SelectBranch").value = 1
    Else
    rs("SelectBranch").value = 0
    End If
    If SelectDept.value = vbChecked Then
    rs("SelectDept").value = 1
    Else
    rs("SelectDept").value = 0
    End If
    If SelctTeam.value = vbChecked Then
    rs("SelectTem").value = 1
    Else
    rs("SelectTem").value = 0
    End If
    If SelectProject.value = vbChecked Then
    rs("SelectProj1").value = 1
    Else
    rs("SelectProj1").value = 0
    End If
    If Option1.value = True Then
    rs("SelectAll").value = 1
    Else
    rs("SelectAll").value = 0
    End If
    If Option2.value = True Then
    rs("SelectEmp").value = 1
    Else
    rs("SelectEmp").value = 0
    End If
    
   ' If Me.Dcterm.BoundText <> "" Then
   '     rs("term_Fullcode").value = IIf(Me.Dcterm.BoundText = "", Null, Me.Dcterm.BoundText)
   ' End If
     
   ' If Me.dcopr.BoundText <> "" Then
   '     rs("opr_Fullcode").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
   ' End If
  rs("Years").value = val(CboYear.ListIndex)
  rs("Months").value = val(CmbMonth.ListIndex)
  If RdTypePay(1).value = True Then
  rs("TypePay").value = 1
  Else
  rs("TypePay").value = Null
  End If
If Me.toid.Text = "" Then
'rs("Start_date").value = Null
'rs("toid").value = Null
Else
'rs("Start_date").value = end_date.value
'rs("toid").value = Me.toid.text
End If

''// 01 06 2015
    rs("PandID").value = IIf(Me.Dcterm.BoundText = "", Null, Me.Dcterm.BoundText)
    rs("OpraID").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
    If Me.Fromdate.value <> "" Then
     rs("FromDate").value = Fromdate.value
     Else
     rs("FromDate").value = Null
     End If
    If Me.todate.value <> "" Then
     rs("ToDate").value = Fromdate.value
     Else
     rs("ToDate").value = Null
     End If
    
    
    rs.update
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("Emp_id"))) <> 0 Then
         
                RsDev.AddNew
                RsDev("pk_id").value = Me.XPTxtID.Text
                RsDev("JobTypeID").value = val(.TextMatrix(i, .ColIndex("JobTypeID")))
                RsDev("DepartmentID").value = val(.TextMatrix(i, .ColIndex("DepartmentID")))
                RsDev("BranchId").value = val(.TextMatrix(i, .ColIndex("BranchId")))
                RsDev("project_id").value = val(.TextMatrix(i, .ColIndex("project_id")))
                RsDev("SpecificationID").value = val(.TextMatrix(i, .ColIndex("SpecificationID")))
                
                RsDev("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
                RsDev("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
                RsDev("NumEkama").value = .TextMatrix(i, .ColIndex("NumEkama"))
                RsDev("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName"))
                RsDev("Emp_id").value = val(.TextMatrix(i, .ColIndex("Emp_id")))
                RsDev("Start_date").value = XPDtbTrans.value
                RsDev("Project_id").value = val(.TextMatrix(i, .ColIndex("ProjectID"))) ''IIf(Me.dcproject.BoundText = "", Null, Me.dcproject.BoundText)
                RsDev("opr_type").value = IIf(Me.txtType.Text = "", 0, Me.txtType.Text)
                RsDev("ContProjSalar").value = IIf(.TextMatrix(i, .ColIndex("ContProjSalar")) = "", 1, val(.TextMatrix(i, .ColIndex("ContProjSalar"))))

                If Me.Dcterm.BoundText <> "" Then
                    RsDev("term_Fullcode").value = IIf(Me.Dcterm.BoundText = "", Null, Me.Dcterm.BoundText)
                End If
            
                If Me.dcopr.BoundText <> "" Then
                    RsDev("opr_Fullcode").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
                End If
                RsDev("PandID").value = val(.TextMatrix(i, .ColIndex("PandID")))
                RsDev("OperID").value = val(.TextMatrix(i, .ColIndex("OperID")))
                RsDev("FromDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("FromDate"))), .TextMatrix(i, .ColIndex("FromDate")), Null)
                RsDev("ToDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("ToDate"))), .TextMatrix(i, .ColIndex("ToDate")), Null)

                RsDev("OldToDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("ToDate"))), .TextMatrix(i, .ColIndex("ToDate")), Null)
                
               ' If .TextMatrix(i, .ColIndex("toid")) <> "" Then
              ' RsDev("toid").value = .TextMatrix(i, .ColIndex("toid"))
             '  RsDev("end_date").value = .TextMatrix(i, .ColIndex("end_date"))
              ' End If
                RsDev("interval").value = val(.TextMatrix(i, .ColIndex("interval")))
            
                RsDev("ProjectID").value = val(.TextMatrix(i, .ColIndex("ProjectID")))
       
                save_employee_current_status val(.TextMatrix(i, .ColIndex("ProjectID"))), val(.TextMatrix(i, .ColIndex("PandID"))), val(.TextMatrix(i, .ColIndex("OperID"))), val(.TextMatrix(i, .ColIndex("Emp_id")))
           '   If .TextMatrix(i, .ColIndex("end_date")) = "" Then
           '     save_employee_prohectt_EndDate .TextMatrix(i, .ColIndex("end_date")), val(Me.xptxtid.text), val(.TextMatrix(i, .ColIndex("Emp_id")))
           '   End If
             If SalaryType(val(.TextMatrix(i, .ColIndex("Emp_id")))) = 4 Then
                SaveSalaryProject val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("ProjectID"))), val(.TextMatrix(i, .ColIndex("interval"))), GetTypeEmployee(val(.TextMatrix(i, .ColIndex("Emp_id")))), IIf(IsDate(.TextMatrix(i, .ColIndex("FromDate"))), .TextMatrix(i, .ColIndex("FromDate")), Fromdate.value), IIf(IsDate(.TextMatrix(i, .ColIndex("ToDate"))), .TextMatrix(i, .ColIndex("ToDate")), "")
             Else
                SaveSalaryCompany val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("ProjectID"))), val(.TextMatrix(i, .ColIndex("interval"))), IIf(IsDate(.TextMatrix(i, .ColIndex("FromDate"))), .TextMatrix(i, .ColIndex("FromDate")), Fromdate.value), IIf(IsDate(.TextMatrix(i, .ColIndex("ToDate"))), .TextMatrix(i, .ColIndex("ToDate")), "")
             End If
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
            Msg = "This is Record Already Saved " & CHR(13)
            Msg = Msg + " You Need To Enter Another Recoed "
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
             MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
    Retrive
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
        Msg = "Can not save data" & CHR(13)
      End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
 Else
  Msg = "Sorry...error douring save data" & CHR(13)
 End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Function SalaryType(Optional Emp_id As Double) As Integer
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     Emp_ID, SalaryType"
sql = sql & " From dbo.TblEmployee"
sql = sql & " Where (Emp_id = " & Emp_id & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
SalaryType = IIf(IsNull(Rs7("SalaryType").value), 0, Rs7("SalaryType").value)
Else
SalaryType = 0
End If
End Function
Sub SaveSalaryCompany(Optional EmpID As Double, Optional ProjID As Double, Optional NoDay As Double, Optional Fromdate As Date, Optional todate As String = "")
    Dim Rs7 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql As String
    Dim Value1 As Double
    
'    If IsDate(ToDate) Then
'        'NoDay =
'        NoDay = DateDiff("d", Fromdate, ToDate) + 1
'    Else
'        NoDay = DateDiff("d", Fromdate, MonthLastDay(Fromdate)) + 1
'    End If
    GetNoOfDays Fromdate, todate
    If NoDay > 30 Then NoDay = 30
    Set Rs7 = New ADODB.Recordset
    Rs7.Open "ProJectMofrdSalar", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    sql = "Select * from dbo.EmpSalaryComponent  where emp_ID=" & EmpID & ""
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs4.RecordCount > 0 Then
    Rs4.MoveFirst
    Dim i As Integer
        For i = 1 To Rs4.RecordCount
                Rs7.AddNew
                Value1 = 0
                Rs7("pk_id").value = Me.XPTxtID.Text
                Rs7("EmpID").value = EmpID
                Rs7("ProjID").value = ProjID
                Rs7("NoDay").value = NoDay
                Rs7("YearID").value = val(CboYear.ListIndex)
                Rs7("MonthID").value = val(CmbMonth.ListIndex)
                
                
                Value1 = IIf(IsNull(Rs4("Value").value), 0, Rs4("Value").value)
                Value1 = Value1 / 30
                
                Rs7("Valuee").value = Round(Value1, 2)
                Rs7("MofrdID").value = IIf(IsNull(Rs4("AccountCode").value), 0, Rs4("AccountCode").value)
                Rs7("Total").value = Round((Value1 * NoDay), 0)
                Rs7("TypeSalary").value = 0
                Rs7("FromDate").value = Fromdate
                If IsDate(todate) Then
                    Rs7("ToDate").value = todate
                Else
                    
                    Rs7("ToDate").value = Null
                End If
                Rs7.update
                Rs4.MoveNext
        Next i
    End If
End Sub
Public Function MonthLastDay(ByVal dCurrDate As Date)
  Dim dFirstDayNextMonth As Date
  
  On Error GoTo lbl_Error
 
  MonthLastDay = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
  MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  MonthLastDay = ""
  Exit Function
lbl_Error:
  MsgBox Err.description, vbOKOnly + vbExclamation
End Function

Public Function MonthLastDay2(ByVal dCurrDate As Date)
  Dim dFirstDayNextMonth As Date
  
  On Error GoTo lbl_Error
 
  MonthLastDay2 = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
  MonthLastDay2 = DateAdd("d", -1, dFirstDayNextMonth)
  'MonthLastDay = ""
  Exit Function
lbl_Error:
  MsgBox Err.description, vbOKOnly + vbExclamation
End Function
Private Sub GetNoOfDays(ByVal mFromDate As String, ByVal mToDate As String)
    If IsDate(mFromDate) And Not IsDate(mToDate) Then
        NoDay = 30 - Day(mFromDate) + 1
    ElseIf IsDate(mToDate) Then
        NoDay = DateDiff("d", mFromDate, mToDate) + 1
        
    End If
    
End Sub
Sub SaveSalaryProject(Optional EmpID As Double, Optional ProjID As Double, Optional NoDay As Double, Optional TypeEmp As Integer, Optional Fromdate As Date, Optional todate As Date)
    Dim Rs7 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql As String
    Dim Value1 As Double
'    If IsDate(ToDate) Then
'        'NoDay =
'        NoDay = DateDiff("d", FromDate, ToDate)
'    Else
'        NoDay = DateDiff("d", FromDate, MonthLastDay(FromDate))
'    End If
    GetNoOfDays Fromdate, todate
    If NoDay > 30 Then NoDay = 30
    
    Set Rs7 = New ADODB.Recordset
    Rs7.Open "ProJectMofrdSalar", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    sql = "Select * from ProJectMofrd  where ProjID=" & ProjID & " and  TypeEmp =" & TypeEmp & ""
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs4.RecordCount > 0 Then
    Rs4.MoveFirst
    Dim i As Integer
        For i = 1 To Rs4.RecordCount
                Rs7.AddNew
                Value1 = 0
                Rs7("pk_id").value = Me.XPTxtID.Text
                Rs7("EmpID").value = EmpID
                Rs7("ProjID").value = ProjID
                Rs7("NoDay").value = NoDay
                Rs7("YearID").value = val(CboYear.ListIndex)
                Rs7("MonthID").value = val(CmbMonth.ListIndex)
                Value1 = IIf(IsNull(Rs4("Valuee").value), 0, Rs4("Valuee").value)
                Value1 = Value1 / 30
                Rs7("Valuee").value = Round(Value1, 2)
                Rs7("MofrdID").value = IIf(IsNull(Rs4("MofrdID").value), 0, Rs4("MofrdID").value)
                Rs7("Total").value = Round((Value1 * NoDay), 0)
                Rs7("TypeSalary").value = 1
                Rs7("FromDate").value = Fromdate
                
                If IsDate(todate) Then
                    Rs7("ToDate").value = todate
                Else
                    
                    Rs7("ToDate").value = Null
                End If
                
                Rs7.update
                Rs4.MoveNext
        Next i
    End If
End Sub
Private Sub DcEmployee_Change()
DcEmployee_Click (0)
End Sub

Private Sub DcEmployee_Click(Area As Integer)

    If val(DCEmployee.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DCEmployee.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ID As Double
If Text1.Text <> "" Then
GetCodeIDProject ID, Text1.Text
DCproject.BoundText = ID
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
           FrmProjectSearch.lblSearchtype.Caption = 30
           FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim ID As Double
If Text2.Text <> "" Then
GetCodeIDProject ID, Text2.Text
DcbProject1.BoundText = ID
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF3 Then
               FrmProjectSearch.lblSearchtype.Caption = 31
               FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode (TxtSearchCode.Text), EmpID
        DCEmployee.BoundText = EmpID
    End If

End Sub
Sub GetEmployee()
 If Me.TxtModFlg.Text <> "R" Then
 'If val(dcproject.BoundText) = 0 Or dcproject.Text = "" Then
 'If SystemOptions.UserInterface = ArabicInterface Then
 'MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
 'Else
 'MsgBox "Please Select Project"
 'End If
 'dcproject.SetFocus
 'Exit Sub
 'End If
If ChckBeginProject(Fromdate.value, val(DCproject.BoundText)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ ﬁ»· »œ«Ì…  «—ÌŒ «·„‘—Ê⁄ "
Else
MsgBox "The allotment date is greater than the start date of the project "
End If
Exit Sub
End If
If Not IsDate(todate) Then
     mToDate = MonthLastDay2(Fromdate)
Else
    mToDate = todate.value
End If
If ChckEndProject(mToDate, val(DCproject.BoundText)) = True Then
mToDate = ""
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ »⁄œ ‰Â«Ì…  «—ÌŒ «·„‘—Ê⁄ "
Else
MsgBox "The allotment date is greater than the end date of the project "
End If
Exit Sub
End If
 
 If SelectBranch.value = vbChecked Then
 If val(DcbBranch1.BoundText) = 0 Or DcbBranch1.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
 Else
 MsgBox "Please Select Branch"
 End If
 DcbBranch1.SetFocus
 Exit Sub
 End If
 End If
  If SelectDept.value = vbChecked Then
 If val(DcpDept1.BoundText) = 0 Or DcpDept1.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «Œ Ì«— «·«œ«—…"
 Else
 MsgBox "Please Select Management "
 End If
 DcpDept1.SetFocus
 Exit Sub
 End If
 End If
   If SelectProject.value = vbChecked Then
 If val(DcbProject1.BoundText) = 0 Or DcbProject1.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
 Else
 MsgBox "Please Select Project "
 End If
 DcbProject1.SetFocus
 Exit Sub
 End If
 End If
    If SelctTeam.value = vbChecked Then
 If val(DcbTeam.BoundText) = 0 Or DcbTeam.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «Œ Ì«— «·›—Ìﬁ"
 Else
 MsgBox "Please Select Team "
 End If
 DcbTeam.SetFocus
 Exit Sub
 End If
 End If
 If val(txtDays.Text) <= 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox " «ﬂœ „‰ ’Õ…  «—ÌŒ «· Œ’Ì’"
 Else
 MsgBox "Please make sure the date"
 End If
 Exit Sub
 End If
 get_all_employee
 End If
End Sub
Sub GetEmployeeProject()
 If Me.TxtModFlg.Text <> "R" Then
If ChckBeginProject(Fromdate.value, val(DCproject.BoundText)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ ﬁ»· »œ«Ì…  «—ÌŒ «·„‘—Ê⁄ "
Else
MsgBox "The allotment date is greater than the start date of the project "
End If
Exit Sub
End If
If Not IsDate(todate) Then
     mToDate = MonthLastDay(Fromdate)
Else
    mToDate = todate.value
End If
If ChckEndProject(mToDate, val(DCproject.BoundText)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ »⁄œ ‰Â«Ì…  «—ÌŒ «·„‘—Ê⁄ "
Else
MsgBox "The allotment date is greater than the end date of the project "
End If
Exit Sub
End If

 If val(DCproject.BoundText) = 0 Or DCproject.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
 Else
 MsgBox "Please Select Project "
 End If
 DCproject.SetFocus
 Exit Sub
 End If
 
 If val(txtDays.Text) <= 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox " «ﬂœ „‰ ’Õ…  «—ÌŒ «· Œ’Ì’"
 Else
 MsgBox "Please make sure the date"
 End If
 Exit Sub
 End If
 get_all_employeeProject
 End If
End Sub
Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap

    Select Case Index
Case 10
If ChckAutoEmp.value = vbChecked Then
GetEmployeeProject
End If
Case 20
GetEmployee
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
                  Me.SelctTeam.value = vbUnchecked
            Me.SelectBranch.value = vbUnchecked
            Me.SelectProject.value = vbUnchecked
            Me.SelectDept.value = vbUnchecked
            SelctTeam_Click
            SelectDept_Click
            SelectDept_Click
            SelectBranch_Click
            SelectProject_Click
            FromDate_Change
ChckAutoEmp_Click
       Accredit.Enabled = True
            XPDtbTrans.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid1.Enabled = True
            
            CboYear.Text = year(Date)
CmbMonth.Text = MonthName(Month(Date))
CmbMonth.ListIndex = Month(Date) - 1
                           GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 1
    

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            If ChKauto.value = vbChecked Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰  ⁄œÌ·  Œ’Ì’ «·Ì ", vbCritical
                Else
                    MsgBox " Can't Delete Auto Employee Allocation ", vbCritical
                End If

                Exit Sub
            End If
FromDate_Change
Dim Msg As String
Dim i As Integer
            TxtModFlg.Text = "E"
           ' Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
        With Me.Grid
              For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("Emp_id"))) <> 0 Then
            If val(.TextMatrix(i, .ColIndex("ProjectID"))) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
             Msg = " Ì—ÃÏ «Œ Ì«—«·„‘—Ê⁄ "
             Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
             Msg = Msg & i
            Else
             Msg = "Please select project "
             Msg = Msg & "In Row " & CHR(13)
             Msg = Msg & i
            End If
            MsgBox Msg
          Exit Sub
          End If
          End If
          Next i
        For i = .FixedRows To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("Emp_id"))) <> 0 Then
            If .TextMatrix(i, .ColIndex("FromDate")) <> "" Then
            If ChckBeginProject(.TextMatrix(i, .ColIndex("FromDate")), val(.TextMatrix(i, .ColIndex("project_id")))) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
             Msg = " ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ ﬁ»· »œ«Ì…  «—ÌŒ «·„‘—Ê⁄ "
             Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
             Msg = Msg & i
            Else
             Msg = "The allotment date is greater than the start date of the project "
             Msg = Msg & "In Row " & CHR(13)
             Msg = Msg & i
            End If
            MsgBox Msg
          Exit Sub
          End If
         
           If ChckDatBeginWork(.TextMatrix(i, .ColIndex("FromDate")), val(.TextMatrix(i, .ColIndex("Emp_id")))) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
             Msg = " ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ ﬁ»·  «—ÌŒ  ⁄ÌÌ‰ «·„ÊŸ› "
             Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
             Msg = Msg & i
            Else
             Msg = "The allotment date is less than the start date of work "
             Msg = Msg & "In Row " & CHR(13)
             Msg = Msg & i
            End If
            MsgBox Msg
          Exit Sub
          End If
          
          End If
          
          If .TextMatrix(i, .ColIndex("ToDate")) <> "" Then
         If ChckEndProject(.TextMatrix(i, .ColIndex("ToDate")), val(.TextMatrix(i, .ColIndex("project_id")))) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
             Msg = " ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰  «—ÌŒ «· Œ’Ì’ »⁄œ ‰Â«Ì…  «—ÌŒ «·„‘—Ê⁄ "
             Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
             Msg = Msg & i
            Else
             Msg = "The allotment date is greater than the end date of the project "
             Msg = Msg & "In Row " & CHR(13)
             Msg = Msg & i
            End If
            MsgBox Msg
          Exit Sub
          End If
          End If
          
            s = " SELECT Null FROM opr_employee_details WHERE "
           s = s & "  IsNull(ToDate,'1-1-2050') > " & SQLDate(.TextMatrix(i, .ColIndex("FromDate")), True) & " And EMp_Id = " & val(.TextMatrix(i, .ColIndex("Emp_id")))
           s = s & " and pk_Id <> " & val(XPTxtID)
           
           Set rsDummy = New ADODB.Recordset
           rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
           If Not rsDummy.EOF Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " «·„ÊŸ› ·Â Õ—ﬂ…  Œ’Ì’ „› ÊÕ… "
                    Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
                    Msg = Msg & i
                Else
                    Msg = "The allotment date is greater than the end date of the project "
                    Msg = Msg & "In Row " & CHR(13)
                    Msg = Msg & i
                End If
                MsgBox Msg
                Exit Sub
            End If
           
         '  End If
'           If val(.TextMatrix(i, .ColIndex("interval"))) <= 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'             Msg = "  «ﬂœ „‰ ’Õ…  «—ÌŒ «· Œ’Ì’ "
'             Msg = Msg & "›Ì «·”ÿ— " & CHR(13)
'             Msg = Msg & i
'            Else
'             Msg = "Make sure that the The allotment date "
'             Msg = Msg & "In Row " & CHR(13)
'             Msg = Msg & i
'            End If
'            MsgBox Msg
'          Exit Sub
'          End If
            End If
          Next i
         End With
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
'
'              Load FrmSearchEmpSalary3
'              FrmSearchEmpSalary3.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
            print_report
  
Case 8
            TxtModFlg.Text = "N"
           ' clear_all Me
            Me.XPTxtID.Text = CStr(new_id("opr_Employee", "ID", "", True))
       
            XPDtbTrans.value = Date
       
            
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
  
  Case 9
  print_report2
        
    End Select

    Exit Sub
ErrTrap:

End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From opr_Employee Where id=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Cn.Execute "delete ProJectMofrdSalar where pk_id=" & val(Me.XPTxtID.Text)
                rs.MoveFirst
                 Cn.Execute "delete opr_employee_details where pk_id=" & val(Me.XPTxtID)

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
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
        XPTxtCurrent.Caption = 0
                        XPTxtCount.Caption = 0
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available.no records there"
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
  Msg = "Sorry error douring delete data"
  End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
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

Private Sub Dcdep_Click(Area As Integer)
'    'CmdOk_Click
End Sub




Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

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

Private Sub dcproject_Change()
dcproject_Click (0)
End Sub

Private Sub dcproject_Click(Area As Integer)
Dim Fullcode As String
  '  If dcproject.BoundText = "" Then Exit Sub
  '  My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
  '  fill_combo Dcterm, My_SQL
  

    If DCproject.BoundText <> "" Then
GetCodeIDProject val(DCproject.BoundText), Fullcode
Text1.Text = Fullcode
        fillterms (val(DCproject.BoundText))
    End If

End Sub
Function fillterms(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.Dcterm, My_SQL
       
        
    dcopr.ReFill
End Function
Private Sub dcproject_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyF5 Then
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Dcombos.GetProjects DCproject
    End If
        If KeyCode = vbKeyF3 Then
           FrmProjectSearch.lblSearchtype.Caption = 30
           FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub Dcterm_Change()
Dcterm_Click (0)
End Sub

Private Sub Dcterm_Click(Area As Integer)

   'If Dcterm.BoundText = "" Then Exit Sub

   ' My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
   ' fill_combo dcopr, My_SQL
    Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
  If DCproject.BoundText <> "" Then
        
         If Me.Dcterm.BoundText <> "" Then
       '  Dcombos.GetProcessOfProjedt
         Dcombos.GetProcessOfProjedt dcopr, val(DCproject.BoundText), , val(Dcterm.BoundText), 2
         End If
       
    End If
End Sub

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

Private Sub Form_Load()

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
 If SystemOptions.UserInterface = ArabicInterface Then
                Grid.ColComboList(Grid.ColIndex("ContProjSalar")) = "#1;ÿ»ﬁ« ·⁄ﬁœ „ÊŸ›|#2;  ÿ»ﬁ« ·„‘—Ê⁄"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Grid.ColComboList(Grid.ColIndex("ContProjSalar")) = "#1;Contract Employees |#2;Project "
            End If
           C1Tab1.CurrTab = 0
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    YearMonth
    Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = " select id,Project_name from projects where not(Project_name is null) and Project_name<>N'""' order by Project_name"
  Else
  My_SQL = " select id,Project_nameE from projects where not (Project_nameE is null) and Project_nameE<>N'""' order by Project_nameE"
  End If
    fill_combo DCproject, My_SQL
    fill_combo DcbProject1, My_SQL
    My_SQL = "    select oprid,des from dbo.projects_des"

    fill_combo Me.Dcterm, My_SQL
   ' My_SQL = " select  fullcode,des from projects_des"
   ' fill_combo Dcterm, My_SQL

   ' My_SQL = " select  fullcode,name from terms_operations"
   ' fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    Dcombos.GetEmpDepartments DcpDept1
    Dcombos.GetBranches DcbBranch1
    'Dcombos.GetProjects DcbProject1
    Dcombos.GetEmpSpecifications Me.DcbTeam
    'Dcombos.GetBranches Dcbranch
    'Dcombos.GetEmpLocations DcbLocation
    Dcombos.GetEmployees DCEmployee
    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
    
Fromdate.value = ""
Me.todate.value = ""
    Set rs = New ADODB.Recordset
    StrSQL = "select * From opr_Employee where opr_type=0 "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print Analysis"
    Cmd(10).Caption = "Add"
    lbl(2).Caption = "To"
    ChckAutoEmp.RightToLeft = False
    ChckAutoEmp.Caption = "Employee From Project"
    Me.SelctTeam.RightToLeft = False
Me.SelctTeam.Caption = "Team"
Me.SelectBranch.RightToLeft = False
Me.SelectBranch.Caption = "Branch"
Me.SelectDept.RightToLeft = False
Me.SelectDept.Caption = " Manage."
Me.SelectProject.RightToLeft = False
Me.SelectProject.Caption = "Project"
lbl(18).Caption = ""
Cmd(20).Caption = "Add"
   ' Frame1.Caption = "Select Employees"
    Option1.Caption = "All Employees"
    Option2.Caption = "Select Emp"
    
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    lbl(12).Caption = "Method"
     RdTypePay(0).RightToLeft = False
     RdTypePay(1).RightToLeft = False
    'CmdHelp.Caption = "Help"
    RdTypePay(0).Caption = "By Contract"
    RdTypePay(1).Caption = "By Project"
Accredit.Caption = "Send To Approve"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    C1Tab1.FirstTab = 0
    C1Tab1.CurrTab = 0
    C1Tab1.TabCaption(0) = "Data"
    C1Tab1.TabCaption(2) = "Data Salary"
    C1Tab1.TabCaption(1) = "Approve Status"
    Me.Caption = "Projects Labors Allocate"
    ELe(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Start Date"
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
    lbl(0).Caption = "Terms"
    lbl(4).Caption = "Operation"
    lbl(5).Caption = "Project"
    lbl(3).Caption = "Year"
    lbl(6).Caption = "Month"
    lbl(9).Caption = "From"
    lbl(10).Caption = "To"
    lbl(11).Caption = "Num Day"
    Cmd(7).Caption = "Print"
    Cmd(8).Caption = "A similar version"
ELe(3).Caption = "Priod"
    Check1.Caption = "Show All Employee"

    CmdRemove.Caption = "Delete Line"
    CmdRemoveAll.Caption = "Delete All"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Serial"
        .TextMatrix(0, .ColIndex("NumEkama")) = "ID Number"
        .TextMatrix(0, .ColIndex("ContProjSalar")) = "Methode"
        .TextMatrix(0, .ColIndex("Emp_code")) = "Emp Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp Name"
        .TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
        .TextMatrix(0, .ColIndex("work_status")) = "Work Status"
        .TextMatrix(0, .ColIndex("project_name")) = "Project Name"
        .TextMatrix(0, .ColIndex("cost_center")) = "Cost Center"
        .TextMatrix(0, .ColIndex("work_days")) = "Work Days"
        .TextMatrix(0, .ColIndex("ATTENDANCE")) = "Absence"
        .TextMatrix(0, .ColIndex("late")) = "Delay"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
        .TextMatrix(0, .ColIndex("net_work_days")) = "Net Work Days"
        .TextMatrix(0, .ColIndex("addition")) = "Over Time"
        .TextMatrix(0, .ColIndex("remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("Project")) = "Project"
        .TextMatrix(0, .ColIndex("pand")) = "Pand"
        .TextMatrix(0, .ColIndex("opra")) = "Process"
        .TextMatrix(0, .ColIndex("FromDate")) = "From"
        .TextMatrix(0, .ColIndex("ToDate")) = "To"
        .TextMatrix(0, .ColIndex("interval")) = "Num Day"
        .TextMatrix(0, .ColIndex("PrjectCode")) = "Project Code"
        
    End With
        With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Emp Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp Name"
        .TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
        .TextMatrix(0, .ColIndex("mofrad_name")) = "Component"
        .TextMatrix(0, .ColIndex("Valuee")) = "Value"
        .TextMatrix(0, .ColIndex("NoDay")) = "No Day"
        .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("TypeSalary")) = "Type Salary"
    End With
lbl(14).Caption = "Curr.Record"
lbl(13).Caption = "No.Record"
With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "LevelName"
        .TextMatrix(0, .ColIndex("EmpName")) = "EmpName"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "ApprovDate"
        .TextMatrix(0, .ColIndex("Remarks")) = " Remarks"
        '.TextMatrix(0, .ColIndex("Convert")) = "Convert To Bill"

    End With
    
    
End Sub

Function GetEmpIDes() As String
    Dim tempString As String
    Dim i As Integer
    tempString = "0,0"
    With Grid
    For i = 0 To .Rows - 1
    If val(.TextMatrix(i, .ColIndex("Emp_id"))) <> 0 Then
        tempString = tempString & "," & val(.TextMatrix(i, .ColIndex("Emp_id")))
    End If
    Next i
    End With
    GetEmpIDes = tempString
End Function
Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer
    Dim CuurRow As Integer
    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details where 1<>-1 "
    sql = sql & " and  BignDateWork <=" & SQLDate(Fromdate.value, True) & ""
    
     If val(DcbTeam.BoundText) <> 0 And DcbTeam.BoundText <> "" Then
    sql = sql & " and  SpecificationID =" & val(DcbTeam.BoundText) & ""
    End If
    If val(DCEmployee.BoundText) <> 0 And DCEmployee.BoundText <> "" Then
    sql = sql & " and  Emp_ID =" & val(DCEmployee.BoundText) & ""
    End If
    If val(DcpDept1.BoundText) <> 0 And DcpDept1.BoundText <> "" Then
    sql = sql & " and  DepartmentID =" & val(DcpDept1.BoundText) & ""
    End If
    If val(DcbBranch1.BoundText) <> 0 And DcbBranch1.BoundText <> "" Then
    sql = sql & " and  BranchId =" & val(DcbBranch1.BoundText) & ""
    End If
    If val(DcbProject1.BoundText) <> 0 And DcbProject1.BoundText <> "" Then
    sql = sql & " and  project_id =" & val(DcbProject1.BoundText) & ""
    End If
    sql = sql & " and  Emp_ID not in (" & GetEmpIDes & " ) "
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid
If .Rows = 1 Then
      .Rows = .Rows + Rs3.RecordCount
        CuurRow = 1
     Else
      CuurRow = .Rows
       .Rows = .Rows + Rs3.RecordCount
     End If
      '  .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
                         Rs3.MoveFirst
         
            For i = CuurRow To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
              
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
               Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Namee").value), "", Rs3.Fields("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeNamee").value), "", Rs3.Fields("JobTypeNamee").value)
               End If
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                .TextMatrix(i, .ColIndex("JobTypeID")) = IIf(IsNull(Rs3.Fields("JobTypeID").value), "", Rs3.Fields("JobTypeID").value)
                .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(Rs3.Fields("DepartmentID").value), "", Rs3.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs3.Fields("BranchId").value), "", Rs3.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("project_id")) = IIf(IsNull(Rs3.Fields("project_id").value), "", Rs3.Fields("project_id").value)
                .TextMatrix(i, .ColIndex("SpecificationID")) = IIf(IsNull(Rs3.Fields("SpecificationID").value), "", Rs3.Fields("SpecificationID").value)
                  .TextMatrix(i, .ColIndex("PandID")) = Dcterm.BoundText
                .TextMatrix(i, .ColIndex("pand")) = Dcterm.Text
                .TextMatrix(i, .ColIndex("OperID")) = dcopr.BoundText
                .TextMatrix(i, .ColIndex("opra")) = dcopr.Text
                If val(txtDays.Text) = 0 Then
                 .TextMatrix(i, .ColIndex("interval")) = 30
                Else
                  .TextMatrix(i, .ColIndex("interval")) = txtDays.Text
                   .TextMatrix(i, .ColIndex("FromDate")) = Fromdate.value & ""
                    .TextMatrix(i, .ColIndex("ToDate")) = Me.todate.value & ""
                   
                 End If
                 .TextMatrix(i, .ColIndex("ProjectID")) = val(DCproject.BoundText)
                .TextMatrix(i, .ColIndex("Project")) = DCproject.Text
                Rs3.MoveNext

            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub get_all_employeeProject()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer
    Dim CuurRow As Integer
    Dim sql As String
    Dim i As Long

    sql = " SELECT     dbo.TblEmpOper.OperCode, dbo.TblEmpOper.[Count], dbo.TblEmpOper.daysalary, dbo.TblEmpOper.JobID, dbo.TblEmpJobsTypes.JobTypeName, "
    sql = sql & "                  dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpOper.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
    sql = sql & "                  dbo.TblEmpOper.Opr, dbo.TblEmpOper.Pand, dbo.TblEmpOper.ProjectID, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName,"
    sql = sql & "                  dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    sql = sql & "                  dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.projects.Project_name, dbo.projects.Project_nameE,"
    sql = sql & "                  dbo.projects.Fullcode AS Expr1, dbo.projects_des.des, dbo.TblEmployee.NumEkama ,dbo.TblEmployee.SpecificationID"
    sql = sql & "       FROM         dbo.projects_des RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblEmpOper ON dbo.projects_des.oprid = dbo.TblEmpOper.Pand LEFT OUTER JOIN"
    sql = sql & "                  dbo.projects ON dbo.TblEmpOper.ProjectID = dbo.projects.id LEFT OUTER JOIN"
    sql = sql & "                  dbo.terms_operations LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID ON dbo.TblEmpOper.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.TblBranchesData.branch_id = dbo.TblEmployee.BranchId RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID ON"
    sql = sql & "                  dbo.TblEmpOper.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmpJobsTypes ON dbo.TblEmpOper.JobID = dbo.TblEmpJobsTypes.JobTypeID"
    sql = sql & "  where  (NOT (dbo.TblEmployee.Fullcode IS NULL))"
 '   Sql = Sql & " and  dbo.TblEmployee.BignDateWork <=" & SQLDate(FromDate.value, True) & ""
    If val(DCproject.BoundText) <> 0 And DCproject.BoundText <> "" Then
    sql = sql & " and  dbo.TblEmpOper.ProjectID  =" & val(DCproject.BoundText) & ""
    End If
        If val(Dcterm.BoundText) <> 0 And Dcterm.BoundText <> "" Then
    sql = sql & " and  dbo.TblEmpOper.Pand  =" & val(Dcterm.BoundText) & ""
    End If
        If val(dcopr.BoundText) <> 0 And dcopr.BoundText <> "" Then
    sql = sql & " and  dbo.terms_operations.OPRIDD  =" & val(dcopr.BoundText) & ""
    End If
    
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid
If .Rows = 1 Then
      .Rows = .Rows + Rs3.RecordCount
        CuurRow = 1
     Else
      CuurRow = .Rows
       .Rows = .Rows + Rs3.RecordCount
     End If
        If Rs3.RecordCount > 0 Then
                         Rs3.MoveFirst
         
            For i = CuurRow To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Fullcode").value), "", Rs3.Fields("Fullcode").value)
                .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(Rs3.Fields("NumEkama").value), "", Rs3.Fields("NumEkama").value)
                .TextMatrix(i, .ColIndex("PrjectCode")) = IIf(IsNull(Rs3.Fields("Expr1").value), "", Rs3.Fields("Expr1").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(Rs3.Fields("Project_name").value), "", Rs3.Fields("Project_name").value)
                .TextMatrix(i, .ColIndex("opra")) = IIf(IsNull(Rs3.Fields("ProcessName").value), "", Rs3.Fields("ProcessName").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
               Else
                .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(Rs3.Fields("Project_nameE").value), "", Rs3.Fields("Project_nameE").value)
                .TextMatrix(i, .ColIndex("opra")) = IIf(IsNull(Rs3.Fields("ProcessNameE").value), "", Rs3.Fields("ProcessNameE").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Namee").value), "", Rs3.Fields("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeNamee").value), "", Rs3.Fields("JobTypeNamee").value)
               End If
            
                .TextMatrix(i, .ColIndex("JobTypeID")) = IIf(IsNull(Rs3.Fields("JobID").value), "", Rs3.Fields("JobID").value)
                .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(Rs3.Fields("DepartmentID").value), "", Rs3.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs3.Fields("BranchId").value), "", Rs3.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("project_id")) = IIf(IsNull(Rs3.Fields("ProjectID").value), "", Rs3.Fields("ProjectID").value)
                .TextMatrix(i, .ColIndex("SpecificationID")) = IIf(IsNull(Rs3.Fields("SpecificationID").value), "", Rs3.Fields("SpecificationID").value)
                .TextMatrix(i, .ColIndex("PandID")) = IIf(IsNull(Rs3.Fields("Pand").value), "", Rs3.Fields("Pand").value)
                .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(Rs3.Fields("des").value), "", Rs3.Fields("des").value)
                .TextMatrix(i, .ColIndex("OperID")) = IIf(IsNull(Rs3.Fields("OPRIDD").value), "", Rs3.Fields("OPRIDD").value)
                .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(Rs3.Fields("ProjectID").value), "", Rs3.Fields("ProjectID").value)
                Grid_AfterEdit i, .ColIndex("Emp_Name")
                If val(txtDays.Text) = 0 Then
                 .TextMatrix(i, .ColIndex("interval")) = 30
                Else
                  .TextMatrix(i, .ColIndex("interval")) = txtDays.Text
                   .TextMatrix(i, .ColIndex("FromDate")) = Fromdate.value & ""
                    .TextMatrix(i, .ColIndex("ToDate")) = Me.todate.value & ""
                   
                 End If
                 .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("EmpID").value), "", Rs3.Fields("EmpID").value)
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .Rows = .Rows + 1
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
             .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "Total"

        End If
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub FromDate_Change()
If Me.Fromdate.value <> "" Then




If Not IsDate(todate) Then
     mToDate = MonthLastDay(Fromdate)
Else
    mToDate = todate.value
End If

GetNoOfDays Me.Fromdate.value, mToDate
'Me.txtDays.Text = DateDiff("d", Me.FromDate.value, mToDate) + 1
txtDays = NoDay
If val(txtDays) > 30 Then txtDays = 30
End If

End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    

    With Grid
    If .TextMatrix(Row, .ColIndex("FromDate")) = "" Then
                    .TextMatrix(Row, .ColIndex("FromDate")) = Fromdate.value & ""
                    .TextMatrix(Row, .ColIndex("ToDate")) = Me.todate.value & ""
                    
                    If Not IsDate(todate) Then
                        If IsDate(Fromdate.value) Then
                            mToDate = MonthLastDay(Fromdate)
                        End If
                    Else
                        mToDate = todate.value
                    End If
                    
                    If IsDate(Fromdate.value) And IsDate(mToDate) Then
                        'NoDay = DateDiff("d", FromDate.value, mToDate) + 1
                        GetNoOfDays Fromdate.value, mToDate
                        If NoDay > 30 Then NoDay = 30
                        .TextMatrix(Row, .ColIndex("interval")) = NoDay
                        
                    End If
          End If
        Select Case .ColKey(Col)
      
 
            Case "Emp_Name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_id")) = StrAccountCode
                If val(.TextMatrix(Row, .ColIndex("ProjectID"))) = 0 Then
                .TextMatrix(Row, .ColIndex("ProjectID")) = DCproject.BoundText
                .TextMatrix(Row, .ColIndex("Project")) = DCproject.Text
                .TextMatrix(Row, .ColIndex("PandID")) = Dcterm.BoundText
                .TextMatrix(Row, .ColIndex("pand")) = Dcterm.Text
                .TextMatrix(Row, .ColIndex("OperID")) = dcopr.BoundText
                .TextMatrix(Row, .ColIndex("opra")) = dcopr.Text
                End If
                If SalaryType(val(.TextMatrix(Row, .ColIndex("Emp_id")))) = 4 Then
                .TextMatrix(Row, .ColIndex("ContProjSalar")) = 2
                Else
                .TextMatrix(Row, .ColIndex("ContProjSalar")) = 1
                End If
                If val(txtDays.Text) = 0 Then
                 .TextMatrix(Row, .ColIndex("interval")) = 30
                Else
                  .TextMatrix(Row, .ColIndex("interval")) = txtDays.Text
                  .TextMatrix(Row, .ColIndex("FromDate")) = Fromdate.value
                    .TextMatrix(Row, .ColIndex("ToDate")) = Me.todate.value & ""
                 End If
             
                StrSQL = "SELECT  * from emp_all_details Where Emp_id=" & val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                    
                        .TextMatrix(Row, .ColIndex("NumEkama")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
                        .TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                         .TextMatrix(Row, .ColIndex("JobTypeID")) = IIf(IsNull(rs("JobTypeID").value), 0, rs("JobTypeID").value)
                        
                        If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                        Else
                        .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                        End If
                            
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "Emp_Code"
            If val(.TextMatrix(Row, .ColIndex("ProjectID"))) = 0 Then
                  .TextMatrix(Row, .ColIndex("ProjectID")) = DCproject.BoundText
                .TextMatrix(Row, .ColIndex("Project")) = DCproject.Text
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
                .TextMatrix(Row, .ColIndex("PandID")) = Dcterm.BoundText
                .TextMatrix(Row, .ColIndex("pand")) = Dcterm.Text
                .TextMatrix(Row, .ColIndex("OperID")) = dcopr.BoundText
                .TextMatrix(Row, .ColIndex("opra")) = dcopr.Text
               End If
                If val(txtDays.Text) = 0 Then
                 .TextMatrix(Row, .ColIndex("interval")) = 30
                Else
                  .TextMatrix(Row, .ColIndex("interval")) = txtDays.Text
                  .TextMatrix(Row, .ColIndex("FromDate")) = Fromdate.value & ""
                    .TextMatrix(Row, .ColIndex("ToDate")) = Me.todate.value & ""
                 End If

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from emp_all_details Where fullcode=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                  Else
                    .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                    .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                  End If
                  .TextMatrix(Row, .ColIndex("JobTypeID")) = IIf(IsNull(rs("JobTypeID").value), 0, rs("JobTypeID").value)
                    .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                   .TextMatrix(Row, .ColIndex("NumEkama")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
              
                Else
                    .TextMatrix(Row, .ColIndex("JobTypeName")) = ""
              
                    .TextMatrix(Row, .ColIndex("Emp_Name")) = ""
              
                    .TextMatrix(Row, .ColIndex("Emp_id")) = ""
              
                End If
                    If SalaryType(val(.TextMatrix(Row, .ColIndex("Emp_id")))) = 4 Then
                .TextMatrix(Row, .ColIndex("ContProjSalar")) = 2
                Else
                .TextMatrix(Row, .ColIndex("ContProjSalar")) = 1
                End If
       Case "Project"
        StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                If StrAccountCode <> "" Then
                StrSQL = " SELECT Fullcode  From dbo.Projects where id =" & val(StrAccountCode) & ""
                End If
                     Set rs = New ADODB.Recordset
      rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      .TextMatrix(Row, .ColIndex("PrjectCode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
       Case "PrjectCode"
       If .TextMatrix(Row, .ColIndex("PrjectCode")) <> "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name , id From dbo.Projects where not(Project_name is null) and Project_name <>N'""' "
           Else
               StrSQL = " SELECT  LTRIM(RTRIM( Project_nameE )) as Project_nameE , id From dbo.Projects where not(Project_nameE is null) and Project_nameE <>N'""' "
       End If
       StrSQL = StrSQL & " and Fullcode= N'" & .TextMatrix(Row, .ColIndex("PrjectCode")) & "'"
       Set rs = New ADODB.Recordset
      rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If rs.RecordCount > 0 Then
       .TextMatrix(Row, .ColIndex("project_id")) = IIf(IsNull(rs("id").value), 0, rs("id").value)
       If SystemOptions.UserInterface = ArabicInterface Then
       .TextMatrix(Row, .ColIndex("Project")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
       Else
       .TextMatrix(Row, .ColIndex("Project")) = IIf(IsNull(rs("Project_nameE").value), "", rs("Project_nameE").value)
       End If
       End If
       End If
 Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("PandID")) = StrAccountCode
                  Case "opra"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("OperID")) = StrAccountCode
              Case "FromDate"
              If (IsDate(.TextMatrix(Row, .ColIndex("FromDate")))) Then
              If (IsDate(.TextMatrix(Row, .ColIndex("ToDate")))) Then
               ' NoDay = DateDiff("d", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))) + 1
                GetNoOfDays .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))
                
                If NoDay > 30 Then NoDay = 30
              .TextMatrix(Row, .ColIndex("interval")) = NoDay
              End If
              End If
              Case "ToDate"
               If (IsDate(.TextMatrix(Row, .ColIndex("FromDate")))) Then
              If (IsDate(.TextMatrix(Row, .ColIndex("ToDate")))) Then
                'NoDay = DateDiff("d", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))) + 1
                GetNoOfDays .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))
              If NoDay > 30 Then NoDay = 30
              .TextMatrix(Row, .ColIndex("interval")) = NoDay
              End If
              End If
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            .TextMatrix(i, .ColIndex("Start_date")) = XPDtbTrans.value
          '  mToDate = .TextMatrix(i, .ColIndex("ToDate"))
            If Not IsDate(.TextMatrix(i, .ColIndex("ToDate"))) And Not IsDate(todate.value) Then
                 If IsDate(.TextMatrix(i, .ColIndex("FromDate"))) Then
                    mToDate = MonthLastDay(.TextMatrix(i, .ColIndex("FromDate")))
                End If
            ElseIf IsDate(.TextMatrix(i, .ColIndex("ToDate"))) Then
                mToDate = .TextMatrix(i, .ColIndex("ToDate"))
            ElseIf IsDate(todate.value) Then
                mToDate = todate.value
            End If

 ' If .TextMatrix(i, .ColIndex("end_date")) <> "" Then
 If IsDate(.TextMatrix(i, .ColIndex("FromDate"))) Then
   ' NoDay = DateDiff("D", .TextMatrix(i, .ColIndex("FromDate")), mToDate) + 1
    GetNoOfDays .TextMatrix(i, .ColIndex("FromDate")), mToDate
    If NoDay > 30 Then NoDay = 30
    .TextMatrix(i, .ColIndex("interval")) = NoDay
End If
 ' End If
            End If

        Next i
   
    End With

End Sub
Function ChckDatBeginWork(Optional RecDate As Date, Optional Emp_id As Double) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT     Emp_ID From dbo.emp_all_details"
sql = sql & " WHERE     (BignDateWork > " & SQLDate(RecDate, True) & ") and Emp_ID=" & Emp_id & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
ChckDatBeginWork = True
Else
ChckDatBeginWork = False
End If
End Function
Function ChckEndProject(Optional ByVal RecDate As String, Optional ProjectID As Double) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset

sql = " SELECT     ID From dbo.projects"
sql = sql & " WHERE     (EndDate < " & SQLDate(CDate(RecDate), True) & ") and ID=" & ProjectID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
ChckEndProject = True
Else
ChckEndProject = False
End If
End Function
Function ChckBeginProject(Optional RecDate As Date, Optional ProjectID As Double) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT     ID From dbo.projects"
sql = sql & " WHERE     (StartDate > " & SQLDate(RecDate, True) & ") and ID=" & ProjectID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
ChckBeginProject = True
Else
ChckBeginProject = False
End If
End Function
Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        Select Case .ColKey(Col)
        Case "PrjectCode"
         .ComboList = ""
            Case "Emp_Code"
                .ComboList = ""

            Case "JobTypeName"
                .ComboList = ""
        
            Case "DepartmentName"
                .ComboList = ""
        
            Case "work_status"
                .ComboList = ""

            Case "work_days"
                .ComboList = ""

            Case "attendance"
                .ComboList = ""

            Case "late"
                .ComboList = ""

            Case "discount"
                .ComboList = ""

            Case "net_work_days"
                .ComboList = ""

            Case "addition"
                .ComboList = ""

            Case "remarks"
                .ComboList = ""

            Case "absence"
                .ComboList = ""
                '  Cancel = True
             Case "interval"
                Cancel = True
                Case "interval"
                .ComboList = ""
                Case "FromDate"
                .ComboList = ""
                Case "ToDate"
                .ComboList = ""
                Case "ContProjSalar"
                Cancel = True
            Case "ToDate"
                .EditMaxLength = 10
           
        End Select

    End With

End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
    With Grid

        Select Case .ColKey(.Col)

                 Case "Emp_Code", "Emp_Name"
              
                  LongRow = .Row


   If KeyCode = vbKeyF3 Then

        FrmEmployeeSearch.lbltype = 30
        Set FrmEmployeeSearch.RetrunFrm = Me
'
        FrmEmployeeSearch.show
  
    End If
    
                               End Select
             End With
        End If
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
        
                'Full Path Display

                StrSQL = "SELECT *  FROM emp_all_details "
            
                'StrSQL = StrSQL & " where  BignDateWork <=" & SQLDate(FromDate.value, True) & ""
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = Grid.BuildComboList(rs, "Emp_Name", "Emp_id")
                Else
                StrComboList = Grid.BuildComboList(rs, "Emp_Namee", "Emp_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                Case "Project"
If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT  LTRIM(RTRIM( Project_name )) as Project_name , id From dbo.Projects where not(Project_name is null) and Project_name <>N'""' "
Else
               StrSQL = " SELECT  LTRIM(RTRIM( Project_nameE )) as Project_nameE , id From dbo.Projects where not(Project_nameE is null) and Project_nameE <>N'""' "
End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If SystemOptions.UserInterface = ArabicInterface Then
                   StrComboList = Grid.BuildComboList(rs, "Project_name", "id")
Else
                    StrComboList = Grid.BuildComboList(rs, "Project_nameE", "id")
End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                          Case "pand"
             If .TextMatrix(Row, .ColIndex("ProjectID")) = "" Then
             If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Else
             MsgBox "Please Select Project"
             End If
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & ")"
           
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = Grid.BuildComboList(rs, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  Case "opra"
                   
If .TextMatrix(Row, .ColIndex("ProjectID")) = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Else
             MsgBox "Please Select Project"
             End If
.TextMatrix(Row, .ColIndex("opra")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("PandID")) = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
Else
MsgBox "Please Select Des"
End If
.TextMatrix(Row, .ColIndex("opra")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("PandID"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & ")"
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = Grid.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
  Case "FromDate"
          .ColComboList(.ColIndex("FromDate")) = "..."
Case "ToDate"
            .ColComboList(.ColIndex("ToDate")) = "..."
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
           If Lngid <> 0 Then
        rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

 
 
    Me.XPTxtID.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
 
    XPDtbTrans.value = IIf(IsNull(rs("Start_date").value), Date, rs("Start_date").value)
    
    
'    If IsNull(rs("toid").value) Then
'           end_date.Visible = False
'             Me.toid.text = ""
'
'    Else
'        end_date.Visible = True
'        end_date.value = IIf(IsNull(rs("end_date").value), Date, rs("end_date").value)
'    Me.toid.text = IIf(IsNull(rs("toid").value), "", rs("toid").value)
'    End If
If Not IsNull(rs("AutoEmp").value) Then
 If rs("AutoEmp").value = 1 Then
 ChckAutoEmp.value = vbChecked
 Else
 ChckAutoEmp.value = vbUnchecked
 End If
 Else
 ChckAutoEmp.value = vbUnchecked
 End If
 
If Not IsNull(rs("TypePay").value) Then
 If rs("TypePay").value = 1 Then
 RdTypePay(1).value = True
 Else
 RdTypePay(0).value = True
 End If
 Else
  RdTypePay(0).value = True
 End If
''////////////
Me.DCEmployee.BoundText = IIf(IsNull(rs("EmpID1").value), "", rs("EmpID1").value)
Me.DcbBranch1.BoundText = IIf(IsNull(rs("BrnchID1").value), "", rs("BrnchID1").value)
Me.DcpDept1.BoundText = IIf(IsNull(rs("DeptID1").value), "", rs("DeptID1").value)
Me.DcbTeam.BoundText = IIf(IsNull(rs("TemID1").value), "", rs("TemID1").value)
Me.DcbProject1.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
If Not IsNull(rs("SelectBranch").value) Then
If (rs("SelectBranch").value) = 1 Then
Me.SelectBranch.value = vbChecked
Else
Me.SelectBranch.value = vbUnchecked
End If
Else
Me.SelectBranch.value = vbUnchecked
End If

If Not IsNull(rs("SelectDept").value) Then
If (rs("SelectDept").value) = 1 Then
Me.SelectDept.value = vbChecked
Else
Me.SelectDept.value = vbUnchecked
End If
Else
Me.SelectDept.value = vbUnchecked
End If

If Not IsNull(rs("SelectTem").value) Then
If (rs("SelectTem").value) = 1 Then
Me.SelctTeam.value = vbChecked
Else
Me.SelctTeam.value = vbUnchecked
End If
Else
Me.SelctTeam.value = vbUnchecked
End If
If Not IsNull(rs("SelectProj1").value) Then
If (rs("SelectProj1").value) = 1 Then
Me.SelectProject.value = vbChecked
Else
Me.SelectProject.value = vbUnchecked
End If
Else
Me.SelectProject.value = vbUnchecked
End If
If Not IsNull(rs("SelectEmp").value) Then
If (rs("SelectEmp").value) = 1 Then
Me.Option2.value = True
Else
Me.Option2.value = False
End If
Else
Me.Option2.value = False
End If
If Not IsNull(rs("SelectAll").value) Then
If (rs("SelectAll").value) = 1 Then
Me.Option1.value = True
Else
Me.Option1.value = False
End If
Else
Me.Option1.value = False
End If
''//////////
    DCproject.BoundText = IIf(IsNull(rs("Project_id").value), "", rs("Project_id").value)
   ' Dcterm.BoundText = IIf(IsNull(rs("term_Fullcode").value), "", rs("term_Fullcode").value)
   ' dcopr.BoundText = IIf(IsNull(rs("opr_Fullcode").value), "", rs("opr_Fullcode").value)

    txtType.Text = IIf(IsNull(rs("opr_type").value), 0, rs("opr_type").value)
     CboYear.ListIndex = IIf(IsNull(rs("Years").value), -1, rs("Years").value)
    CmbMonth.ListIndex = IIf(IsNull(rs("Months").value), -1, rs("Months").value)

    If IsNull(rs("auto").value) Then
        ChKauto.value = vbUnchecked
    Else
        ChKauto.value = vbChecked
    End If
  ''// 01 06 2015
    Dcterm.BoundText = IIf(IsNull(rs("PandID").value), "", rs("PandID").value)
    dcopr.BoundText = IIf(IsNull(rs("OpraID").value), "", rs("OpraID").value)
 
    Fromdate.value = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
    todate.value = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
 

    
   
          If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
      XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
     
  StrSQL = "SELECT     dbo.projects.Project_name, dbo.opr_employee_details.*, dbo.projects_des.des, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, "
  StrSQL = StrSQL & "                     dbo.TblProcessDEF.TblProcessDEFID, dbo.projects_des.oprid, dbo.TblEmployee.Fullcode, dbo.opr_employee_details.ContProjSalar, dbo.projects.Project_nameE,"
  StrSQL = StrSQL & "                    dbo.projects.Fullcode AS ProFullcode, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.opr_employee_details.JobTypeID,"
  StrSQL = StrSQL & "                    dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee"
  StrSQL = StrSQL & " FROM         dbo.opr_employee_details LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmpJobsTypes ON dbo.opr_employee_details.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.projects_des ON dbo.opr_employee_details.PandID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblProcessDEF ON dbo.opr_employee_details.OperID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id"
StrSQL = StrSQL & " Where (dbo.opr_employee_details.pk_id =" & XPTxtID & ")"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_id").value), "", RsDev("Emp_id").value)
                 .TextMatrix(i, .ColIndex("ContProjSalar")) = IIf(IsNull(RsDev("ContProjSalar").value), 1, RsDev("ContProjSalar").value)
                .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("fullcode").value), "", RsDev("fullcode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(RsDev("emp_name").value), "", RsDev("emp_name").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
               Else
                      .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
               End If
                .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(RsDev("NumEkama").value), "", RsDev("NumEkama").value)
           '     .TextMatrix(i, .ColIndex("Start_date")) = IIf(IsNull(RsDev("Start_date").value), XPDtbTrans.value, RsDev("Start_date").value)
           '     If IsNull(RsDev("toid").value) Then
             '            .TextMatrix(i, .ColIndex("end_date")) = ""
              '               .TextMatrix(i, .ColIndex("toid")) = ""
              '               .TextMatrix(i, .ColIndex("interval")) = ""
                    
              '      Else
                   
              '               .TextMatrix(i, .ColIndex("toid")) = IIf(IsNull(RsDev("toid").value), "", RsDev("toid").value)
                             .TextMatrix(i, .ColIndex("interval")) = IIf(IsNull(RsDev("interval").value), "", RsDev("interval").value)
                    
                
              '      End If
     
              .TextMatrix(i, .ColIndex("SpecificationID")) = IIf(IsNull(RsDev("SpecificationID").value), "", RsDev("SpecificationID").value)
              .TextMatrix(i, .ColIndex("project_id")) = IIf(IsNull(RsDev("project_id").value), "", RsDev("project_id").value)
              .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDev("BranchId").value), "", RsDev("BranchId").value)
              .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(RsDev("DepartmentID").value), "", RsDev("DepartmentID").value)
              .TextMatrix(i, .ColIndex("JobTypeID")) = IIf(IsNull(RsDev("JobTypeID").value), "", RsDev("JobTypeID").value)
              .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(RsDev("ProjectID").value), "", RsDev("ProjectID").value)
              .TextMatrix(i, .ColIndex("PrjectCode")) = IIf(IsNull(RsDev("ProFullcode").value), "", RsDev("ProFullcode").value)
            
                   .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(RsDev("FromDate").value), "", RsDev("FromDate").value)
                    .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(RsDev("ToDate").value), "", RsDev("ToDate").value)
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    .TextMatrix(i, .ColIndex("PandID")) = IIf(IsNull(RsDev("oprid").value), "", RsDev("oprid").value)
                    .TextMatrix(i, .ColIndex("OperID")) = IIf(IsNull(RsDev("TblProcessDEFID").value), "", RsDev("TblProcessDEFID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(RsDev("Project_name").value), "", RsDev("Project_name").value)
                    .TextMatrix(i, .ColIndex("opra")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                     Else
                     .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                     .TextMatrix(i, .ColIndex("opra")) = IIf(IsNull(RsDev("ProcessNameE").value), "", RsDev("ProcessNameE").value)
                     End If
                
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RetriveProjectSalar
 fillapprovData
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub ToDate_Change()
If Me.Fromdate.value <> "" Then

If Not IsDate(todate) Then
     mToDate = MonthLastDay(Fromdate)
Else
    mToDate = todate.value
End If

'NoDay = DateDiff("d", Me.FromDate.value, mToDate)
GetNoOfDays Me.Fromdate.value, mToDate
If NoDay > 30 Then NoDay = 30
Me.txtDays.Text = NoDay

End If
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        CmdRemoveAll.Enabled = True
        ELe(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        CmdRemoveAll.Enabled = True
        ELe(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
       ' Ele(1).Enabled = False

        CmdRemove.Enabled = False
        CmdRemoveAll.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

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

Private Sub XPDtbTrans_Change()
ReLineGrid
End Sub
