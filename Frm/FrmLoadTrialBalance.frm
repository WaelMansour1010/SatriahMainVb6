VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmLoadTrialBalance 
   Caption         =   "«” œ⁄«¡ „Ì“«‰ „—«Ã⁄Â"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
   HelpContextID   =   450
   Icon            =   "FrmLoadTrialBalance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15735
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
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9435
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15735
      _cx             =   27755
      _cy             =   16642
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   6045
         Left            =   15
         TabIndex        =   1
         Top             =   1995
         Width           =   15720
         _cx             =   27728
         _cy             =   10663
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
         FrontTabForeColor=   -2147483630
         Caption         =   "«·ﬁÌÊœ|«·‘—Õ «·⁄«„"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   6
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
            Height          =   5955
            Index           =   0
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   45
            Width           =   14760
            _cx             =   26035
            _cy             =   10504
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
            BackColor       =   16777215
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   2
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
            GridRows        =   2
            GridCols        =   4
            Frame           =   1
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmLoadTrialBalance.frx":030A
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic EleOpt 
               Height          =   945
               Left            =   30
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   4980
               Width           =   11025
               _cx             =   19447
               _cy             =   1667
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
               ForeColorDisabled=   -2147483630
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   990
                  Left            =   -60
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   11610
                  _cx             =   20479
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
                  Begin VB.Frame Frame2 
                     Caption         =   "ÿ—Ìﬁ… ⁄—÷ «·Õ”«»« "
                     Height          =   735
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   136
                     Top             =   120
                     Width           =   4455
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "ÃœÊ·Ì"
                        Height          =   345
                        Index           =   2
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   139
                        Top             =   240
                        Width           =   1260
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "„”«—"
                        Height          =   345
                        Index           =   1
                        Left            =   1680
                        RightToLeft     =   -1  'True
                        TabIndex        =   138
                        Top             =   240
                        Width           =   1140
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "‘Ã—Ì"
                        Height          =   345
                        Index           =   0
                        Left            =   3120
                        RightToLeft     =   -1  'True
                        TabIndex        =   137
                        Top             =   240
                        Width           =   1140
                     End
                  End
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     Caption         =   "  — Ì» »«·œ·Ì· «·„Õ«”»Ì"
                     Height          =   345
                     Index           =   1
                     Left            =   7395
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   615
                     Width           =   2940
                  End
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     Caption         =   "  — Ì» «»ÃœÏ"
                     Height          =   345
                     Index           =   0
                     Left            =   7395
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   255
                     Width           =   2940
                  End
                  Begin ALLButtonS.ALLButton CmdRemove 
                     Height          =   420
                     Left            =   525
                     TabIndex        =   77
                     Tag             =   "Delete Row"
                     Top             =   0
                     Width           =   1140
                     _ExtentX        =   2011
                     _ExtentY        =   741
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
                     MICON           =   "FrmLoadTrialBalance.frx":0375
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton cmdAdd 
                     Height          =   420
                     Left            =   480
                     TabIndex        =   124
                     Tag             =   "Delete Row"
                     Top             =   480
                     Width           =   1140
                     _ExtentX        =   2011
                     _ExtentY        =   741
                     BTYPE           =   3
                     TX              =   "«œ—«Ã ”ÿ—"
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
                     BCOL            =   65280
                     BCOLO           =   65280
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmLoadTrialBalance.frx":0391
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   " — Ì» «·Õ”«»« "
                     Height          =   375
                     Index           =   12
                     Left            =   8415
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   0
                     Width           =   1890
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                  Height          =   870
                  Left            =   11445
                  TabIndex        =   61
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   12285
                  _cx             =   21669
                  _cy             =   1535
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
                  Begin VB.OptionButton Optx 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰Ÿ«„ «·„”«—"
                     Height          =   270
                     Index           =   1
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   1575
                  End
                  Begin VB.OptionButton Optx 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·‰Ÿ«„ «·‘Ã—Ï"
                     Height          =   270
                     Index           =   0
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   360
                     Width           =   1575
                  End
                  Begin VB.OptionButton Optx 
                     Alignment       =   1  'Right Justify
                     Caption         =   "⁄—÷ ÃœÊ·Ï"
                     Height          =   285
                     Index           =   2
                     Left            =   5000
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   360
                     Width           =   1455
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "⁄—÷ «·Õ”«»« "
                     Height          =   240
                     Index           =   11
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   120
                     Width           =   1965
                  End
               End
               Begin C1SizerLibCtl.C1Elastic EleSortOpt 
                  Height          =   555
                  Left            =   26820
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   13125
                  _cx             =   23151
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
                  Begin VB.OptionButton OptSort 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " — Ì» »«·œ·Ì· «·„Õ«”»Ï"
                     Height          =   195
                     Index           =   11
                     Left            =   -2460
                     RightToLeft     =   -1  'True
                     TabIndex        =   3
                     Top             =   -90
                     Value           =   -1  'True
                     Width           =   46710
                  End
               End
               Begin VB.Image ImgNote 
                  Height          =   240
                  Left            =   0
                  Picture         =   "FrmLoadTrialBalance.frx":03AD
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   4920
               Left            =   30
               TabIndex        =   2
               Top             =   30
               Width           =   14700
               _cx             =   25929
               _cy             =   8678
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
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   10
               Rows            =   10
               Cols            =   28
               FixedRows       =   2
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmLoadTrialBalance.frx":0937
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
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   4035
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   4035
                  ScaleWidth      =   9405
                  TabIndex        =   10
                  Top             =   810
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.TextBox TxtDese 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1485
                     Left            =   120
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   29
                     Top             =   2040
                     Width           =   8955
                  End
                  Begin VB.TextBox txtcodesub 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   4920
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   3600
                     Width           =   855
                  End
                  Begin VB.CommandButton Command4 
                     Caption         =   "«÷«›… ‘—Õ"
                     Height          =   375
                     Left            =   7440
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   3600
                     Width           =   1350
                  End
                  Begin VB.CommandButton Command3 
                     Caption         =   "«” œ⁄«¡ ‘—Õ"
                     Height          =   375
                     Left            =   6240
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   3600
                     Width           =   1095
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                     Height          =   4620
                     Left            =   120
                     TabIndex        =   25
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   10905
                     _cx             =   19235
                     _cy             =   8149
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial (Arabic)"
                        Size            =   20.25
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
                     BackColor       =   -2147483633
                     ForeColor       =   4210688
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
                     CaptionPos      =   7
                     WordWrap        =   -1  'True
                     MaxChildSize    =   0
                     MinChildSize    =   0
                     TagWidth        =   0
                     TagPosition     =   0
                     Style           =   0
                     TagSplit        =   2
                     PicturePos      =   7
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
                     Begin VB.TextBox TxtDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1380
                        Left            =   0
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   28
                        Top             =   525
                        Visible         =   0   'False
                        Width           =   8955
                     End
                     Begin VB.Shape Shape3 
                        Height          =   375
                        Left            =   360
                        Top             =   3600
                        Width           =   1695
                     End
                     Begin VB.Label Label10 
                        Alignment       =   1  'Right Justify
                        Caption         =   "„Ê«›ﬁ"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   12
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   420
                        Left            =   600
                        RightToLeft     =   -1  'True
                        TabIndex        =   31
                        Top             =   3600
                        Width           =   855
                     End
                     Begin VB.Label LblDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000C&
                        Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                        ForeColor       =   &H0000C8FF&
                        Height          =   360
                        Left            =   6840
                        RightToLeft     =   -1  'True
                        TabIndex        =   26
                        Top             =   0
                        Width           =   2445
                     End
                  End
                  Begin VB.Shape Shape2 
                     BorderWidth     =   5
                     FillStyle       =   2  'Horizontal Line
                     Height          =   375
                     Left            =   480
                     Top             =   3480
                     Width           =   3255
                  End
                  Begin VB.Shape Shape1 
                     Height          =   495
                     Left            =   360
                     Top             =   3480
                     Width           =   1335
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Code"
                     Height          =   495
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   3480
                     Width           =   735
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     Height          =   495
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   1200
                     Width           =   975
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Code"
                     Height          =   255
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   1320
                     Width           =   735
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   9
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   2475
                  _cx             =   1973752078
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmLoadTrialBalance.frx":0DE0
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   113
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   945
               Left            =   30
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   4980
               Width           =   14700
               _cx             =   25929
               _cy             =   1667
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   20.25
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
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   6
               ChildSpacing    =   4
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
               PicturePos      =   7
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
               Begin VB.CheckBox ChkLastAccount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "⁄—÷ «·Õ”«»«  «·‰Â«∆Ì… ›ﬁÿ"
                  Height          =   195
                  Left            =   26145
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   600
                  Value           =   1  'Checked
                  Width           =   7770
               End
               Begin DBPIXLib.DBPix20 DBPix202 
                  Height          =   615
                  Left            =   26145
                  TabIndex        =   15
                  Top             =   -120
                  Width           =   5340
                  _Version        =   131072
                  _ExtentX        =   9419
                  _ExtentY        =   1085
                  _StockProps     =   1
                  BackColor       =   16777215
                  _Image          =   "FrmLoadTrialBalance.frx":137A
                  ImageResampleWidth=   100
                  ImageResampleHeight=   100
                  ImageResampleMode=   1
                  ImageSaveFormat =   0
                  JPEGQuality     =   75
                  JPEGEncoding    =   0
                  JPEGColorMode   =   0
                  JPEGNoRecompress=   -1  'True
                  JPEGRotateWarning=   0
                  PNGColorDepth   =   0
                  PNGCompression  =   0
                  PNGFilter       =   0
                  PNGInterlace    =   1
                  ImageDitherMethod=   3
                  ImagePaletteMethod=   4
                  ImagePreviewMode=   0   'False
                  ImageKeepMetaData=   -1  'True
                  UseAmbientBackcolor=   -1  'True
                  ViewAsyncDecoding=   -1  'True
                  ViewEnableMouseZoom=   -1  'True
                  ViewInitialZoom =   0
                  ViewHAlign      =   1
                  ViewVAlign      =   1
                  ViewMenuMode    =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· ÊﬁÌ⁄"
                  Height          =   240
                  Index           =   5
                  Left            =   31080
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Tag             =   "51"
                  Top             =   0
                  Width           =   2625
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5955
            Index           =   1
            Left            =   16365
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   14760
            _cx             =   26035
            _cy             =   10504
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
            Begin VB.TextBox Txte 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   1590
               Left            =   4440
               MaxLength       =   1000
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   129
               Top             =   2520
               Width           =   10095
            End
            Begin VB.Frame Frame3 
               Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
               Height          =   675
               Left            =   14970
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   4935
               Visible         =   0   'False
               Width           =   8955
            End
            Begin VB.TextBox Txtcode 
               Alignment       =   1  'Right Justify
               Height          =   480
               Left            =   31425
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   5355
               Width           =   3645
            End
            Begin VB.CommandButton Command2 
               Caption         =   "«” œ⁄«¡ ﬁ«·» ‘—Õ"
               Height          =   630
               Left            =   4500
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   4155
               Width           =   3525
            End
            Begin VB.CommandButton Command1 
               Caption         =   "«÷«›… ﬁ«·» ‘—Õ"
               Height          =   690
               Left            =   8205
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   4110
               Width           =   3480
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   1590
               Left            =   4455
               MaxLength       =   1000
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   420
               Width           =   10095
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‘—Õ «‰Ã·Ì“Ì"
               Height          =   255
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‘—Õ ⁄—»Ì"
               Height          =   255
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
               Height          =   315
               Left            =   18270
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   3060
               Width           =   5565
            End
            Begin VB.Label Lb_note_value_by_characters 
               Alignment       =   1  'Right Justify
               Height          =   660
               Left            =   16215
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   7545
               Width           =   19485
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Code"
               Height          =   615
               Left            =   27495
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   5355
               Width           =   3300
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄·Ìﬁ:"
               Height          =   195
               Index           =   6
               Left            =   27975
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Tag             =   "22"
               Top             =   570
               Width           =   7425
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   1365
         Left            =   15
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   615
         Width           =   15705
         _cx             =   27702
         _cy             =   2408
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
         Begin VB.CommandButton CmdImport 
            Caption         =   "«” Ì—«œ «·„·›"
            Height          =   375
            Left            =   8640
            TabIndex        =   143
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   3960
            TabIndex        =   141
            Top             =   105
            Width           =   7455
         End
         Begin VB.CommandButton CMDSelectFile 
            Caption         =   "Õœœ «·„·›"
            Height          =   375
            Left            =   10080
            TabIndex        =   140
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtManualNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   17160
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   133
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "⁄—÷ "
            Height          =   285
            Left            =   9000
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   -3705
            Width           =   615
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   285
            Left            =   12600
            TabIndex        =   130
            Top             =   480
            Visible         =   0   'False
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   503
         End
         Begin VB.Frame Frame1 
            Caption         =   "‰”Œ ﬁÌœ ”«»ﬁ"
            Height          =   735
            Left            =   -15
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   -2100
            Width           =   3345
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   240
               Width           =   2205
            End
            Begin VB.CommandButton Command5 
               Caption         =   "‰”Œ"
               Height          =   330
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„Â"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   3600
               TabIndex        =   83
               Top             =   240
               Width           =   435
            End
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ‰"
            Height          =   165
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   -1425
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«∆‰"
            Height          =   165
            Left            =   3450
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   -1695
            Width           =   990
         End
         Begin VB.CheckBox Auto_cost_center 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Ê“Ì⁄ «·Ï ·„—«ﬂ“ «· ﬂ·›…"
            Enabled         =   0   'False
            Height          =   270
            Left            =   165
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   -1860
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9720
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   -3705
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9030
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   -4095
            Width           =   2415
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   12555
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   105
            Width           =   1770
         End
         Begin VB.TextBox TxtDEVID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   600
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.TextBox TxtValue 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   510
            Left            =   8130
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1470
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   495
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   105
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.PictureBox DtHijriTrans 
            BackColor       =   &H00FFFFC0&
            Height          =   270
            Left            =   3285
            ScaleHeight     =   210
            ScaleWidth      =   1740
            TabIndex        =   33
            Top             =   -1335
            Visible         =   0   'False
            Width           =   1800
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   12600
            Top             =   1560
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic ElePost 
            Height          =   270
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1500
            Visible         =   0   'False
            Width           =   4020
            _cx             =   7091
            _cy             =   476
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
            ForeColorDisabled=   -2147483630
            Caption         =   "Õ«·… «·”‰œ"
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   2
            ChildSpacing    =   1
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
            Frame           =   4
            FrameStyle      =   3
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CheckBox ChkPost 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·”‰œ"
               Height          =   225
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   45
               Width           =   1485
            End
            Begin VB.Image Img 
               Height          =   225
               Index           =   0
               Left            =   90
               Top             =   90
               Width           =   270
            End
            Begin VB.Image Img 
               Height          =   180
               Index           =   1
               Left            =   1635
               Top             =   285
               Width           =   285
            End
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmLoadTrialBalance.frx":1392
            Height          =   315
            Left            =   105
            TabIndex        =   44
            Top             =   -1455
            Width           =   3180
            _ExtentX        =   5609
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
         Begin MSDataListLib.DataCombo dcprojects 
            Bindings        =   "FrmLoadTrialBalance.frx":13A7
            Height          =   315
            Left            =   4575
            TabIndex        =   45
            Top             =   -585
            Width           =   2835
            _ExtentX        =   5001
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
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   285
            Left            =   12600
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   495
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98238467
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   4575
            TabIndex        =   78
            Top             =   -1935
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   3120
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„·›"
            Height          =   210
            Index           =   15
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Tag             =   "53"
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   315
            Index           =   14
            Left            =   18960
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Tag             =   "53"
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Lblnotes_all 
            Alignment       =   1  'Right Justify
            Caption         =   "Label17"
            Height          =   15
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label LblTransactionsId 
            Alignment       =   1  'Right Justify
            Caption         =   "Label17"
            Height          =   255
            Left            =   16440
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ «·⁄«„"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   7440
            TabIndex        =   79
            Top             =   -735
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   315
            Index           =   10
            Left            =   15615
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Tag             =   "53"
            Top             =   645
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1005
            Width           =   2175
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            Height          =   225
            Index           =   9
            Left            =   14325
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Tag             =   "53"
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   4860
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1005
            Width           =   2070
         End
         Begin VB.Label lblPost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   1005
            Width           =   2265
         End
         Begin VB.Label LblKaleb 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   9750
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1005
            Width           =   2250
         End
         Begin VB.Label LblDawry 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1005
            Width           =   2160
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄ «·⁄«„"
            Height          =   225
            Left            =   7410
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   -585
            Width           =   1350
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
            Height          =   210
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   -1935
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   225
            Left            =   11325
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   -3705
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„’œ—Â"
            Height          =   210
            Left            =   11565
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   -4095
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   210
            Index           =   3
            Left            =   14325
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Tag             =   "53"
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·”‰œ"
            Height          =   510
            Index           =   4
            Left            =   10455
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Tag             =   "54"
            Top             =   1470
            Visible         =   0   'False
            Width           =   1770
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleTop 
         Height          =   660
         Left            =   0
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   0
         Width           =   15735
         _cx             =   27755
         _cy             =   1164
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         BackColor       =   12648447
         ForeColor       =   8421376
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "«” œ⁄«¡ „Ì“«‰ „—«Ã⁄Â"
         Align           =   1
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
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
         PicturePos      =   7
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
         Begin VB.TextBox oldTxtSerial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1380
            TabIndex        =   73
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmLoadTrialBalance.frx":13BC
            ColorButton     =   12648447
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
            Left            =   135
            TabIndex        =   74
            Top             =   120
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmLoadTrialBalance.frx":1756
            ColorButton     =   12648447
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
            Left            =   1875
            TabIndex        =   75
            Top             =   120
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmLoadTrialBalance.frx":1AF0
            ColorButton     =   12648447
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
            Left            =   675
            TabIndex        =   76
            Top             =   120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmLoadTrialBalance.frx":1E8A
            ColorButton     =   12648447
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
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
            Height          =   315
            Index           =   27
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   120
            Width           =   7155
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   1230
         Left            =   0
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   8160
         Width           =   15675
         _cx             =   27649
         _cy             =   2170
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
         Begin VB.TextBox TXTResults 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   120
            Width           =   1890
         End
         Begin VB.Frame Frame17 
            Height          =   510
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   6660
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.TextBox TxtDEV_NO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   16560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   120
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               Caption         =   "„·€Ì"
               Height          =   195
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   120
               Width           =   1095
            End
            Begin VB.TextBox txt_salary 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   16440
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Text            =   "Text1"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œÌ„ «· √ÀÌ—"
               Height          =   315
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   120
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «⁄ „«œÂ"
               Height          =   315
               Left            =   1020
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   120
               Width           =   1335
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁ«·»"
               Height          =   315
               Left            =   -480
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   120
               Width           =   1095
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌœ œÊ—Ì"
               Height          =   195
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   120
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”·"
               Height          =   435
               Index           =   7
               Left            =   19050
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Tag             =   "57"
               Top             =   240
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«· «—ÌŒ"
               Height          =   180
               Index           =   0
               Left            =   14220
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Tag             =   "52"
               Top             =   555
               Width           =   915
            End
         End
         Begin VB.TextBox TxtTotalCredit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   5895
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   90
            Width           =   1890
         End
         Begin VB.TextBox TxtTotalDebit 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   9570
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   90
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Left            =   30
            TabIndex        =   99
            Top             =   90
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   12648447
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   0
            Left            =   13590
            TabIndex        =   100
            Top             =   480
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   1
            Left            =   12120
            TabIndex        =   101
            Top             =   435
            Width           =   1740
            _ExtentX        =   3069
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
            Height          =   315
            Index           =   2
            Left            =   10860
            TabIndex        =   102
            Top             =   480
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
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
            Height          =   330
            Index           =   3
            Left            =   9495
            TabIndex        =   103
            Top             =   435
            Width           =   1215
            _ExtentX        =   2143
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
            Height          =   330
            Index           =   4
            Left            =   5910
            TabIndex        =   104
            Top             =   435
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
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
            Height          =   330
            Index           =   5
            Left            =   4830
            TabIndex        =   105
            Top             =   435
            Width           =   990
            _ExtentX        =   1746
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
            Height          =   330
            Index           =   6
            Left            =   0
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   435
            Width           =   1275
            _ExtentX        =   2249
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
            Height          =   330
            Index           =   7
            Left            =   3045
            TabIndex        =   107
            Top             =   435
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
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
            Height          =   330
            Left            =   1500
            TabIndex        =   108
            Top             =   435
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   8
            Left            =   8040
            TabIndex        =   109
            Top             =   480
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
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
         Begin ALLButtonS.ALLButton ALLButton20 
            Height          =   345
            Left            =   11565
            TabIndex        =   110
            Top             =   795
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«⁄ „«œ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":2224
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton6 
            Height          =   345
            Left            =   9960
            TabIndex        =   111
            Top             =   795
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«‰‘«¡ ﬁÌœ œÊ—Ì"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":2240
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton7 
            Height          =   345
            Left            =   6615
            TabIndex        =   112
            Top             =   795
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   " ÕÊÌ· «·Ï ﬁ«·»"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":225C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton8 
            Height          =   345
            Left            =   3915
            TabIndex        =   113
            Top             =   795
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«·€«¡ «· √ÀÌ—"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":2278
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton9 
            Height          =   345
            Left            =   2505
            TabIndex        =   114
            Top             =   795
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«‰‘«¡ ﬁÌœ ⁄ﬂ”Ì"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65535
            BCOLO           =   65535
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":2294
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton10 
            Height          =   345
            Left            =   5145
            TabIndex        =   115
            Top             =   795
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«” œ⁄«¡ ﬁ«·»"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":22B0
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
            Height          =   345
            Left            =   13080
            TabIndex        =   116
            Top             =   795
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":22CC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   345
            Left            =   1320
            TabIndex        =   117
            Top             =   795
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«·„—›ﬁ« "
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":22E8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton3 
            Height          =   345
            Left            =   8385
            TabIndex        =   118
            Top             =   795
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«” œ⁄«¡ ﬁÌœ œÊ—Ï"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   65280
            BCOLO           =   65280
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLoadTrialBalance.frx":2304
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—ﬁ"
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   13
            Left            =   5115
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Tag             =   "56"
            Top             =   165
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ…"
            Height          =   225
            Index           =   8
            Left            =   2130
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Tag             =   "51"
            Top             =   135
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·œ«∆‰"
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   2
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Tag             =   "56"
            Top             =   135
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ÿ—› «·„œÌ‰"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   1
            Left            =   11505
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Tag             =   "55"
            Top             =   90
            Width           =   1980
         End
      End
   End
End
Attribute VB_Name = "FrmLoadTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim line_no1 As Integer
Dim last_line_id As Double
Dim numbering_type As Integer
Dim TTP As New clstooltip
Dim BolEditOnMainAccounts As Boolean
Dim PicHeight As Long
Dim PicWidth As Long
Dim Dcombos As ClsDataCombos
Dim DCboSearch As New clsDCboSearch
Dim rs As New ADODB.Recordset
Private Enum PrintTarget
    WindowTarget
    PrinterTarget
End Enum

Public StrOldTransID As String
  
Function sand_numbering() As String
    'On Error Resume Next
    'Dim start_at As Integer
    'Dim end_at As Integer
    'Dim auto_sanad_no As String
    'auto_sanad_no = ""
    'departement_name = 1
    'branch_no = 1
    'connection_string = Cn.ConnectionString
    'numbering.ConnectionString = connection_string
    'numbering.CommandType = adCmdText
    'numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=0"
    'numbering.Refresh
    'If numbering.Recordset.RecordCount = 0 Then
    'numbering_type = 0
    'Else
    'numbering_type = numbering.Recordset.Fields!numbering_id
    'start_at = numbering.Recordset.Fields!start_at
    'end_at = numbering.Recordset.Fields!end_at
    '
    'End If

    'If numbering_type = 1 Then
    'detect_no.ConnectionString = connection_string
    'detect_no.CommandType = adCmdText
    'detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
    'detect_no.Refresh
    '
    ' If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
    '
    ' If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
    '
    ' If detect_no.Recordset.Fields!last_sand_no >= end_at Then
    ' sand_numbering = "error"
    ' Exit Function
    ' End If
    ' End If
    'Else
    'If numbering_type = 2 Then
 
    'detect_no.ConnectionString = connection_string
    'detect_no.CommandType = adCmdText
    'detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 4, 2)

    'detect_no.Refresh
    '
    'If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
    '   no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
    '   If end_at = 0 Then end_at = no + 1
    ' If no >= end_at Then
    ' sand_numbering = "error"
    ' Exit Function
    ' End If
    ' End If

    'Else
    'If numbering_type = 3 Then
    '
    'detect_no.ConnectionString = connection_string
    'detect_no.CommandType = adCmdText
    'detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=200 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4)

    'detect_no.Refresh
    'If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
    'no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
    'If end_at = 0 Then end_at = no + 1
    ' If no >= end_at Then
    ' sand_numbering = "error"
    ' Exit Function
    ' End If
    ' End If
 
    'End If
 
    'End If
    'End If

    'If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

    '                If numbering_type = 0 Then

    '                Else
    '                    If numbering_type = 1 Then
    '                    auto_sanad_no = start_at
    '                Else
                
    '                    If numbering_type = 2 Then
    '                    auto_sanad_no = Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 4, 2) & start_at
    '
    '                Else
    '                     If numbering_type = 3 Then
    '                        auto_sanad_no = Mid(Format$(DTP_Date.value, "dd/mm/yyyy"), 7, 4) & start_at
    '
    '                  End If
    '                  End If
    '                  End If
    '                  End If
    '
    'Else
    '                If numbering_type = 0 Then

    '                Else
    '                    If numbering_type = 1 Then
    '                  auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
    '                Else
                
    '                    If numbering_type = 2 Then
    '                     no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
    '                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (no + 1)
    '
                      
    '                Else
    '                     If numbering_type = 3 Then
    '                           no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
    '                     auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (no + 1)

    '                  End If
    '                  End If
    '                  End If
    '                  End If
    '
    'End If
    'sand_numbering = auto_sanad_no

End Function

Private Sub ALLButton1_Click()
    'On Error GoTo ErrTrap

    On Error Resume Next

    If DcCostCenter.BoundText <> "" Then

        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Exit Sub
    End If

    If Me.Auto_cost_center.value = vbUnchecked Then
        Dim StrSQL As String
        StrSQL = "Delete  marakes_taklefa_temp  where auto_des=1 and  kedno =" & val(Text1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    
    End If

    If DcCostCenter.BoundText <> "" Then

        'MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        'Exit Sub
    End If

    If DcCostCenter.BoundText <> "" Then
        StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 and kedno =" & val(Text1.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = Text1.Text
    'Else
    'opr_id = TxtDEV_NO.text
    'End If
    Unload marakes_taklefa_tawze3

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "0" Then
            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            
            marakes_taklefa_tawze3.txtAccountSerial = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("Account_Serial"))
            
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        Else

            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "0" Then
                marakes_taklefa_tawze3.show
            
                marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) 'Text5.Text
                marakes_taklefa_tawze3.depit_or_credit.Caption = "œ«∆‰"
                marakes_taklefa_tawze3.kedno = opr_id
                    
                marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
                marakes_taklefa_tawze3.txtAccountSerial = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("Account_Serial"))
                marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
                marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
              
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„…  ", vbCritical
                Else
                    MsgBox "Enter Value First ", vbCritical
                End If

                Exit Sub
            End If
        End If

        marakes_taklefa_tawze3.opr_type = "”‰œ ﬁÌœ"
        marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
        marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
        marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
        marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        marakes_taklefa_tawze3.Adodc3.Refresh
        Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    End If

    If Me.TxtModFlg.Text = "R" Then
        marakes_taklefa_tawze3.ALLButton1.Enabled = False
    Else
        marakes_taklefa_tawze3.ALLButton1.Enabled = True
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub ALLButton10_Click()

    If SystemOptions.UserInterface = ArabicInterface Then

        If Me.TxtModFlg.Text <> "N" Then MsgBox "·«»œ „‰ «·÷€ÿ ⁄·Ï ÃœÌœ «Ê·« ·«” œ⁄«¡ «·ﬁ«·» ": Exit Sub
    Else

        If Me.TxtModFlg.Text <> "N" Then MsgBox "Press New To Call Templates": Exit Sub
    End If
  
    'If Fg_Journal.Rows > 4 Then MsgBox "ÌÊÃœ «”ÿ— ›Ì Â–« «·ﬁÌœ ·–·ﬂ ·«Ì„ﬂ‰ «” œ⁄«¡ ﬁ«·» «·ﬁÌœ": Exit Sub

    Unload KALEB
    KALEB.show
End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    If TxtSerial.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
         
            MsgBox "·«»œ „‰ «Õ Ì«— ﬁÌœ «Ê·«": Exit Sub
        Else
            MsgBox "Select Voucher Firstly": Exit Sub
        End If
 
    End If

    Unload imaged
    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Voucher #"
        imaged.Caption = "Voucher Attachment"
        imaged.txtopeation_type = "„—›ﬁ«  «·ﬁÌœ"
        imaged.SUBJECT_NO = TxtSerial.Text
        imaged.Label6.Caption = "Voucher #"
    Else

        imaged.Label9.Caption = "„—›ﬁ«  ”‰œ ﬁÌœ  —ﬁ„"
        imaged.Caption = "„—›ﬁ«  «·ﬁÌœ  "
        imaged.txtopeation_type = "„—›ﬁ«  «·ﬁÌœ"
        imaged.SUBJECT_NO = TxtSerial.Text
        imaged.Label6.Caption = "—ﬁ„  «·ﬁÌœ"

    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·ﬁÌœ' and subject_no='" & TxtSerial.Text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub ALLButton20_Click()
 
    If Dir(App.path & "\images\sign" & user_id & ".JPG") <> "" Then
        DBPix202.ImageLoadFile (App.path & "\images\sign" & user_id & ".JPG")
   
        Check2.value = 1

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« ÌÕﬁ ·Â–« «·„” Œœ„ «⁄ „«œ «·”‰œ« "
        Else
            MsgBox "Not allow to do this"
        End If

    End If

End Sub

Private Sub ALLButton3_Click()

    If SystemOptions.UserInterface = ArabicInterface Then
        If Me.TxtModFlg.Text <> "N" Then MsgBox "·«»œ „‰ «·÷€ÿ ⁄·Ï ÃœÌœ «Ê·« ·«’œ«— «·ﬁÌœ «·œÊ—Ì": Exit Sub
    Else

        If Me.TxtModFlg.Text <> "N" Then MsgBox " Press New To Create Repeated Voucher": Exit Sub
    End If

    Unload keddawrym
    keddawrym.show

End Sub

Private Sub ALLButton6_Click()

    'If Me.TxtModFlg.text <> "E" And Me.TxtModFlg.text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    If TxtSerial.Text = "" Then MsgBox "«Œ — ﬁÌœ «Ê·«", vbCritical: Exit Sub
    If Text2.Text <> "Manual" And Text2.Text <> "ÌœÊÌ" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰ «‰‘«¡ ﬁÌœ œÊ—Ì „‰ ﬁÌœ «·Ì", vbCritical
        Else
            MsgBox "Can't create Repeated Voucher Form Auto vouchers"
        
        End If

        Exit Sub
    End If

    Unload ked_dawry
    ked_dawry.show
    ked_dawry.ID = Me.TXTNoteID ' TxtDEV_NO.text
    ked_dawry.desc = Txt.Text
    ked_dawry.TxtSerial = Me.TxtSerial
    Check4.value = vbChecked
End Sub

Private Sub ALLButton7_Click()
Cmd_Click (1)
    If SystemOptions.UserInterface = ArabicInterface Then
        If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    Else

        If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "Press New Or Edit Firstly", vbCritical: Exit Sub
    End If

    If Text2.Text <> "Manual" And Text2.Text <> "ÌœÊÌ" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰  ÕÊÌ· ﬁÌœ «·Ì «·Ï ﬁ«·»", vbCritical
        Else
            MsgBox "Can't convert Auto Voucher To Template"
        
        End If

        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox(" √ﬂÌœ «· ÕÊÌ· «·Ï ﬁ«·»", vbInformation + vbYesNo)
    Else
        X = MsgBox("Confirm Convert To Template?", vbInformation + vbYesNo)
    End If

    If X = vbYes Then
        Check3.value = 1
    End If
Cmd_Click (2)
End Sub

Private Sub ALLButton8_Click()

    If SystemOptions.UserInterface = ArabicInterface Then
        If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "«÷€ÿ  ⁄œÌ·  «Ê ÃœÌœ «Ê·«", vbCritical: Exit Sub
    Else

        If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then MsgBox "press New or modify Firstlty to do that", vbCritical: Exit Sub
    End If

    If Check1.value = vbChecked Then
        Check1.value = 1
        Check1.value = Unchecked
    Else
        Check1.value = vbChecked
    End If

End Sub

Private Sub ALLButton9_Click()

    If SystemOptions.UserInterface = ArabicInterface Then
        Text2.Text = "ÌœÊÌ"
    Else
        Text2.Text = "Manual"
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Text3.Text = "ﬁÌœ ÌÊ„Ì…"
    Else
        Text3.Text = "JL Entry"
    End If

    Me.Txt.Text = ""
    Check1.value = Unchecked
    Check2.value = Unchecked
    Check3.value = Unchecked
    Check4.value = Unchecked
    Check5.value = Unchecked

    Me.TXTNoteID.Text = ""
    Me.TxtDEVID.Text = ""
    Me.DTP_Date.value = Date
    Me.TxtSerial.Text = ""
    Me.TxtValue.Text = ""

    Me.ChkPost.value = vbUnchecked

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.ChkPost.Caption = "€Ì— „—Õ·"
    Else
        Me.ChkPost.Caption = "Not Poasted"
    End If

    Me.ChkPost.ForeColor = vbBlack
    'Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Me.DcboUsers.BoundText = user_id
 
    Me.TxtModFlg.Text = "N"
    setfoxy
    DcCostCenter.Text = ""
    Option1.value = True
    Dim temp_value As Double

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Single

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) > 0 Then
                
                    .TextMatrix(i, .ColIndex("CreditValue")) = val(.TextMatrix(i, .ColIndex("DebitValue")))
                    .TextMatrix(i, .ColIndex("DebitValue")) = 0
                Else
                    .TextMatrix(i, .ColIndex("DebitValue")) = val(.TextMatrix(i, .ColIndex("CreditValue")))
                    .TextMatrix(i, .ColIndex("CreditValue")) = 0
                End If
            
            End If

        Next i

    End With

End Sub

Private Sub Auto_cost_center_Click()
   
    If Auto_cost_center.value = vbUnchecked Then
        ALLButton1.Enabled = True
    Else
    
        ALLButton1.Enabled = False
    End If

End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                '    PicDes.Height = PicHeight
                '    PicDes.Width = PicWidth
            Else
                '    PicDes.Width = CboDes.Width - 10
                '    PicDes.Height = CboDes.Height * 8
            End If

            '  Debug.Print PicHeight
            '  Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            TxtDese.Visible = True
            TxtDese.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("dese")) ' Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Dese"))
            TxtDes.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("des"))
            TxtDese.Text = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("dese"))
    
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            '  Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 200

    End If

End Sub

Private Sub CboDes_KeyUp(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 200

    End If

End Sub

Private Sub ChkPost_Click()

    'Stop
    If ChkPost.value = vbChecked Then
        img(1).Visible = True
        img(0).Visible = False
        ChkPost.ForeColor = vbRed
    ElseIf ChkPost.value = vbUnchecked Then
        img(0).Visible = True
        img(1).Visible = False
        ChkPost.ForeColor = vbBlack
    End If

End Sub

Function setfoxy_Line() As Double
    
    last_line_id = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = last_line_id
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = last_line_id
 
    rs.update
    
End Function

Function setfoxy()
    Text1.Text = CStr(new_id("foxy", "id", "", True))
    'last_line_id = CStr(new_id("foxy", "id1", "", True))
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.Text
 
    rs.update
    
End Function

Private Sub Cmd_Click(Index As Integer)
 
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            SetForNew
            oldTxtSerial.Text = ""
        
            Me.TxtModFlg.Text = "N"
            setfoxy
            DcCostCenter.Text = ""
            lbl(27).Caption = ""
    
            Option1.value = True
            Opt(1).value = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Text2.Text = "ÌœÊÌ"
            Else
                Text2.Text = "Manual"
            End If

            Me.dcBranch.BoundText = Current_branch
            Text3.Text = ""
            Txt_DateHigri.value = ToHijriDate(Date)

        Case 1
    
            If ChekClodePeriod(DTP_Date.value) = True Then
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

            If Text2.Text <> "Manual" And Text2.Text <> "ÌœÊÌ" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰  ⁄œÌ· ﬁÌœ «·Ì «»œ«", vbCritical
                    Exit Sub
                Else
                    MsgBox "Can't Modify Auto vouchers"
                    Exit Sub
                End If

                Exit Sub
            End If

            Opt(1).value = True
            Me.TxtModFlg.Text = "E"
  
            Fg_Journal.Rows = Fg_Journal.Rows + 1
 
            CuurentLogdata

        Case 2
            If ChekClodePeriod(DTP_Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If val(TxtTotalDebit.Text) = 0 And val(TxtTotalCredit.Text) = 0 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = " There is no iAccounts in vouchers"
                Else
                    Msg = "·„ Ì „ «œŒ«· Õ”«»«  ›Ì «·ﬁÌœ"
                End If

                MsgBox Msg, vbCritical
                Exit Sub
            End If

            '  Me.DcboUsers.BoundText = user_id
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            SaveData

        Case 3
            Undo
        
        Case 4
            Frame3.Visible = True
      
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Unload Voucher_search
            Voucher_search.show

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            ShowGL_cc TxtSerial.Text, , 200

        Case 8
            If ChekClodePeriod(DTP_Date.value) = True Then
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

            Del_Trans
    End Select

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If Text2.Text <> "Manual" And Text2.Text <> "ÌœÊÌ" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰ Õ–› ﬁÌœ «·Ì «»œ«", vbCritical
        Else
            MsgBox "Can't Delete Auto vouchers"

        End If

        Exit Sub
    End If

    If TXTNoteID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·ﬁÌœ —ﬁ„ " & CHR(13)
            Msg = Msg + (Me.TxtSerial.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
   
            Msg = "Delete Voucher No. " & CHR(13)
            Msg = Msg + (Me.TxtSerial.Text) & CHR(13)
            Msg = Msg + " Confirm Delete?"
  
        End If
    
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
    
            CuurentLogdata ("D")
    
            StrSQL = "Delete  Notes  where NoteSerial =" & TxtSerial
            Cn.Execute StrSQL, , adExecuteNoRecords
     
            ' StrSQL = "Delete  Notes  where NoteID =" & Val(TxtNoteID.text)
            '  Cn.Execute StrSQL, , adExecuteNoRecords
  
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            Dim rs As New ADODB.Recordset

            'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
             "From notes where (((notes.NoteType)=200)) " & _
             "    ORDER BY NOTES.NoteID "
            StrSQL = "SELECT  Noteserial From gl_cc    where notetype<>1000  group by   Noteserial     ORDER BY  Noteserial"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           
            If rs.RecordCount < 1 Then
                clear_all Me
                '  Fg_Journal.Clear flexClearScrollable, flexClearEverything
                
                TxtModFlg_Change
               
                Fg_Journal.Clear flexClearScrollable, flexClearEverything
                Me.TxtTotalCredit.Text = 0
                Me.TxtTotalDebit.Text = 0
               
                Me.TXTResults.Text = 0

            Else

                If Not (IsNull(rs("NoteSerial").value)) Then
                    Me.Retrive rs("NoteSerial").value
                    StrOldTransID = rs("NoteSerial").value
                End If

            End If
        
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
                  
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
    
        Case "E"
    
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            '   Rs.find "id='" & Val(Me.TXTid.text) & "'", , adSearchForward, adBookmarkFirst
            '         If Rs.EOF Or Rs.BOF Then
            '            Me.TxtModFlg.text = "R"
            '            Exit Sub
            '         End If
            Retrive (val(TxtDEV_NO.Text))
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer

    With Fg_Journal

        If Not .TextMatrix(Fg_Journal.Row, .ColIndex("AccountCode")) = "" Then

            .AddItem " ", Fg_Journal.Row
        End If

    End With

End Sub

Private Sub CmdImport_Click()
On Error Resume Next
If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim BranchID As String
Dim account_serial As String
Dim des As String
Dim DebitValue As String
Dim CreditValue As String
  

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    Do Until .cells(i, 7) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
   BranchID = .cells(i, 1)
    account_serial = .cells(i, 7)
         des = .cells(i, 4)
        DebitValue = Abs(.cells(i, 5))
         CreditValue = Abs(.cells(i, 6))
         
        
 With Fg_Journal

     
  .TextMatrix(i, .ColIndex("des")) = (des)
  
   .TextMatrix(i, .ColIndex("account_serial")) = val(account_serial)
   
   Fg_Journal_AfterEdit i, .ColIndex("account_serial")
   
     .TextMatrix(i, .ColIndex("BranchId")) = val(BranchID)
          .TextMatrix(i, .ColIndex("BranchName")) = val(BranchID) 'GetBrancheName(val(BranchID))
          
  
   If val(DebitValue) > 0 Then
      .TextMatrix(i, .ColIndex("DebitValue")) = val(DebitValue)
         Fg_Journal_AfterEdit i, .ColIndex("DebitValue")

    End If
    
       If val(CreditValue) > 0 Then
     .TextMatrix(i, .ColIndex("CreditValue")) = val(CreditValue)
     Fg_Journal_AfterEdit i, .ColIndex("CreditValue")
      End If
   
        Fg_Journal.Row = i
                            Fg_Journal.Col = Fg_Journal.ColIndex("lineno")
                            Fg_Journal.ShowCell i, Fg_Journal.ColIndex("lineno")
                            
                            Fg_Journal.SetFocus


 End With
 If .cells(i, 7) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With

       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String

    If Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) <> "" Then
        sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        Cn.Execute sgl, , adExecuteNoRecords
    End If

    If Fg_Journal.Rows > 1 Then
        If Fg_Journal.Rows = 2 Then
            Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Fg_Journal.Rows > 1 Then
                If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                    Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

    With Fg_Journal
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With
            
End Sub

Private Sub CMDSelectFile_Click()
CD1.ShowOpen
txtFile.Text = CD1.filename
 End Sub

Private Sub Command1_Click()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("ked").value = Txt.Text
    rs("code").value = txtCode.Text
        
    rs.update
    '    Cn.CommitTrans
    rs.Close
End Sub

Private Sub Command2_Click()
    Unload KEDDES
    KEDDES.show
End Sub

Private Sub Command3_Click()
    Unload KEDDES
    KEDDES.show
    KEDDES.case_id = 1
    KEDDES.rowno = Fg_Journal.Row
    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)

End Sub

Private Sub Command4_Click()

    If Len(TxtDes.Text) = 0 Then Exit Sub
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[ked_desc]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("ked").value = TxtDes.Text
    rs("code").value = txtcodesub.Text
        
    rs.update
    '    Cn.CommitTrans
    rs.Close
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Dim X As Long

    If Len(Text4.Text) = 0 Then Exit Sub
    'x = get_Notes_id(Text4.text)
    X = Text4.Text

    If X <> 0 Then
        Me.Retrive3 (X)
        'Frame3.Visible = False
        ReLineGrid
        Fg_Journal.Rows = Fg_Journal.Rows + 1
        Text4.Text = ""
    End If

End Sub

Private Sub Command6_Click()
    ' .Cell(flexcpData, .Row, .ColIndex("Des")) = "Hiiiiiii"
    '                   .TextMatrix(I, .ColIndex("des")) = IIf(IsNull(Rs("Double_Entry_Vouchers_Description").value), _
                        "", Rs("Double_Entry_Vouchers_Description").value)


        Select Case lbl(10)
'07112013
  Case 170: ' ›« Ê—… „»Ì⁄« 
   
   frmsalebill.Retrive val(LblTransactionsId.Caption)
   
            Case 220: '„—œÊœ«  „»Ì⁄« 
 FrmReturnSalling.show
FrmReturnSalling.Retrive val(LblTransactionsId.Caption)

            Case 150: ' „‘ —Ì« 
 
   FrmBillBuy.Retrive val(LblTransactionsId.Caption)

            Case 230: ' „—œÊœ«  „‘ —Ì« 
        FrmReturnpurchases.show
          FrmReturnpurchases.Retrive val(LblTransactionsId.Caption)
  


'3       „’—Ê›«  Expenses
'4       „ﬁ»Ê÷«  Revenue
'5       „œ›Ê⁄«  Payments
'    14   ÕÊÌ·«  „«·ÌÂ       Financial Transfer

Case 3
FrmExpenses5.show
FrmExpenses5.Retrive val(Lblnotes_all.Caption)


Case 4 '„ﬁ»Ê÷« 
FrmCashing.show
FrmCashing.Retrive val(TXTNoteID.Text)

Case 5 '„œ›Ê⁄« 
FrmPayments.show
FrmPayments.Retrive val(TXTNoteID.Text)

Case 50
FrmPayments1.show
FrmPayments1.Retrive val(TXTNoteID.Text)

Case 14 ' ÕÊÌ·« 
FrmBoxDrawing.show
FrmBoxDrawing.Retrive val(TXTNoteID.Text)

'Case 80 ' ›« Ê—… „«·Ì…
Case 80 ' ‘—«¡ «’Ê· À =«» …
If GetFinInvoiceType(val(Lblnotes_all.Caption)) = 2 Then
        FrmExpenses4.show
        FrmExpenses4.Retrive val(Lblnotes_all.Caption)
Else
FrmExpenses3.show
FrmExpenses3.Retrive val(Lblnotes_all.Caption)

End If




Case 350  '    350 ”‰œ   ”ÊÌ…  ⁄Âœ…        Era Voucher
FrmExpenses30.show
FrmExpenses30.Retrive val(Lblnotes_all.Caption)

Case 20
FrmBankDeposite.show
     FrmBankDeposite.Retrive , val(TXTNoteID.Text)
     
Case 21
FrmBankDeposite1.show
     FrmBankDeposite1.Retrive , val(TXTNoteID.Text)
        
Case 18
 FrmReceiptPart.show
 FrmReceiptPart.Retrive , val(TXTNoteID.Text)
   ' 20   «Ìœ«⁄«  »‰ﬂÌÂ  Banks Deposite
   ' 21    Õ’Ì·«   »‰ﬂÌÂ Collection and payment of checks
    
    
'    160 ”‰œ «” ·«„  Recieve Voucher
 
'    180   ”‰œ ’—›   Issue Voucher
'    190  Õ“Ì· »÷«⁄Â »Ì‰ «·„Œ«“‰           Moning Items Between Inv

Case 160 '160 ”‰œ «” ·«„  Recieve Voucher
 FrmInpout.show
FrmInpout.Retrive val(LblTransactionsId.Caption)

Case 180 '180   ”‰œ ’—›   Issue Voucher
FrmOut.show
FrmOut.Retrive val(LblTransactionsId.Caption)
Case 190 '190  Õ“Ì· »÷«⁄Â »Ì‰ «·„Œ«“‰
FrmMoving.show
FrmMoving.Retrive val(LblTransactionsId.Caption)


        End Select

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.Text = ""

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 1
    End If

End Sub

Private Sub DTP_Date_Change()

    If Trim(TxtSerial.Text) <> "" Then
        oldTxtSerial.Text = TxtSerial.Text

    End If

    TxtSerial.Text = ""
    Txt_DateHigri.value = ToHijriDate(DTP_Date.value)

End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·ﬁÌœ  " & TxtSerial.Text & CHR(13) & "   «· «—ÌŒ  " & DTP_Date & CHR(13) & "   «·›—⁄ «·⁄«„   " & dcBranch & CHR(13) & "     „—ﬂ“ «· ﬂ·›… «·⁄«„     " & DcCostCenter & CHR(13) & "    «·„‘—Ê⁄ «·⁄«„     " & dcprojects & CHR(13) & "     «·‘—Õ «·⁄«„    " & Txt & CHR(13) & "     «·«Ã„«·Ì    " & TxtTotalDebit
                   
    '
                     
     LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr No     " & TxtSerial.Text & CHR(13) & "   Date  " & DTP_Date & CHR(13) & "   General Branch  " & dcBranch & CHR(13) & "     General Cost Center    " & DcCostCenter & CHR(13) & "    General Project     " & dcprojects & CHR(13) & "     General Des      " & Txt & CHR(13) & "     Total    " & TxtTotalDebit
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 200, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial
    Else
        AddToLogFile CInt(user_id), 200, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial
    End If
    
End Function

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)

    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "DebitValue", "CreditValue"
 
                Dim NO_OF_row As Integer
                Dim row_value As Double
                Dim cuttent_value As Double
                'remove destribution
                NO_OF_row = get_NO_OF_row(val(Text1.Text), Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")), val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))))

                If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")) = "0" Then
                    cuttent_value = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("DebitValue")))
                ElseIf Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "" And Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")) = "0" Then
                    cuttent_value = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("CreditValue")))
            
                End If

                If NO_OF_row = 0 Then

                Else
                    row_value = cuttent_value / NO_OF_row
                    sgl = "update  marakes_taklefa_temp  set value=" & row_value & "  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                    Cn.Execute sgl, , adExecuteNoRecords
                End If
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0

                If check_cost_center(Row) = False Then
                    Exit Sub
                End If
        
            Case "BranchName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("BranchId")) = StrAccountCode
                
             '   Case "BranchId"
             '   StrAccountCode = .ComboItem
             '   .TextMatrix(Row, .ColIndex("BranchName")) = StrAccountCode

 
                
                        
            Case "NEmpName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("NEmpid")) = StrAccountCode
                
                Case "Departement"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("Departementid")) = StrAccountCode
        
        

        
                   Case "FixedAsset"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("FixedAssetId")) = StrAccountCode
                
                
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If
                
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
            
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
                    Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
                    Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                If check_cost_center(Row) = False Then
                    Exit Sub
                End If

            Case "Account_Serial"
        
                .TextMatrix(Row, .ColIndex("BranchId")) = val(Me.dcBranch.BoundText)
                .TextMatrix(Row, .ColIndex("BranchName")) = Me.dcBranch.Text
          
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center_id,ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        If LastAccount(rs("Account_Code").value) = False Then
                            .TextMatrix(Row, Col) = ""
                            .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                            Exit Sub
                        End If
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    'Account_NameEng
                    End If
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    .TextMatrix(Row, .ColIndex("cost_center_id")) = IIf((rs("cost_center").value) = False, "", rs("cost_center_id").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
                .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                .TextMatrix(Row, .ColIndex("DebitValue")) = 0
        
                .TextMatrix(Row, .ColIndex("BranchId")) = val(Me.dcBranch.BoundText)
                .TextMatrix(Row, .ColIndex("BranchName")) = Me.dcBranch.Text
                sgl = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then
                    If LastAccount(StrAccountCode) = False Then
                        .TextMatrix(Row, Col) = ""
                        .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Else

                        .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                        .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    End If

                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

                Set ClsAcc = Nothing
            
                StrSQL = "SELECT  ACCOUNTS.cost_center_id,ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_code='" & StrAccountCode & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
                    .TextMatrix(Row, .ColIndex("cost_center_id")) = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
    
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ReLineGrid
 
        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·Õ”«» «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("DebitValue") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„… «·„œÌ‰…   «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("DebitValue")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change  debit value" & .Cell(flexcpTextDisplay, Row, .ColIndex("DebitValue")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        ElseIf Col = .ColIndex("CreditValue") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„… «·œ«∆‰…   «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("CreditValue")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change  Credit value" & .Cell(flexcpTextDisplay, Row, .ColIndex("CreditValue")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
 
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        ElseIf Col = .ColIndex("BranchName") Then
            LogTextA = "   ⁄œÌ· «·›—⁄  «·Ï   " & .Cell(flexcpTextDisplay, Row, .ColIndex("BranchName")) & "    ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " «·”ÿ— —ﬁ„ " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
            LogTexte = "  Change Branch Name " & .Cell(flexcpTextDisplay, Row, .ColIndex("BranchName")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName")) & " Line No " & .Cell(flexcpTextDisplay, Row, .ColIndex("LineNo"))
        
        End If

        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial)

    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
'Exit Sub
    With Fg_Journal

        If Row > .FixedRows Then
            If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
                Cancel = True
            End If
        End If

        Select Case .ColKey(Col)
Case "AccountName"
  
        
  
    
            Case "LineNo"
                .ComboList = ""
                Cancel = True
                Exit Sub

            Case "DebitValue", "CreditValue", "Account_Serial"
                .ComboList = ""

            Case "DebitValueE", "CreditValuEe", "Account_Serial"
                .ComboList = ""
            
            Case "DebitCode", "CreditCode"
                .ComboList = ""

            Case "Des"
                .ComboList = ""
                '  Cancel = True
          
            Case "Dese"
                .ComboList = ""
                '  Cancel = True
          
        End Select

    End With

End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, _
                                       ByVal Col As Long)

    With Me.Fg_Journal

        Select Case .ColKey(Col)

            Case "CC"
                ALLButton1_Click
        End Select

    End With

End Sub

Private Sub Fg_Journal_Click()
    On Error Resume Next

    'If user_id = 1 Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = CStr(user_id) Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = "" Then
    '
    If SystemOptions.usertype = UserAdminAll Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = CStr(user_id) Or Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")) = "" Then
    Else

        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Can't Edit this Record because it created by user : " & get_user_name(val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")))), vbCritical: Exit Sub
        Else
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ›Ì Â–« «·”ÿ— ·«‰Â  „ «÷«› … »Ê«”ÿ… „” Œœ„ «Œ— ÊÂÊ   : " & get_user_name(val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("userid")))), vbCritical: Exit Sub
        End If
    End If

    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) = "CC" And .TextMatrix(r, .ColIndex("AccountCode")) <> "" Then
            '        ALLButton1_Click
        End If
    
    End With

End Sub

Function check_cost_center(Row As Long) As Boolean
    check_cost_center = False

    If Auto_cost_center.value = vbChecked Then Exit Function

    'If Fg_Journal.Row = 2 Then Exit Function

    If Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("cost_center")) <> "True" Or Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("cost_center_id")) <> "" Then
        check_cost_center = True
        Exit Function

    Else

        If Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("cost_center")) = "True" And Fg_Journal.TextMatrix(Row, Fg_Journal.ColIndex("distributed")) = "" Then

            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Must select Cost Center For this Account ", vbCritical
            Else
                MsgBox " ·«»œ „‰ «œŒ«· „—ﬂ“ «· ﬂ·›…    " & " ›Ì «·”ÿ— —ﬁ„ : " & Row - 1 & " ·«‰ Â–« «·Õ”«» ·Â „—ﬂ“  ﬂ·›…  ", vbCritical
            End If

            Fg_Journal.Row = Row
            Fg_Journal.Col = 10
            Exit Function
        End If
    End If

    check_cost_center = True
End Function

Private Sub Fg_Journal_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
' SendKeys "{F4}"
End Sub

Private Sub Fg_Journal_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
'SendKeys "{BACKSPACE}"

End Sub

Private Sub Fg_Journal_DblClick()

    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" And Fg_Journal.ColKey(c) <> "Dese" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        TxtDes.Text = Fg_Journal.TextMatrix(r, Fg_Journal.ColIndex("des"))
        TxtDese.Text = Fg_Journal.TextMatrix(r, Fg_Journal.ColIndex("dese"))
        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus

        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c

        If SystemOptions.UserInterface = ArabicInterface Then
            '    TxtDes.SetFocus
        Else
            '    TxtDese.SetFocus
        End If
    
    End With

End Sub

Private Sub Fg_Journal_GotFocus()

' SendKeys "{F4}"
End Sub

Private Sub Fg_Journal_KeyDown(KeyCode As Integer, Shift As Integer)
'SendKeys "{F4}"
   '  SendKeys "{BACKSPACE}"
End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
SendKeys "{F4}"
SendKeys "{BACKSPACE}"
SendKeys CHR(KeyAscii)
End Sub

Private Sub Fg_Journal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'  SendKeys "{F4}"
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
 
        update_accounts
    End If

    If KeyCode = 46 Then
        CmdRemove_Click
    End If

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 200

    End If

    If KeyCode = vbKeyReturn Then

        With Fg_Journal

            If .Col = 7 And val(.TextMatrix(.Row, 7)) = 0 Then
                .Col = .Col + 2
            ElseIf .Col = 7 And val(.TextMatrix(.Row, 7)) <> 0 Then
                .Row = .Row + 1
                .Col = 5
           
            ElseIf .Col = 9 Then
                .Row = .Row + 1
                .Col = 5
            Else
                .Col = .Col + 1
            End If

            .ShowCell .Row, .Col + 1
            
            .SetFocus
        End With

    End If

    '.ColIndex("Account_Serial")
 
End Sub

Private Sub Fg_Journal_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    With Fg_Journal

        If Button = vbRightButton Then
            '    If .FixedRows <= .Row And .Row < .Rows - 1 Then
            '        If .TextMatrix(.Row, .ColIndex("AccountCode")) <> "" Then
            '            MDIFrmamin.MnuPopJournal_Parent.Tag = .Row
            '            MDIFrmamin.MnuPopJournal(0).Enabled = True
            '            Me.PopupMenu MDIFrmamin.MnuPopJournal_Parent
            '        Else
            '            MDIFrmamin.MnuPopJournal_Parent.Tag = .Row
            '            MDIFrmamin.MnuPopJournal(0).Enabled = False
            '            Me.PopupMenu MDIFrmamin.MnuPopJournal_Parent
            '        End If
            '    End If
        End If

    End With

End Sub

Function update_accounts()
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal
    
        If Opt(0).value = True Then
            'Tree display
            StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "

            If ChkLastAccount.value = vbChecked Then
                'StrSQL = StrSQL + " And(((ACCOUNTS.last_account) = True)) "
            End If

            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "DisName", "Account_Code")
                
        ElseIf Opt(1).value = True Then

            'Full Path Display
            If SystemOptions.UserInterface = EnglishInterface Then
                
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                End If

                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                End If
                
            Else
                
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                End If

                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                End If
                
            End If
                
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
            Debug.Print StrSQL
        ElseIf Opt(2).value = True Then 'the normal Display
            StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "

            If ChkLastAccount.value = vbChecked Then
                If SystemOptions.SysDataBaseType = AccessDataBase Then
                    StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                Else
                    StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                End If
            End If

            If OptSort(1).value = True Then
                StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
            Else
                StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
            End If

            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
        End If

        If StrComboList <> "" Then
            StrComboList = "|" & StrComboList
        End If

        .ComboList = StrComboList
   
    End With

End Function
Function IntializeGrid()
Exit Function
Dim rs As New ADODB.Recordset
 With Fg_Journal

 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                

rs.Close
 

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentName FROM         dbo.TblEmpDepartments ORDER BY DepartmentName  "
                Else
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentNamee FROM         dbo.TblEmpDepartments ORDER BY DepartmentNamee   "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList


         rs.Close

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Name from dbo.TblEmployee ORDER BY Emp_Name "
                Else
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Namee from dbo.TblEmployee ORDER BY Emp_Namee "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        rs.Close
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "  select branch_id,branch_name from TblBranchesData   "
                Else
                    StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_namee", "branch_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
 rs.Close

                If Opt(0).value = True Then
                    'Tree display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "

                    If ChkLastAccount.value = vbChecked Then
                        'StrSQL = StrSQL + " And(((ACCOUNTS.last_account) = True)) "
                    End If

                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "DisName", "Account_Code")
                
                ElseIf Opt(1).value = True Then

                    'Full Path Display
                    If SystemOptions.UserInterface = EnglishInterface Then
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                        End If
                
                    Else
                
                        '    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & _
                             "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & _
                             " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
                             "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
                             "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
                             "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code,  REPLACE(REPLACE(REPLACE(ACCOUNTS.Account_Name, CHAR(10), ''), CHAR(13), ''), CHAR(9), '')  As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                
                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                        End If
                
                    End If
 
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                    Debug.Print StrSQL
                ElseIf Opt(2).value = True Then 'the normal Display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "

                    If ChkLastAccount.value = vbChecked Then
                        If SystemOptions.SysDataBaseType = AccessDataBase Then
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                        Else
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                        End If
                    End If

                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
rs.Close
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
 

    End With

End Function
Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
'Exit Sub
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
Case "FixedAsset"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                


            Case "Departement"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentName FROM         dbo.TblEmpDepartments ORDER BY DepartmentName  "
                Else
                    StrSQL = " SELECT     TOP 100 PERCENT DeparmentID, DepartmentNamee FROM         dbo.TblEmpDepartments ORDER BY DepartmentNamee   "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList


            Case "NEmpName"


                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Name from dbo.TblEmployee ORDER BY Emp_Name "
                Else
                    StrSQL = "   SELECT     TOP 100 PERCENT Emp_ID, Emp_Namee from dbo.TblEmployee ORDER BY Emp_Namee "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "BranchName"

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "  select branch_id,branch_name from TblBranchesData   "
                Else
                    StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
                End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_name", "branch_id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "branch_namee", "branch_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "AccountName"

                If Opt(0).value = True Then
                    'Tree display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, Space(2*(Len(Account_Code)))" & "+ ACCOUNTS.Account_Name   As DisName , ACCOUNTS.Parent_Account_Code," & "ACCOUNTS.last_account, ACCOUNTS.cannot_del" & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <> 'r' "

                    If ChkLastAccount.value = vbChecked Then
                        'StrSQL = StrSQL + " And(((ACCOUNTS.last_account) = True)) "
                    End If

                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "DisName", "Account_Code")
                
                ElseIf Opt(1).value = True Then

                    'Full Path Display
                    If SystemOptions.UserInterface = EnglishInterface Then
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                        End If
                
                    Else
                
                        '    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & _
                             "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & _
                             " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
                             "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
                             "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
                             "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                
                        StrSQL = "SELECT ACCOUNTS.Account_Code,  REPLACE(REPLACE(REPLACE(ACCOUNTS.Account_Name, CHAR(10), ''), CHAR(13), ''), CHAR(9), '')  As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                
                        If ChkLastAccount.value = vbChecked Then
                            If SystemOptions.SysDataBaseType = AccessDataBase Then
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                            Else
                                StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                            End If
                        End If

                        If OptSort(1).value = True Then
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                        Else
                            StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                        End If
                
                    End If
                
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                    Debug.Print StrSQL
                ElseIf Opt(2).value = True Then 'the normal Display
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del " & "From ACCOUNTS Where  ACCOUNTS.Account_Code <>'r' "

                    If ChkLastAccount.value = vbChecked Then
                        If SystemOptions.SysDataBaseType = AccessDataBase Then
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                        Else
                            StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                        End If
                    End If

                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If

                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
        End Select

    End With
 'SendKeys "{F4}"
End Sub

Private Sub Form_Activate()
    'Application_Mode Me.TxtModFlg.text
End Sub

Private Sub Form_Load()
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If
    
    ScreenNameArabic = "  ﬁÌœ «·ÌÊ„Ì…"
    ScreenNameEnglish = "GL Entry"
    
    Dim StrSQL  As String
    Dim GrdBck As New ClsBackGroundPic

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL
    'StrSQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
    StrSQL = "  select id,Project_name from projects " '

    fill_combo Me.dcprojects, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select branch_id,branch_name from TblBranchesData   "
    Else
        StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
    End If

    fill_combo dcBranch, StrSQL

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(8).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Me.TxtModFlg.Text = "R"
    SetDtpickerDate Me.DTP_Date
    Me.TabMain.CurrTab = 0

    ' adjust the grid
    With Fg_Journal

        If SystemOptions.usertype = UserAdminAll Then
            .ColHidden(.ColIndex("BranchName")) = False
        Else
            .ColHidden(.ColIndex("BranchName")) = True
         
        End If

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(.ColIndex("LineNo")) = True
        .Cell(flexcpText, 0, .ColIndex("LineNo"), 1, .ColIndex("LineNo")) = "—ﬁ„ «·”ÿ—"

        .MergeCol(.ColIndex("DebitValue")) = True
        .MergeCol(.ColIndex("CreditValue")) = True
        .MergeCol(.ColIndex("Account_Serial")) = True
        .MergeCol(.ColIndex("AccountName")) = True
    
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
        .ColWidth(.ColIndex("Account_Serial")) = 1500

        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "«”„ «·Õ”«»"
        .ColWidth(.ColIndex("AccountName")) = 4500
    
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·„Õ·Ì… "
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter

        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "„œÌ‰"
        .ColWidth(.ColIndex("DebitValue")) = 1590
        .ColFormat(.ColIndex("DebitValue")) = "#,###.00" ' SystemOptions.SysDefCurrencyForamt
     
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "œ«∆‰"
        .ColWidth(.ColIndex("CreditValue")) = 1590
        .ColFormat(.ColIndex("CreditValue")) = "#,###.00"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = " ﬁÌ„… «·ﬁÌœ »«·⁄„·… «·«Ã‰»Ì… "
    
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter
        
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "œ«∆‰"
        .ColFormat(.ColIndex("DebitValueE")) = "#,###.00"
        .ColFormat(.ColIndex("CreditValueE")) = "#,###.00"

        '.MergeCol(.ColIndex("Des")) = True
        '.Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "«·‘—Õ"
        '.ColWidth(.ColIndex("Des")) = 2200
        Set .WallPaper = GrdBck.Picture

        ' .Cols = .Cols + 1
        ' .ColKey(.Cols - 1) = "Remarks"
        .ColComboList(.ColIndex("CC")) = "..."
    
    End With

    'If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    ChangeLang
    'End If
    'Me.Img(0).Picture = MDIFrmamin.ImgLstMenuIcons.ListImages("Unlock").Picture
    'Img(0).Visible = True
    'Me.Img(1).Picture = MDIFrmamin.ImgLstMenuIcons.ListImages("Lock").Picture
    'Img(1).Visible = False
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DcboUsers
    AddTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    StrSQL = "SELECT  Noteserial  From gl_cc    where notetype <>1000  group by    Noteserial     ORDER BY  Noteserial"

    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = "SELECT  Noteserial  From gl_cc    where branch_no=" & branch_id & " and  notetype <>1000  group by    Noteserial     ORDER BY  Noteserial"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    'IntializeGrid
  '  Resize_Form Me, TransactionSize
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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
      
                Cmd_Click (2)
      
                ' SaveData
            Case vbNo

                If Me.TxtModFlg.Text = "N" Then
                    sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
                    Cn.Execute sgl, , adExecuteNoRecords
                End If
      
            Case vbCancel
                Cancel = True
        End Select
      
    End If

    Exit Sub
ErrTrap:

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    '    Select Case QueryCloseMsg(Me.TxtModFlg.text, Me.Caption)
    '        Case vbYes
    '            Cancel = True
    '            Do_Action Do_save
    '        Case vbNo
    '            Cancel = False
    '            Application_Mode "R"
    '        Case vbCancel
    '            Cancel = True
    '    End Select
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Dcombos = Nothing
    Set DCboSearch = Nothing
    Set TTP = Nothing
   rs.Close
    Set rs = Nothing
End Sub


 

Private Sub Label10_Click()
    PicDes.Visible = False
End Sub

Private Sub Opt_Click(Index As Integer)

    Select Case Index

        Case 0
            ChkLastAccount.Enabled = False

        Case 1
            ChkLastAccount.Enabled = True

        Case 2
            ChkLastAccount.Enabled = True
    End Select

End Sub

Private Function LastAccount(StrAccountCode As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    If StrAccountCode = "" Then
        LastAccount = False
        Exit Function
    End If

    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account,ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Code='" & StrAccountCode & "'"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs("last_account").value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«·Õ”«» " & rs("Account_Name").value & CHR(13)
            Msg = Msg & "Õ”«» €Ì— ‰Â«∆Ï Ê·«Ì„ﬂ‰ ﬂ «»… ﬁÌœ ⁄·ÌÂ " & CHR(13)
            Msg = Msg & "»—Ã«¡  ÕœÌœ √Ï Õ”«» ›—⁄Ï  Õ  Â–« «·Õ”«»" & CHR(13)
            Msg = Msg & "√Ê ﬁ„ » ⁄—Ì› Õ”«»«  ›—⁄Ì… ÃœÌœ  Õ  Â–« «·Õ”«»"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            Msg = "The " & IIf(IsNull(rs("Account_NameEng").value), rs("Account_Name").value, rs("Account_NameEng").value) & " Account " & CHR(13)
            Msg = Msg & "is not a last account..!" & CHR(13)
            Msg = Msg & "and it is not accepted."
            MsgBox Msg, vbExclamation, App.title
        End If

        LastAccount = False
    Else
        LastAccount = True
    End If

Exit_Function:
    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    LastAccount = False
    Resume Exit_Function
End Function

Private Sub SetForNew()
    Me.Txt.Text = ""
    Check1.value = Unchecked
    Check2.value = Unchecked
    Check3.value = Unchecked
    Check4.value = Unchecked
    Check5.value = Unchecked
    txt_salary.Text = ""
    Me.TXTNoteID.Text = ""
    Me.TxtDEVID.Text = ""
    Me.DTP_Date.value = Date
    Me.TxtSerial.Text = ""
    Me.TxtValue.Text = ""

    Me.ChkPost.value = vbUnchecked

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.ChkPost.Caption = "€Ì— „—Õ·"
    Else
        Me.ChkPost.Caption = "Not Poasted"
    End If

    Me.ChkPost.ForeColor = vbBlack
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    
    Me.TxtTotalCredit.Text = 0
    Me.TxtTotalDebit.Text = 0
    Me.TXTResults.Text = 0
    Me.DcboUsers.BoundText = user_id
    txtFile.Text = ""
End Sub

Public Property Let Cmd_New(ByVal vNewValue As Boolean)
    m_Cmd_New = vNewValue
End Property

Public Property Get Cmd_Undo() As Boolean
    'Dim Msg As String
    'Dim BolTemp  As Boolean
    'Cmd_Undo = m_Cmd_Undo
    'On Error GoTo ErrTrap
    'Select Case TxtModFlg.text
    '    Case "N"
    '        If QueryUndoMsg(Me.TxtModFlg.text, Me.Caption) = vbYes Then
    '            BolTemp = Cmd_New
    '        Else
    '            Cmd_Undo = False
    '            Exit Property
    '        End If
    '    Case "E"
    '        If QueryUndoMsg(Me.TxtModFlg.text, Me.Caption) = vbYes Then
    '           Me.Retrive Me.TxtNoteID
    '            Cmd_Undo = True
    '        Else
    '            Cmd_Undo = False
    '            Exit Property
    '        End If
    'End Select
    'Cmd_Undo = True
    'Exit Property
    'ErrTrap:
End Property

Public Property Let Cmd_Undo(ByVal vNewValue As Boolean)
    m_Cmd_Undo = vNewValue
End Property

Private Sub PicDes_Resize()

    With PicDes
        '  LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        '  TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub Txt_DateHigri_LostFocus()
    DTP_Date.value = ToGregorianDate(Txt_DateHigri.value)
    'DTP_Date_Change
End Sub

Private Sub TxtDes_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
    'TxtDes.RightToLeft = True
    TxtDes.Alignment = 1

End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        '    PutData
        '    CboDes.CloseUp
    End If

End Sub

Private Sub TxtDes_LostFocus()
    'PicHeight = PicDes.Height
    'PicWidth = PicDes.Width
    'CboDes.CloseUp
    'CboDes.Visible = False
End Sub

Private Sub TxtDesE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtModFlg_Change()

    Select Case TxtModFlg.Text

        Case "N"
            Me.EleHeader.Enabled = True
            Me.Fg_Journal.Editable = flexEDKbdMouse
            EleOpt.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = True
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
            CmdRemove.Enabled = True
            Fg_Journal.Enabled = True
            ALLButton1.Enabled = True
            Auto_cost_center.value = vbUnchecked

        Case "E"
            Me.EleHeader.Enabled = True
            Me.Fg_Journal.Editable = flexEDKbdMouse
            EleOpt.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = True
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
            CmdRemove.Enabled = True
            Fg_Journal.Enabled = True

            If Auto_cost_center.value = vbUnchecked Then
                ALLButton1.Enabled = True
            Else
    
                ALLButton1.Enabled = False
            End If
  
        Case "R"
            Fg_Journal.Editable = flexEDNone
         '   Me.EleHeader.Enabled = False
            '   Me.Fg_Journal.Editable = flexEDNone
            EleOpt.Enabled = False
            CboDes.CloseUp
            CboDes.Visible = False
        
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(4).Enabled = False
            Cmd(5).Enabled = True
            Cmd(7).Enabled = True
            CmdRemove.Enabled = False
            '   Fg_Journal.Enabled = False
            ' ALLButton1.Enabled = False
    End Select

End Sub

Public Function ReLineGridP()
    ReLineGrid
End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
            
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter

                If .TextMatrix(i, .ColIndex("LineNo1")) = "" Then
                    ' setfoxy_Line
                    .TextMatrix(i, .ColIndex("LineNo1")) = setfoxy_Line  'last_line_id
        
                End If
            
            End If
 
        Next i

    End With

    line_no1 = IntCounter
    Coloring
End Sub

Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 21) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 21) = vbWhite
            End If

        Next i

    End With

    line_no1 = IntCounter

End Sub

Public Property Get Cmd_Search() As Boolean
    Cmd_Search = m_Cmd_Search
    Frm_SandSearch.show vbModal
    Cmd_Search = True
End Property

Public Property Let Cmd_Search(ByVal vNewValue As Boolean)
    m_Cmd_Search = vNewValue
End Property

Public Sub Retrive(LngNoteID As String)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    'StrSQL = "SELECT NOTES.project_id, NOTES.project_depit_or_credit,  NOTES.foxy_no,NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & _
     "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & _
     "NOTES.Remark,NOTES.general_cost_center, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,DOUBLE_ENTREY_VOUCHERS.USERID," & _
     "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & _
     "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,DOUBLE_ENTREY_VOUCHERS.Valuee,DOUBLE_ENTREY_VOUCHERS.currency,DOUBLE_ENTREY_VOUCHERS.rate," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & _
     ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & _
     " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
     " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & _
     "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    'StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    'StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    StrSQL = "select * from gl_cc_new where Noteserial='" & LngNoteID & "'"
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    ' rs.find "Noteserial=" & LngNoteID & "'", , adSearchForward, 1
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    If rs("DAWRY").value = 0 Then
        Check4.value = vbUnchecked
        LblDawry.Caption = ""
    Else
        Check4.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            LblDawry.Caption = "ﬁÌœ œÊ—Ì"
        Else
            LblDawry.Caption = "Repeated JLE"
        End If
    End If
  
    If rs("KALEB").value = 0 Then
        Check3.value = vbUnchecked
        LblKaleb.Caption = ""
    Else
        Check3.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            LblKaleb.Caption = "ﬁ«·»"
        Else
            LblKaleb.Caption = "Template"
        End If
    End If
  
    If rs("auto_des").value = 0 Or IsNull(rs("auto_des").value) Then
        Me.Auto_cost_center.value = vbUnchecked
        ALLButton1.Enabled = True
    Else
        Auto_cost_center.value = vbChecked
        ALLButton1.Enabled = False
    End If
  
    ' Check3.value = RsNetes("KALEB").value
    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
 
    Me.txt_salary.Text = IIf(IsNull(rs("salary").value), 0, rs("salary").value)
 
    Me.TXTNoteID.Text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.dcprojects.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    If Not IsNull(rs("project_depit_or_credit").value) Then
        If rs("project_depit_or_credit").value = 0 Then
            Option1.value = True
        ElseIf rs("project_depit_or_credit").value = 1 Then
            Option2.value = True
        End If
    End If

    Dim NotesTypeNameE As String

    If SystemOptions.UserInterface = ArabicInterface Then
        Text3.Text = get_note_type_name(rs("Notetype").value)
    Else
        Text3.Text = get_note_type_name(rs("Notetype").value, NotesTypeNameE)
        Text3.Text = NotesTypeNameE
    End If

    Me.TxtDEVID.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)
    Me.TxtDEV_NO.Text = ""
    Me.TxtValue.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    Me.TxtDEV_NO.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)

    Me.DTP_Date.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(DTP_Date.value), rs("NoteDateH").value)

    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), Date, rs("NoteSerial").value)
    Me.oldTxtSerial.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value), rs("OldNoteSerial1").value)
    Me.TxtManualNO.Text = IIf(IsNull(rs("ManualNO").value), "", rs("ManualNO").value)
 
    If rs("Notetype").value = 200 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Text2.Text = "ÌœÊÌ"
        Else
            Text2.Text = "Manual"
        End If

        lbl(27).Caption = showLabel(TxtSerial, oldTxtSerial)
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Text2.Text = "«·Ì"
        Else
            Text2.Text = "Auto"
        End If

        lbl(27).Caption = ""
    End If
lbl(10).Caption = IIf(IsNull(rs("Notetype").value), "", rs("Notetype").value)
LblTransactionsId.Caption = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
Lblnotes_all.Caption = IIf(IsNull(rs("notes_all").value), 0, rs("notes_all").value)
    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    Me.DcboUsers.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    Me.Txte.Text = IIf(IsNull(rs("RemarkE").value), "", rs("RemarkE").value)

    If Not (IsNull(rs("NoteType").value)) Then
        If rs("NoteType").value = "2" Then
            'Me.OptType(0).Value = True
        ElseIf rs("NoteType").value = 1 Then
            'Me.OptType(1).Value = True
        End If
    End If

    If rs("NotePosted").value = True Then
        ChkPost.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "„—Õ·"
            lblPost.Caption = "„—Õ·"
        Else
            ChkPost.Caption = "Posted"
            lblPost.Caption = "Posted"
        End If

        ChkPost.ForeColor = vbRed
    Else
        ChkPost.value = vbUnchecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "€Ì— „—Õ·"
        Else
            ChkPost.Caption = "Not Posted"
        End If

        ChkPost.ForeColor = vbBlack
    End If

    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            
            End If
            
            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Nameeng").value), "", rs("Account_Nameeng").value)
                 
            Else
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            End If
            
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If

            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
             
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
            
            
            
            .TextMatrix(i, .ColIndex("Departementid")) = IIf(IsNull(rs("Departementid").value), "", rs("Departementid").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                 
            Else
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
            End If
            
            
                   .TextMatrix(i, .ColIndex("FixedAssetId")) = IIf(IsNull(rs("FixedAssetId").value), "", rs("FixedAssetId").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("FixedAsset")) = IIf(IsNull(rs("fixedassetname").value), "", rs("fixedassetname").value)
                 
            Else
                .TextMatrix(i, .ColIndex("FixedAsset")) = IIf(IsNull(rs("fixedassetnamee").value), "", rs("fixedassetnamee").value)
            End If
            
            
                     .TextMatrix(i, .ColIndex("NEmpid")) = IIf(IsNull(rs("NEmpid").value), "", rs("NEmpid").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(rs("NEmpName").value), "", rs("NEmpName").value)
                 
            Else
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(rs("NEmpNamee").value), "", rs("NEmpNamee").value)
            End If
            
            
            
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
    
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
    
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        '    Me.TxtTotalCredit.text =Round(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
    
        '    Me.TxtTotalDebit.text =Round(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
    
    End With

    Coloring
End Sub

Public Sub Retrive3(LngNoteID As String)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

 
    StrSQL = "select * from gl_cc_new where Noteserial='" & LngNoteID & "'"
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    ' rs.find "Noteserial=" & LngNoteID & "'", , adSearchForward, 1
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    If rs("DAWRY").value = 0 Then
        Check4.value = vbUnchecked
        LblDawry.Caption = ""
    Else
        Check4.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            LblDawry.Caption = "ﬁÌœ œÊ—Ì"
        Else
            LblDawry.Caption = "Repeated JLE"
        End If
    End If
  
    If rs("KALEB").value = 0 Then
        Check3.value = vbUnchecked
        LblKaleb.Caption = ""
    Else
        Check3.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            LblKaleb.Caption = "ﬁ«·»"
        Else
            LblKaleb.Caption = "Template"
        End If
    End If
  
    If rs("auto_des").value = 0 Or IsNull(rs("auto_des").value) Then
        Me.Auto_cost_center.value = vbUnchecked
        ALLButton1.Enabled = True
    Else
        Auto_cost_center.value = vbChecked
        ALLButton1.Enabled = False
    End If
  
    ' Check3.value = RsNetes("KALEB").value
    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
 
    Me.txt_salary.Text = IIf(IsNull(rs("salary").value), 0, rs("salary").value)
 
'    Me.TxtNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.dcprojects.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    If Not IsNull(rs("project_depit_or_credit").value) Then
        If rs("project_depit_or_credit").value = 0 Then
            Option1.value = True
        ElseIf rs("project_depit_or_credit").value = 1 Then
            Option2.value = True
        End If
    End If

    Dim NotesTypeNameE As String

    If SystemOptions.UserInterface = ArabicInterface Then
        Text3.Text = get_note_type_name(rs("Notetype").value)
    Else
        Text3.Text = get_note_type_name(rs("Notetype").value, NotesTypeNameE)
        Text3.Text = NotesTypeNameE
    End If

    Me.TxtDEVID.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)
    Me.TxtDEV_NO.Text = ""
    Me.TxtValue.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    Me.TxtDEV_NO.Text = IIf(IsNull(rs("Double_Entry_Vouchers_ID").value), "", rs("Double_Entry_Vouchers_ID").value)

    Me.DTP_Date.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(DTP_Date.value), rs("NoteDateH").value)

'    Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), Date, rs("NoteSerial").value)
'    Me.oldTxtSerial.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value), rs("OldNoteSerial1").value)
'    Me.txtManualNo.text = IIf(IsNull(rs("ManualNO").value), "", rs("ManualNO").value)
 
    If rs("Notetype").value = 200 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Text2.Text = "ÌœÊÌ"
        Else
            Text2.Text = "Manual"
        End If

        lbl(27).Caption = showLabel(TxtSerial, oldTxtSerial)
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Text2.Text = "«·Ì"
        Else
            Text2.Text = "Auto"
        End If

        lbl(27).Caption = ""
    End If
lbl(10).Caption = IIf(IsNull(rs("Notetype").value), "", rs("Notetype").value)
LblTransactionsId.Caption = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
Lblnotes_all.Caption = IIf(IsNull(rs("notes_all").value), 0, rs("notes_all").value)
    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    Me.DcboUsers.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    Me.Txte.Text = IIf(IsNull(rs("RemarkE").value), "", rs("RemarkE").value)

    If Not (IsNull(rs("NoteType").value)) Then
        If rs("NoteType").value = "2" Then
            'Me.OptType(0).Value = True
        ElseIf rs("NoteType").value = 1 Then
            'Me.OptType(1).Value = True
        End If
    End If

    If rs("NotePosted").value = True Then
        ChkPost.value = vbChecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "„—Õ·"
            lblPost.Caption = "„—Õ·"
        Else
            ChkPost.Caption = "Posted"
            lblPost.Caption = "Posted"
        End If

        ChkPost.ForeColor = vbRed
    Else
        ChkPost.value = vbUnchecked

        If SystemOptions.UserInterface = ArabicInterface Then
            ChkPost.Caption = "€Ì— „—Õ·"
        Else
            ChkPost.Caption = "Not Posted"
        End If

        ChkPost.ForeColor = vbBlack
    End If

    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            
            End If
            
            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Nameeng").value), "", rs("Account_Nameeng").value)
                 
            Else
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            End If
            
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If

            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", Round(rs("Value").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", Round(rs("Valuee").value, SystemOptions.SysDefCurrencyForamt))
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
             
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
            
            
            
            .TextMatrix(i, .ColIndex("Departementid")) = IIf(IsNull(rs("Departementid").value), "", rs("Departementid").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                 
            Else
                .TextMatrix(i, .ColIndex("Departement")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
            End If
            
            
                   .TextMatrix(i, .ColIndex("FixedAssetId")) = IIf(IsNull(rs("FixedAssetId").value), "", rs("FixedAssetId").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("FixedAsset")) = IIf(IsNull(rs("fixedassetname").value), "", rs("fixedassetname").value)
                 
            Else
                .TextMatrix(i, .ColIndex("FixedAsset")) = IIf(IsNull(rs("fixedassetnamee").value), "", rs("fixedassetnamee").value)
            End If
            
            
                     .TextMatrix(i, .ColIndex("NEmpid")) = IIf(IsNull(rs("NEmpid").value), "", rs("NEmpid").value)
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(rs("NEmpName").value), "", rs("NEmpName").value)
                 
            Else
                .TextMatrix(i, .ColIndex("NEmpName")) = IIf(IsNull(rs("NEmpNamee").value), "", rs("NEmpNamee").value)
            End If
            
            
            
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
    
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
    
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        '    Me.TxtTotalCredit.text =Round(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
    
        '    Me.TxtTotalDebit.text =Round(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
    
    End With

    Coloring
End Sub

Public Sub Retrive2(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    'StrSQL = "SELECT  NOTES.foxy_no,NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & _
     "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & _
     "NOTES.Remark, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,DOUBLE_ENTREY_VOUCHERS.USERID," & _
     "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & _
     "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,DOUBLE_ENTREY_VOUCHERS.Valuee,DOUBLE_ENTREY_VOUCHERS.currency,DOUBLE_ENTREY_VOUCHERS.rate," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & _
     ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & _
     " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
     " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & _
     "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    'StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    'StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"
    StrSQL = "select * from gl_cc_new where NoteID='" & LngNoteID & "'"
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    'If Rs("DAWRY").value = 0 Then
    'Check4.value = vbUnchecked
    'Else
    ' Check4.value = vbChecked
    'End If
  
    '  If Rs("KALEB").value = 0 Then
    'Check3.value = vbUnchecked
    'Else
    ' Check3.value = vbChecked
    'End If
  
    ' Check3.value = RsNetes("KALEB").value
    
    'Me.TxtNoteID.text = IIf(IsNull(Rs("NoteID").value), "", Rs("NoteID").value)
    'Me.Text1.text = IIf(IsNull(Rs("foxy_no").value), "", Rs("foxy_no").value)

    'If Rs("Notetype").value = 200 Then
    'Text2.text = "Manual"

    'Else
    'Text2.text = "Auto"

    'End If

    'Text3.text = get_note_type_name(Rs("Notetype").value)

    'Me.TxtDEVID.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)
    'Me.TxtDEV_NO.text = ""
    'Me.TxtValue.text = IIf(IsNull(Rs("Note_Value").value), "", Rs("Note_Value").value)
    'Me.TxtDEV_NO.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)

    'Me.DTP_Date.value = IIf(IsNull(Rs("NoteDate").value), Date, Rs("NoteDate").value)
    'Me.TxtSerial.text = IIf(IsNull(Rs("NoteSerial").value), Date, Rs("NoteSerial").value)

    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    'Me.DcboUsers.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
    'Me.Txt.text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
    'If Not (IsNull(Rs("NoteType").value)) Then
    '    If Rs("NoteType").value = "2" Then
    '        'Me.OptType(0).Value = True
    '    ElseIf Rs("NoteType").value = 1 Then
    '        'Me.OptType(1).Value = True
    '    End If
    'End If
    'If Rs("NotePosted").value = True Then
    '    ChkPost.value = vbChecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "„—Õ·"
    '    Else
    '        ChkPost.Caption = "Posted"
    '    End If
    '    ChkPost.ForeColor = vbRed
    'Else
    '    ChkPost.value = vbUnchecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "€Ì— „—Õ·"
    '    Else
    '        ChkPost.Caption = "Not Posted"
    '    End If
    '    ChkPost.ForeColor = vbBlack
    'End If
    Dim last_row As Integer
    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = 3
        last_row = .Rows
        .Rows = .Rows + rs.RecordCount - 1

        For i = last_row - 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            
            End If

            .TextMatrix(i, .ColIndex("LineNo")) = i ' IIf(IsNull(Rs("DEV_ID_Line_No").value), "", Rs("DEV_ID_Line_No").value)
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Nameeng").value), "", rs("Account_Nameeng").value)
                 
            Else
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            End If
            
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If

            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("DebitValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = "0"
            
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
            
                .TextMatrix(i, .ColIndex("CreditValuee")) = IIf(IsNull(rs("Valuee").value), "", rs("Valuee").value)
                .TextMatrix(i, .ColIndex("DebitValuee")) = "0"
                
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("userid")) = IIf(IsNull(rs("userid").value), "", rs("userid").value)
            
            .TextMatrix(i, .ColIndex("currenct_code")) = IIf(IsNull(rs("currency").value), "", rs("currency").value)
            
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(rs("rate").value), "", rs("rate").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
                       
    End With

End Sub

Public Sub retrive1(LngNoteID As Long)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i  As Integer

    If LngNoteID = 0 Then
        Exit Sub
    End If

    'StrSQL = "SELECT  NOTES.KALEB, NOTES.DAWRY, NOTES.NoteID,  NOTES.NoteType," & _
     "NOTES.NoteDate, NOTES.Note_Value,NOTES.NoteHijriDate," & _
     "NOTES.Remark, NOTES.NotePosted,NOTES.UserID,NoteSerial ," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID," & _
     "DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, DOUBLE_ENTREY_VOUCHERS.Account_Code," & _
     "DOUBLE_ENTREY_VOUCHERS.Value, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit," & _
     "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,ACCOUNTS.Account_Name  " & _
     ",ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial " & _
     " FROM ACCOUNTS INNER JOIN (NOTES INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
     " ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON " & _
     "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code "

    'StrSQL = StrSQL + " Where NOTES.NoteID=" & LngNoteID & ""
    'StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    StrSQL = "select * from gl_cc_new where NoteID='" & LngNoteID & "'"
    StrSQL = StrSQL + "Order By (DEV_ID_Line_No)"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Exit Sub
    End If

    ' If Rs("DAWRY").value = 0 Then
    ' ' Check3.value = vbUnchecked
    '' Else
    ' Check3.value = vbChecked
    'End If
  
    '    If Rs("KALEB").value = 0 Then
    '  Check4.value = vbUnchecked
    '  Else
    '   Check4.value = vbChecked
    '  End If
    '
    ' Check3.value = RsNetes("KALEB").value
    
    'Me.TxtNoteID.text = IIf(IsNull(Rs("NoteID").value), "", Rs("NoteID").value)

    'Me.TxtDEVID.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)
    'Me.TxtDEV_NO.text = ""
    'Me.TxtValue.text = IIf(IsNull(Rs("Note_Value").value), "", Rs("Note_Value").value)
    'Me.TxtDEV_NO.text = IIf(IsNull(Rs("Double_Entry_Vouchers_ID").value), "", Rs("Double_Entry_Vouchers_ID").value)

    'Me.DTP_Date.value = IIf(IsNull(Rs("NoteDate").value), Date, Rs("NoteDate").value)
    'Me.TxtSerial.text = IIf(IsNull(Rs("NoteSerial").value), Date, Rs("NoteSerial").value)

    'Me.DtHijriTrans.value = IIf(IsNull(Rs("NoteHijriDate").value), "", Rs("NoteHijriDate").value)
    'Me.DcboUsers.BoundText = IIf(IsNull(Rs("UserID").value), "", Rs("UserID").value)
    'Me.Txt.text = IIf(IsNull(Rs("Remark").value), "", Rs("Remark").value)
    'If Not (IsNull(Rs("NoteType").value)) Then
    '    If Rs("NoteType").value = "2" Then
    '        'Me.OptType(0).Value = True
    '    ElseIf Rs("NoteType").value = 1 Then
    '        'Me.OptType(1).Value = True
    '    End If
    'End If
    'If Rs("NotePosted").value = True Then
    '    ChkPost.value = vbChecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "„—Õ·"
    '    Else
    '        ChkPost.Caption = "Posted"
    '    End If
    '    ChkPost.ForeColor = vbRed
    'Else
    '    ChkPost.value = vbUnchecked
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        ChkPost.Caption = "€Ì— „—Õ·"
    '    Else
    '        ChkPost.Caption = "Not Posted"
    '    End If
    '    ChkPost.ForeColor = vbBlack
    'End If

    rs.MoveFirst

    With Me.Fg_Journal
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("branch_id").value), "", rs("branch_id").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            
            End If

            .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(rs("DEV_ID_Line_No").value), "", rs("DEV_ID_Line_No").value)
            
            .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(rs("DEV_ID_Line_No1").value), "", rs("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
            .Cell(flexcpData, i, .ColIndex("Des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
            
            .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("Double_Entry_Vouchers_Description").value), "", rs("Double_Entry_Vouchers_Description").value)
             
            .TextMatrix(i, .ColIndex("dese")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)

            If Trim(.Cell(flexcpData, i, .ColIndex("Des"))) <> "" Then
                .Cell(flexcpPicture, i, .ColIndex("Des")) = ImgNote.Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("Des")) = flexAlignLeftCenter
            Else
                .Cell(flexcpPicture, i, .ColIndex("Des")) = Empty
            End If
        
            If rs("Credit_Or_Debit").value = 0 Then
                .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                .TextMatrix(i, .ColIndex("CreditValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignRightCenter
            Else
                .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                .TextMatrix(i, .ColIndex("DebitValue")) = "0"
                .Cell(flexcpAlignment, i, .ColIndex("AccountName"), i, .ColIndex("AccountName")) = flexAlignLeftCenter
            End If

            .TextMatrix(i, .ColIndex("USERID")) = IIf(IsNull(rs("USERID").value), "", rs("USERID").value)
            
            rs.MoveNext
        Next i

        Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))

        Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
    
        Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)
        Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
    
    End With

End Sub

Public Property Get Cmd_Edit() As Boolean
    Dim Msg As String
    Cmd_Edit = m_Cmd_Edit

    If Trim(Me.TXTNoteID.Text) = "" Then
        'Msg = "·«ÌÊÃœ ”Ã· Õ«÷— ·· ⁄œÌ·"
        GetMsgs 72, vbExclamation
        Cmd_Edit = False
        Exit Property
    ElseIf Me.ChkPost.value = vbChecked Then
        'Msg = "Â–« «·”‰œ „—Õ· ...!!" & Chr(13)
        'Msg = Msg & "Ê·« Ì„ﬂ‰  ⁄œÌ· «·ﬁÌœ"
        GetMsgs 73, vbExclamation
        Cmd_Edit = False
        Exit Property
    Else
        Me.DcboUsers.BoundText = user_id 'LngUserID
        Cmd_Edit = True
        Exit Property
    End If

End Property

Public Property Let Cmd_Edit(ByVal vNewValue As Boolean)
    m_Cmd_Edit = vNewValue
End Property

Public Property Get Cmd_Delete() As Boolean
    Dim StrSQL  As String
    Dim Msg As String
    Dim BolTemp As Boolean
    Dim TransBegine As Boolean
    Dim rs As New ADODB.Recordset
    Dim IntRes As Integer
    On Error GoTo ErrTrap
    Cmd_Delete = m_Cmd_Delete

    If Me.TXTNoteID.Text = "" Then
        Cmd_Delete = True
        Exit Property
    End If

    If Me.ChkPost.value = vbChecked Then
        'Msg = "Â–« «·”‰œ „—Õ· ...!!" & Chr(13)
        'Msg = Msg & "Ê·« Ì„ﬂ‰ Õ–› «·ﬁÌœ...!!"
        GetMsgs 74, vbExclamation
        Cmd_Delete = True
        Exit Property
    End If

    StrSQL = "Delete * From Notes Where Notes.Note_ID='" & Trim(Me.TXTNoteID.Text) & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ê› Ì „ Õ–› Â–« «·”‰œ —ﬁ„ " & Trim(Me.TxtSerial.Text) & CHR(13)
        Msg = Msg & "›Â· √‰  „ √ﬂœ „‰ «·√” „—«— ...!!"
        IntRes = MsgBox(Msg, vbQuestion + vbOKCancel + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
    Else
        Msg = "This voucher " & Trim(Me.TxtSerial.Text) & CHR(13)
        Msg = Msg & "will be deleted " & CHR(13)
        Msg = Msg & "are you sure to continue ..?"
        IntRes = MsgBox(Msg, vbQuestion + vbOKCancel, App.title)
    End If

    If IntRes = vbOK Then
        Cn.BeginTrans
        TransBegine = True
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.CommitTrans
        TransBegine = False
    
        'Msg = " „ Õ–› «·”Ã·."
        GetMsgs 75, vbInformation
    End If

    Cmd_Delete = True
    Exit Property
ErrTrap:

    If TransBegine = True Then
        Cn.RollbackTrans
    End If

    'Msg = "ÕœÀ Œÿ√ √À‰«¡ Õ–› «·”Ã·"
    GetMsgs 76, vbExclamation
    Cmd_Delete = True
End Property

Public Property Let Cmd_Delete(ByVal vNewValue As Boolean)
    m_Cmd_Delete = vNewValue
End Property

Private Sub PutData()
    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)

    With Fg_Journal

        If Len(TxtDes.Text) > 0 And Len(TxtDese.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = TxtDes.Text
        
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("dese")) = TxtDese.Text
        ElseIf Len(TxtDes.Text) > 0 And Len(TxtDese.Text) = 0 Then
    
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("dese")) = ""
        ElseIf Len(TxtDes.Text) = 0 And Len(TxtDese.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
            .TextMatrix(.Row, .ColIndex("des")) = ""
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = TxtDes.Text
            .TextMatrix(.Row, .ColIndex("dese")) = TxtDese.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Dese")) = flexAlignLeftCenter
        ElseIf Len(TxtDes.Text) = 0 And Len(TxtDese.Text) = 0 Then
            .TextMatrix(.Row, .ColIndex("des")) = ""
            .TextMatrix(.Row, .ColIndex("dese")) = ""
    
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
            .Cell(flexcpData, .Row, .ColIndex("Dese")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Dese")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        
        End If

    End With

End Sub

Public Property Get Cmd_Print() As Boolean

    If Me.TXTNoteID.Text = "" Then
        GetMsgs 140, vbExclamation
        Cmd_Print = False
    Else
        Cmd_Print = FireReport(PrinterTarget)
    End If

End Property

Public Property Let Cmd_Print(ByVal vNewValue As Boolean)
    m_Cmd_Print = vNewValue
End Property

Private Function FireReport(m_Destination As PrintTarget) As Boolean
    'Dim RsData As New ADODB.Recordset
    'Dim Rs As New ADODB.Recordset
    'Dim xApp As New CRAXDRT.Application
    'Dim xReport As CRAXDRT.Report
    'Dim Msg As String
    'Dim StrSQL As String
    'Dim StrPrinterName As String
    'Dim XPrinter As Object
    'Dim Frm As FrmPrint
    'Dim I As Integer
    'Dim StrFileName As String
    'On Error GoTo FireReportErrTrap
    'If Me.TxtNoteID.text = "" Then
    '    FireReport = False
    '    Exit Function
    'End If
    'StrSQL = "SELECT NOTES.NoteID, NOTES.Employee_ID, NOTES.NoteType, NOTES.NoteDate," & _
    '    "NOTES.Value, NOTES.Remark, NOTES.Chique_Serial_No, NOTES.Transaction_Header_ID," & _
    '    "NOTES.Dealer_Code, NOTES.NotePosted, NOTES.PostedBy, NOTES.PostDate," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Value as DEV_Value, DOUBLE_ENTREY_VOUCHERS." & _
    '    "Credit_Or_Debit, DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Remark," & _
    '    "DOUBLE_ENTREY_VOUCHERS.Notes_Id,ACCOUNTS.Account_Name, EMPLOYEES.Employee_Name," & _
    '    "USERS.UserName AS UserIssued, USERS_1.UserName AS UserPosted ,ACCOUNTS.Account_Serial "
    'StrSQL = StrSQL + " FROM (EMPLOYEES RIGHT JOIN ((USERS INNER JOIN NOTES ON USERS.User_ID = " & _
    '    "NOTES.Issued_BY) LEFT JOIN USERS AS USERS_1 ON NOTES.PostedBy = USERS_1.User_ID) " & _
    '    "ON EMPLOYEES.Employee_Code = NOTES.Employee_ID) INNER JOIN  " & _
    '    "(ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code =  " & _
    '    "DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_Id" & _
    '    " where NOTES.Note_ID='" & Me.TxtNoteID.text & "'" & _
    '    " ORDER BY Val(DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No);"
    'If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.Path & "\Reports\Journal.rpt"
    'Else
    '    StrFileName = App.Path & "\Reports\Journal_Eng.rpt"
    'End If
    'If Dir(StrFileName) = "" Then
    '    GetMsgs 139, vbExclamation
    '    FireReport = False
    '    Exit Function
    'End If
    'RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    'If RsData.BOF Or RsData.EOF Then
    '    GetMsgs 138, vbExclamation
    '    FireReport = False
    '    RsData.Close
    '    Set RsData = Nothing
    '    Exit Function
    'End If
    'Screen.MousePointer = vbArrowHourglass
    'Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData
    'Rs.Open "Options", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    'xReport.ParameterFields(1).AddCurrentValue Rs("Company_Name_Arabic").Value
    'xReport.ParameterFields(2).AddCurrentValue Rs("Comment_Arabic").Value
    'xReport.ParameterFields(3).AddCurrentValue Rs("Company_Name_Eng").Value
    'xReport.ParameterFields(4).AddCurrentValue Rs("Comment_Eng").Value
    'xReport.ParameterFields(5).AddCurrentValue StrUserName
    'If SystemOptions.UserInterface = ArabicInterface Then
    '     xReport.ReportTitle = "ÿ»«⁄… ﬁÌœ «·ÌÊ„Ì… —ﬁ„ " & Me.TxtSerial.text
    'Else
    '     xReport.ReportTitle = "Journal Voucher NO." & Me.TxtSerial.text
    'End If
    'xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    '
    ''xReport.PaperSize=
    'If Not (IsNull(Rs("DefaultPrinter").Value)) Then
    '    StrPrinterName = Rs("DefaultPrinter").Value
    '    For I = 0 To Printers.count - 1
    '        If StrPrinterName = Printers(I).DeviceName Then
    '            Set XPrinter = Printers.Item(I)
    '            Exit For
    '        End If
    '    Next I
    '    If Not XPrinter Is Nothing Then
    '        xReport.SelectPrinter XPrinter.DriverName, XPrinter.DeviceName, XPrinter.Port
    '    End If
    'End If
    '
    'Set Frm = New FrmPrint
    'With Frm
    '    .CRViewerMain.ReportSource = xReport
    '    Do While .CRViewerMain.IsBusy
    '        DoEvents
    '    Loop
    '    .CRViewerMain.Zoom IIf(IsNull(Rs("RptZoom").Value), 100, Rs("RptZoom").Value)
    '    If m_Destination = WindowTarget Then
    '        .CRViewerMain.ViewReport
    '        .WindowState = vbMaximized
    '    Else
    '        'xReport.PrintOut "⁄œœ «·‰”Œ", 12
    '        xReport.PrintOut
    '        .CRViewerMain.PrintReport
    '    End If
    '
    '    If m_Destination = WindowTarget Then
    '        .Show
    '    Else
    '        Unload Frm
    '    End If
    'End With
    'Set xApp = Nothing
    'Set xReport = Nothing
    ''SendCrystalSetting cr, "ﬁÌÊœ «·ÌÊ„Ì…"
    'FireReport = True
    'Screen.MousePointer = vbDefault
    'Exit Function
    'FireReportErrTrap:
    'FireReport = False
    'Screen.MousePointer = vbDefault
End Function

Private Sub ChangeLang()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Command6.Caption = "Show"
    Frame2.Caption = "Acc Show Type"
    lbl(15).Caption = "FileName"
    CMDSelectFile.Caption = "Select File."
    CmdImport.Caption = "Import File"
    Me.Caption = " Load Trial Balance"
    
    Me.EleTop.Caption = Me.Caption
    lbl(9).Caption = "Date"
    lbl(10).Caption = "Notes"
    lbl(11).Caption = "Accounts View"
    lbl(12).Caption = "Accounts Sort "
       lbl(14).Caption = "Manual No."
    CmdRemove.Caption = "Remove Line"
    CmdAdd.Caption = "Add Line"
    Label14.Caption = "Eng DES"
    Frame3.Caption = "Enter Voucher No. To copy it"
    Label7.Caption = "Voucher #"
    Command5.Caption = "Copy"
    Label8.Caption = "General C.C."
    Label9.Caption = "Project"
    Option1.Caption = "Depit"
    Option2.Caption = "Credit"
    Cmd(8).Caption = "Delete"
    Auto_cost_center.Caption = "Auto C.C."
    Label11.Caption = "General Branch"
    Frame1.Caption = "Copy From JL"
    Label12.Caption = "No:"
    Command5.Caption = "Copy"

    'Rs.Open "Lang", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    'Rs.MoveFirst
    'For I = Me.lbl.LBound To Me.lbl.UBound
    '    If Trim(lbl(I).Tag) <> "" Then
    '        Rs.MoveFirst
    '        Rs.find "ID=" & Val(Me.lbl(I).Tag) & "", , adSearchForward, 1
    '        If Not (Rs.BOF Or Rs.EOF) Then
    '            Me.lbl(I).Caption = IIf(IsNull(Rs("Eng").value), "", Rs("Eng").value) & ":"
    '        End If
    '    End If
    'Next I
    'Rs.Close
    'Set Rs = Nothing
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Label1.Caption = "Source"
    Label2.Caption = "Based ON"

    lbl(7).Caption = "ID"
    lbl(0).Caption = "Date"
    lbl(3).Caption = "Serial"
    lbl(4).Caption = "Value"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Modify"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Insert"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    Label13.Caption = "Des"
    TabMain.TabCaption(0) = "Journal"
    TabMain.TabCaption(1) = "Comment"
    ElePost.Caption = "Posting State"
    ChkPost.Caption = "Voucher State"
    Check3.Caption = "Template"
    Check2.Caption = "Approved"
    Check1.Caption = "Cancel Action"
    Check5.Caption = "Deleted"
    Check4.Caption = "Periodic"
    lbl(1).Caption = "Depit Sum"
    lbl(2).Caption = "Credit Sum"
    lbl(13).Caption = "Result"
    lbl(8).Caption = "By"
    lbl(5).Caption = "Signature"
    ALLButton1.Caption = "Cost Center"
    ALLButton20.Caption = "Approved"
    ALLButton3.Caption = "Call Repeated Vchr."
    ALLButton6.Caption = "Create Repeated Vchr."
    ALLButton7.Caption = "template"
    ALLButton10.Caption = "Insert template"
    ALLButton8.Caption = "Cancel Action"
    ALLButton9.Caption = "Perview"
    ALLButton2.Caption = "Attachments"

    Command1.Caption = "Add to Explain Template"
    Command2.Caption = "Call Explain Template"

    EleOpt.Caption = "Show Of Accounts"
    Opt(0).Caption = "Hierarchy "
    Opt(1).Caption = "Parent Path "
    Opt(2).Caption = "Tabular "
    ChkLastAccount.Caption = "Show Last Accounts Only"
    OptSort(0).Caption = "A-Z"
    OptSort(1).Caption = "Chart Sort"

    With Fg_Journal
        .Cell(flexcpText, 0, .ColIndex("LineNo"), 1, .ColIndex("LineNo")) = "Line NO."
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = "Current Currency value"
        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("DebitValueE"), 0, .ColIndex("CreditValueE")) = "Forign Currency value"
        .Cell(flexcpText, 1, .ColIndex("DebitValueE"), 1, .ColIndex("DebitValueE")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValueE"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        '  .Cell(flexcpText, 0, .ColIndex("DebitValuee"), 0, .ColIndex("CreditValueE")) = "ValueE"
        '   .Cell(flexcpText, 1, .ColIndex("DebitValuee"), 1, .ColIndex("DebitValueE")) = "Debit"
        '   .Cell(flexcpText, 1, .ColIndex("CreditValuee"), 1, .ColIndex("CreditValueE")) = "Credit"
    
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "Account Name"
        .Cell(flexcpText, 0, .ColIndex("Des"), 1, .ColIndex("Des")) = "Comment A"
        .Cell(flexcpText, 0, .ColIndex("DesE"), 1, .ColIndex("DesE")) = "Comment E"
    
        .Cell(flexcpText, 0, .ColIndex("currenct_code"), 1, .ColIndex("currenct_code")) = "currency"
     
        .Cell(flexcpText, 0, .ColIndex("rate"), 1, .ColIndex("rate")) = "rate"
        .Cell(flexcpText, 0, .ColIndex("BranchName"), 1, .ColIndex("BranchName")) = "BranchName"
        .Cell(flexcpText, 0, .ColIndex("CC"), 1, .ColIndex("CC")) = "CC"
       
       .Cell(flexcpText, 0, .ColIndex("Departement"), 1, .ColIndex("Departement")) = "Departement"
        .Cell(flexcpText, 0, .ColIndex("NEmpName"), 1, .ColIndex("NEmpName")) = "NEmpName"
        .Cell(flexcpText, 0, .ColIndex("FixedAsset"), 1, .ColIndex("FixedAsset")) = "Equipments"
  
  
    End With

    LblDes.Caption = "Write your comment."
End Sub

Private Sub AddTip()

    Dim Wrap As String
    Dim Msg As String

    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "—ﬁ„ «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "—ﬁ„ «·ﬁÌœ «·Œ«’ »«·„” ‰œ"
            .AddControl TxtDEV_NO, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "„”·”·", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "„”·”· Â–« «·„” ‰œ ›Ï  Õ—Ì— «·ﬁÌÊœ"
            .AddControl TxtSerial, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "ﬁÌ„… «·”‰œ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "«·ﬁÌ„… «·√Ã„«·Ì… ··ﬁÌœ"
            .AddControl TxtValue, Msg, True
        End With

        With TTP
            .Create Me.hwnd, " «—ÌŒ «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = " «—ÌŒ  Õ—Ì— «·ﬁÌœ." & Wrap & "≈› —«÷Ì« ÌﬂÊ‰  «—ÌŒ «·ÌÊ„."
            .AddControl DTP_Date, Msg, True
        End With

        With TTP
            .Create Me.hwnd, " ⁄·Ìﬁ ⁄·Ï «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Ì„ﬂ‰ﬂ Â‰« ﬂ «»…  ⁄·Ìﬁ „‰«”»" & Wrap & "⁄·Ï Â–« «·Õ”«» ·ÌŸÂ— »ÃÊ«—Â" & Wrap & "›Ï ⁄„·Ì… „—«Ã⁄… «·ﬁÌÊœ √Ê " & Wrap & "«·ÿ»«⁄…."
            .AddControl TxtDes, Msg, True
        End With

        '
        With TTP
            .Create Me.hwnd, " ⁄·Ìﬁ ⁄·Ï «·ﬁÌœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "≈÷€ÿ Â‰« · ŸÂ— ·ﬂ ‰«›–…" & Wrap & " Õ—Ì— «· ⁄·Ìﬁ · ﬂ »  ⁄·Ìﬁ" & Wrap & "„‰«”» ⁄·Ï Â–« «·Õ”«»."
            .AddControl CboDes, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "⁄—÷ «·Õ”«» «·‰Â«∆Ï ›ﬁÿ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì„ﬂ‰ﬂ ÕÃ»" & Wrap & " «·Õ”«» «·—∆Ì”Ì… Ê≈ŸÂ«— «·Õ”«»« " & Wrap & "«·‰Â«∆Ì… Ê«· Ï Ì„ﬂ‰ﬂ  ”ÃÌ· " & Wrap & "«·ﬁÌÊœ ·Â«."
            .AddControl ChkLastAccount, Msg, True
        End With

        'OptSort
        With TTP
            .Create Me.hwnd, Opt(1).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ÃœÊ·Ï Ì⁄—÷ «”„ «·Õ”«» «·‰Â«∆Ï Ê«”„" & Wrap & "«·Õ”«» «·„ ›—⁄ „‰Â Ê«Ì÷« «”„ «·Õ”«» " & Wrap & "«·√⁄·Ï „‰Â( À·«À… „” ‰ÊÌ« )."
            .AddControl Opt(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, Opt(2).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ÃœÊ·Ï Ì⁄—÷ «”„ «·Õ”«» ›ﬁÿ."
            .AddControl Opt(2), Msg, True
        End With

        With TTP
            .Create Me.hwnd, Opt(0).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· ⁄—÷ «”„«¡ «·Õ”«»«  «· Ï " & Wrap & "Ì„ﬂ‰ﬂ ﬂ «»… Ê ”ÃÌ· «·ﬁÌœ ·Â«  ŸÂ— ›Ï " & Wrap & "‘ﬂ· ‘Ã—Ï »«·Ÿ»ÿ „À· «·œ·Ì· «·„Õ«”»Ï."
            .AddControl Opt(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, OptSort(1).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· «”„«¡ «·Õ”«»« " & Wrap & " „— »… Õ”» „Êﬁ⁄Â« Ê — Ì»Â« " & Wrap & "««·œ·Ì· «·„Õ«”»Ï »«·Ÿ»ÿ. "
            .AddControl OptSort(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, OptSort(0).Caption, 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌÃ⁄· «”„«¡ «·Õ”«»« " & Wrap & " „— »…  —ÌÌ»« √»ÃœÌ« »€÷ " & Wrap & "«·‰Ÿ— ⁄‰ „Êﬁ⁄Â« ›Ï «·œ·Ì·" & Wrap & "«·„Õ«”»Ï."
            .AddControl OptSort(0), Msg, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "DEV NO.", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The serial of double entery voucher "
            .AddControl TxtDEV_NO, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Serial", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The Serial of the voucher in the " & Wrap & "editing journals transactions"
            .AddControl TxtSerial, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Voucher Value", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "The total talue which will be" & Wrap & "recorded"
            .AddControl TxtValue, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Date", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Data of editing the voucher" & Wrap & "by default it is current ." & Wrap & "system date."
            .AddControl DTP_Date, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Comment", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Write your comment here to" & Wrap & " appear in auditing journal" & Wrap & "screen or in auditing report "
            .AddControl TxtDes, Msg, False
        End With

        '
        With TTP
            .Create Me.hwnd, "Write comment", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Click here to show the " & Wrap & "editing window to write" & Wrap & "your comment."
            .AddControl CboDes, Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkLastAccount.Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option when enabled shows" & Wrap & "the last accounts only."
            .AddControl ChkLastAccount, Msg, False
        End With

        'OptSort
        With TTP
            .Create Me.hwnd, Opt(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in tabluar form !! and display " & Wrap & "the last three levels of chart" & Wrap & "of accounts."
            .AddControl Opt(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, Opt(2).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in tabluar form !! and display" & Wrap & "just only the last account."
            .AddControl Opt(2), Msg, False
        End With

        With TTP
            .Create Me.hwnd, Opt(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts" & Wrap & "in hierarchy view exactly like" & Wrap & "the view of chart of accounts."
            .AddControl Opt(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, OptSort(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option shows the accounts " & Wrap & "sorted by it is index in the" & Wrap & "chart of accounts "
            .AddControl OptSort(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, OptSort(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This Option shows the accounts" & Wrap & "sorted alphabetically regardless " & Wrap & "it is index in the chart of " & Wrap & "accounts."
            .AddControl OptSort(0), Msg, False
        End With

    End If

End Sub

Public Function RefreshData() As Boolean

End Function

Public Property Get Cmd_Preview() As Boolean

    If Me.TXTNoteID.Text = "" Then
        GetMsgs 140, vbExclamation
        Cmd_Print = False
    Else
        Cmd_Print = FireReport(WindowTarget)
    End If

End Property

Public Property Let Cmd_Preview(ByVal vNewValue As Boolean)
    m_Cmd_Preview = vNewValue
End Property

Private Sub SaveData()
    Dim TransBegine As Boolean
    Dim Msg As String
    Dim i As Long
    Dim StrSQL As String
    Dim RsTemp  As New ADODB.Recordset
    Dim RsNetes As New ADODB.Recordset
    Dim RsDev As New ADODB.Recordset
    Dim IntNoteType As Integer
    Dim StrInsertSQL  As String
    Dim IntAutoAccPost As Integer
    Dim StrPost As String
    Dim StrUnPost As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrPost = "„—Õ·"
        StrUnPost = "€Ì— „—Õ·"
    Else
        StrPost = "Posted"
        StrUnPost = "Not Posted"
    End If

    'On Error GoTo ErrTrap

    If val(TxtValue.Text) = 0 Then
        TxtValue.Text = 0
        '  Msg = "„‰ ›÷·ﬂ ﬁ„ »≈œŒ«· ﬁÌ„… «·”‰œ"
        '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '  'GetMsgs 59, vbExclamation
        '  TxtValue.SetFocus
        '  Exit Sub
    End If

    With Fg_Journal

        i = .FixedRows

        Do While i <= .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) = "" Or .TextMatrix(i, .ColIndex("Account_Serial")) = "" Then
                .RemoveItem i
                i = i
            Else
                i = i + 1
            End If
 If .Rows > 2 Then
            Me.TxtTotalDebit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
            Me.TxtTotalCredit.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
 End If
 
            Me.TXTResults.Text = val(Me.TxtTotalDebit.Text) - val(Me.TxtTotalCredit.Text)
            Me.TxtTotalDebit.Text = Round(Me.TxtTotalDebit.Text, SystemOptions.SysDefCurrencyForamt)
            Me.TxtTotalCredit.Text = Round(Me.TxtTotalCredit.Text, SystemOptions.SysDefCurrencyForamt)


        Loop

        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 And val(.TextMatrix(i, .ColIndex("CreditValue"))) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                
                        Msg = "«·Õ”«» " & .TextMatrix(i, .ColIndex("AccountName")) & CHR(13)
                        Msg = Msg & "·„  Õœœ ·Â Â· ÂÊ ÿ—› œ«∆‰ √Ê „œÌ‰.øø!!" & CHR(13)
                        Msg = Msg & "»—Ã«¡ ﬂ «»… ﬁÌ„… –·ﬂ «·Õ”«»"
                        Msg = Msg & " ”ÿ— —ﬁ„" & .TextMatrix(i, .ColIndex("lineno"))
                
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Else
                        Msg = "The Account " & .TextMatrix(i, .ColIndex("AccountName")) & CHR(13)
                        Msg = Msg & "not set as a Credit Or as Debit.??" & CHR(13)
                        Msg = Msg & "Please Write this account value.!"
                        MsgBox Msg, vbExclamation, App.title
                    End If
        Fg_Journal.Row = i
                            Fg_Journal.Col = Fg_Journal.ColIndex("DebitValue")
                            Fg_Journal.ShowCell i, Fg_Journal.ColIndex("DebitValue")
                            
                            Fg_Journal.SetFocus
                            
                            
                    Exit Sub
                End If
            End If

        Next i

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If check_cost_center(i) = False Then
                    Exit Sub
                End If
            End If

        Next i

    End With

    If Round(Me.TxtTotalCredit.Text, 1) <> Round(Me.TxtTotalDebit.Text, 1) Then

        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Depit And Credit not matched ..!!" & CHR(13)
            Msg = Msg & "please correct this error."
        Else
            Msg = "ÿ—›Ï «·ﬁÌœ €Ì— „ “‰Ì‰ ..!!" & CHR(13)
            Msg = Msg & "„‰ ›÷·ﬂ ﬁ„ »„—«Ã⁄… ÿ—›Ï «·ﬁÌœ."
        End If

        'GetMsgs 60, vbExclamation
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    'If Val(Me.TxtValue.text) <> Val(Me.TxtTotalDebit.text) Then
    '    Msg = "ﬁÌ„… «·”‰œ €Ì— „ﬁ»Ê·… ..!!" & Chr(13)
    '    Msg = Msg & "„‰ ›÷·ﬂ ﬁ„ »„—«Ã⁄… ÿ—›Ï «·ﬁÌœ."
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    'GetMsgs 61, vbExclamation
    '    Exit Sub
    'End If
    '---------------------------Get the serial--------------
    If Me.TxtModFlg.Text = "N" Then
        ' Me.TxtSerial.text = ModAccounts.GetNewDEV_Serial(Me.DTP_Date.value)
    End If

    IntNoteType = 20



    If Me.TxtSerial.Text = "" Then
        my_branch = val(Me.dcBranch.BoundText)
  
        If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "can't Add new voucher because you exceed the numbering  ": Exit Sub
            Else
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            End If
 
        Else

            If Notes_coding(val(my_branch), DTP_Date.value) = "" Then
                If TxtSerial.Text = "" Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        MsgBox "Enter Voucher code ": Exit Sub
                    Else
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·ﬁÌœ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                 
                    End If
                End If

            Else
  
                TxtSerial.Text = Notes_coding(val(my_branch), DTP_Date.value)
          
            End If
        End If
    End If

    If TxtSerial.Text = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Enter Voucher code ": Exit Sub
        Else
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·ﬁÌœ    ": Exit Sub
                 
        End If

        Exit Sub
    End If
    Cn.BeginTrans
    TransBegine = True
    
    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete   Notes Where Notes.NoteID='" & Trim(TXTNoteID.Text) & "'"
        Cn.Execute StrSQL, , adExecuteNoRecords
    
        If DcCostCenter.BoundText <> "" Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
    
    ElseIf Me.TxtModFlg.Text = "N" Then
        '---------------------------Get The Note ID ------------
        Me.TXTNoteID.Text = CStr(new_id("notes", "NoteID", ""))
        Me.TxtDEVID.Text = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
        Me.TxtDEV_NO.Text = Me.TxtDEVID.Text
        Me.oldTxtSerial.Text = Trim$(Me.TxtSerial.Text)
    
        '---------------------------Begine of Saving------------
    End If

    Set RsNetes = New ADODB.Recordset
    'RsNetes.Open "NOTES", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT      * from dbo.NOTES Where (1 = -1)"
   RsNetes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    RsNetes.AddNew
    RsNetes("branch_no").value = val(Me.dcBranch.BoundText)
    RsNetes("salary").value = val(Me.txt_salary.Text)
    RsNetes("NoteID").value = val(Me.TXTNoteID.Text)
    RsNetes("NoteType").value = 200
    RsNetes("NoteSerial").value = (Me.TxtSerial.Text)
      RsNetes("ManualNo").value = (Me.TxtManualNO.Text)
      
    RsNetes("OldNoteSerial1").value = (Me.oldTxtSerial.Text)  '
    RsNetes("numbering_type").value = sand_numbering_type(0) ' numbering_type
    RsNetes("sanad_year").value = year(DTP_Date.value)
    RsNetes("sanad_month").value = Month(DTP_Date.value)
    RsNetes("foxy_no").value = val(Text1.Text)
     RsNetes("NoteDate").value = Me.DTP_Date.value
    
 '   RsNetes("NoteDate").value = Format$(Date, "dd-mm-yyyy")
    RsNetes("NoteDateH").value = Me.Txt_DateHigri.value
     
    RsNetes("Note_Value").value = val(Me.TxtValue.Text)
    RsNetes("Double_Entry_Vouchers_ID").value = val(Me.TxtDEVID.Text)
    RsNetes("DAWRY").value = Check4.value
    RsNetes("KALEB").value = Check3.value
    RsNetes("auto_des").value = Me.Auto_cost_center.value
    
    RsNetes("Remark").value = Trim$(Me.Txt.Text)
    RsNetes("RemarkE").value = Trim$(Me.Txte.Text)
    RsNetes("UserID").value = val(Me.DcboUsers.BoundText)
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtTotalDebit.Text, "0.00"), 0, True, ".")
    RsNetes("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)
    RsNetes("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    
    If Me.dcprojects.BoundText <> "" Then
        Dim project_id As Integer
        project_id = IIf(Me.dcprojects.BoundText = "", 0, Me.dcprojects.BoundText)
        RsNetes("project_id").value = project_id
        Dim project_depit_or_credit As Integer
    
        If Option1.value = True Then
            project_depit_or_credit = 0
        Else
            project_depit_or_credit = 1
        End If
    
        RsNetes("project_depit_or_credit").value = project_depit_or_credit
    
    End If
    
    RsNetes.update
    Dim valuee As Variant

    With Fg_Journal

        For i = .FixedRows To .Rows - 1
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Variant

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                If val(.TextMatrix(i, .ColIndex("DebitValue"))) > 0 Then
                    IntDEV_Type = 0
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("DebitValue")))
                Else
                    IntDEV_Type = 1
                    SngDEV_Value = val(.TextMatrix(i, .ColIndex("CreditValue")))
                End If
            
                project_id = IIf(Me.dcprojects.BoundText = "", 0, Me.dcprojects.BoundText)
            
                If IntDEV_Type = 0 And Option1.value = True Then
               
                ElseIf IntDEV_Type = 1 And Option2.value = True Then
            
                Else
                    project_id = 0
                End If
            
                If val(.TextMatrix(i, .ColIndex("DebitValuee"))) > 0 Then
               
                    valuee = val(.TextMatrix(i, .ColIndex("DebitValuee")))
                Else
                 
                    valuee = val(.TextMatrix(i, .ColIndex("CreditValuee")))
                End If

                ' CStr(.Cell(flexcpData, I, .ColIndex("Des")))
                If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(val(Me.dcBranch.BoundText) = 0, 1, val(Me.dcBranch.BoundText))
                End If
                    
                      
                       
            
                If ModAccounts.AddNewDev(val(Me.TxtDEVID.Text), .TextMatrix(i, .ColIndex("LineNo")), .TextMatrix(i, .ColIndex("AccountCode")), SngDEV_Value, IntDEV_Type, .TextMatrix(i, .ColIndex("des")), val(Me.TXTNoteID.Text), , , SystemOptions.SysCurrentAccountIntervalID, Me.DTP_Date.value, val(.TextMatrix(i, .ColIndex("userid"))), , Me.TxtSerial.Text, , valuee, .TextMatrix(i, .ColIndex("currenct_code")), val(.TextMatrix(i, .ColIndex("rate"))), , .TextMatrix(i, .ColIndex("dese")), IIf(val(.TextMatrix(i, .ColIndex("LineNo1"))) <> 0, val(.TextMatrix(i, .ColIndex("LineNo1"))), setfoxy_Line), , project_id, , , , val(.TextMatrix(i, .ColIndex("FixedassetId"))), , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("Departementid"))), val(.TextMatrix(i, .ColIndex("NEmpid")))) = False Then
                    GoTo ErrTrap
                End If
            End If

        Next i

    End With

    Cn.CommitTrans
    TransBegine = False
    CuurentLogdata

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "Saved"
    Else
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ"
    End If

    lbl(27).Caption = showLabel(TxtSerial, oldTxtSerial)

    'Õ›Ÿ „—ﬂ“ «· ﬂ·›… «·⁄«„
    '   If Me.DcCostCenter.BoundText <> "" Then
    save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "”‰œ ﬁÌœ", Me.DTP_Date.value
    save_cost_center

    '   End If
    'Õ›Ÿ „—ﬂ“ «· ﬂ·›… «·„Ê“⁄Â «·Ì«
    'If Me.Auto_cost_center.value = vbChecked Then
    save_Auto_cost_center "”‰œ ﬁÌœ", Me.DTP_Date.value
    save_cost_center
    
    'End If
    
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.TxtModFlg.Text = "R"
    '------------------------End of Saving--------------
    Exit Sub
ErrTrap:

    If TransBegine = True Then
        Cn.RollbackTrans
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        Msg = "error During Saving"
    Else
        Msg = "⁄›Ê« ... ÕœÀ Œÿ« «À‰«¡ ⁄„·Ì… «·Õ›Ÿ."
    End If

    'Msg = Msg & Chr(13) & Err.Remark
    MsgBox Msg, vbExclamation, App.title
End Sub

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.Text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.Text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = DTP_Date.value
        rs("NoteSerial").value = TxtSerial.Text
        ' rs("Remark").value = Txt.text
        rs("Remark").value = "”‰œ ﬁÌœ   »—ﬁ„ " & TxtSerial.Text & "    " & Me.TxtDes
        rs.update
        rs.MoveNext
    Next i

End Function
 
Public Function save_Auto_cost_center(opr_type As String, _
                                      record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String
    StrSQL = "Delete  marakes_taklefa_temp  where auto_des=1 and  kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.Auto_cost_center.value = vbUnchecked Then
        'Exit Function
    End If
 
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And .TextMatrix(i, .ColIndex("cost_center_id")) <> "" Then
                'Õ«·…  Ê“Ì⁄ „—«ﬂ“ «· ﬂ·›… «·Ì«
     
                rs.AddNew
                rs("cost_center_id").value = .TextMatrix(i, .ColIndex("cost_center_id"))
                rs("cost_center").value = get_COST_CENTER_NAME(.TextMatrix(i, .ColIndex("cost_center_id")), "account_name")

                If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 Then
                    rs("value").value = .TextMatrix(i, .ColIndex("CreditValue"))
                    rs("depit_or_credit").value = "œ«∆‰"
            
                Else
                    rs("value").value = .TextMatrix(i, .ColIndex("DebitValue"))
                    rs("depit_or_credit").value = "„œÌ‰"
            
                End If
        
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("general_des").value = 0
                rs("auto_des").value = 1
        
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs("NoteDate").value = DTP_Date.value
                rs("NoteSerial").value = TxtSerial.Text
                rs("Remark").value = Txt.Text
 
                rs.update
            End If
   
        Next i

    End With

    rs.Close
End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String
    StrSQL = "Delete  marakes_taklefa_temp  where   general_des=1 and kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
 
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center

                If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 Then
                    rs("value").value = .TextMatrix(i, .ColIndex("CreditValue"))
                    rs("depit_or_credit").value = "œ«∆‰"
            
                Else
                    rs("value").value = .TextMatrix(i, .ColIndex("DebitValue"))
                    rs("depit_or_credit").value = "„œÌ‰"
            
                End If
        
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("general_des").value = 1
        
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs.update

                'Õ«·…  Ê“Ì⁄ „—«ﬂ“ «· ﬂ·›… «·Ì«
                If Auto_cost_center.value = vbChecked Then
                    rs.AddNew
                    rs("cost_center_id").value = cost_center_id
                    rs("cost_center").value = get_COST_CENTER_NAME(cost_center_id, "account_name")

                    If val(.TextMatrix(i, .ColIndex("DebitValue"))) = 0 Then
                        rs("value").value = .TextMatrix(i, .ColIndex("CreditValue"))
                        rs("depit_or_credit").value = "œ«∆‰"
            
                    Else
                        rs("value").value = .TextMatrix(i, .ColIndex("DebitValue"))
                        rs("depit_or_credit").value = "„œÌ‰"
            
                    End If
        
                    rs("opr_id").value = Me.Text1.Text
                    rs("kedno").value = Me.Text1.Text
                    rs("general_des").value = 0
                    rs("auto_des").value = 1
        
                    rs("opr_type").value = opr_type
                    rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                    rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                    rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                    rs("record_date").value = record_date
                    rs.update
                End If
        
            End If

        Next i

    End With

    rs.Close
End Function

Private Sub TXTResults_Change()
    On Error Resume Next
    Me.TXTResults.Text = Round(Me.TXTResults.Text, 2)

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    'Dim rs As New ADODB.Recordset

    Dim StrSQL As String

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     " From notes where (((notes.NoteType) =200)) " & _
     " ORDER BY NOTES.NoteID "
    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     "From notes where (((notes.NoteType)=200)) " & _
     "    ORDER BY NOTES.NoteID "
    
    'StrSQL = "SELECT NOTES.NoteID, NOTES.NoteType " & _
     "From notes      ORDER BY NOTES.NoteID  "
    'StrSQL = " SELECT  Noteserial From gl_cc   group by   Noteserial     ORDER BY  Noteserial   where notetype<>101 "
    
    
'    StrSQL = "SELECT  Noteserial  From gl_cc    where notetype <>1000  group by    Noteserial     ORDER BY  Noteserial"

'    If SystemOptions.usertype <> UserAdminAll Then
'        StrSQL = "SELECT  Noteserial  From gl_cc    where branch_no=" & branch_id & " and  notetype <>1000  group by    Noteserial     ORDER BY  Noteserial"
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    If StrOldTransID <> "" Then
        rs.find "Noteserial='" & StrOldTransID & "'", , adSearchForward, 1

        If rs.BOF Or rs.EOF Then
            rs.MoveFirst
        End If

    Else
        rs.MoveFirst
    End If

    Select Case Index

        Case 1 'First

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst
            End If

        Case 0 'Previous

            If Not (rs.BOF Or rs.EOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveNext
            End If

        Case 3 'NEXT

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MovePrevious
            End If

        Case 2 'Last

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

    End Select

    If Not (IsNull(rs("Noteserial").value)) Then
        Me.Retrive rs("Noteserial").value
        StrOldTransID = rs("Noteserial").value
    End If

'    rs.Close
'    Set rs = Nothing
End Sub

