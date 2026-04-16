VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalary5 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰ ÿ»⁄ » «—ÌŒ"
   ClientHeight    =   9240
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   18570
   HelpContextID   =   580
   Icon            =   "FrmEmpSalary5.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   18570
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
      Height          =   9240
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18570
      _cx             =   32755
      _cy             =   16298
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
         Height          =   7185
         Left            =   150
         TabIndex        =   3
         Top             =   30
         Width           =   18360
         _cx             =   32385
         _cy             =   12674
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   " Œ’Ì’ «·—« »| ’œÌ— ··»‰ﬂ|”œ«œ «·—Ê« »|ﬁÌÊœ «·”œ«œ| ›«’Ì· «·„‘«—Ì⁄| ›«’Ì· «· ‰ﬁ·« "
         Align           =   0
         CurrTab         =   4
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
         Picture(0)      =   "FrmEmpSalary5.frx":038A
         Picture(1)      =   "FrmEmpSalary5.frx":0724
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6720
            Left            =   -18915
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin VB.CommandButton CMDShow 
               Caption         =   "⁄—÷ «·ﬁÌÊœ"
               Height          =   375
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   240
               Width           =   975
            End
            Begin VSFlex8Ctl.VSFlexGrid GridGE 
               Height          =   4260
               Left            =   9120
               TabIndex        =   95
               Top             =   120
               Width           =   9015
               _cx             =   15901
               _cy             =   7514
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
               Cols            =   67
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary5.frx":0ABE
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
            Height          =   6720
            Index           =   1
            Left            =   -19815
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ ”«⁄«  «·‘Â—"
               Height          =   630
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   0
               Visible         =   0   'False
               Width           =   1545
               Begin VB.TextBox TxtMonthHours 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Text            =   "176"
                  Top             =   330
                  Width           =   705
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "»Ì«‰«  ﬁÌœ «·”œ«œ"
               Height          =   555
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   -315
               Visible         =   0   'False
               Width           =   5520
               Begin VB.TextBox txtnoteserial2 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   120
                  Width           =   1455
               End
               Begin ALLButtonS.ALLButton ALLButton5 
                  Height          =   345
                  Left            =   240
                  TabIndex        =   58
                  Top             =   120
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   609
                  BTYPE           =   2
                  TX              =   "ÿ»«⁄Â"
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
                  BCOL            =   15790320
                  BCOLO           =   15790320
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmEmpSalary5.frx":12F9
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   255
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ﬁÌœ «·«” Õﬁ«ﬁ"
               Height          =   585
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   0
               Width           =   7605
               Begin VB.TextBox txtnoteserial 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   120
                  Width           =   1455
               End
               Begin ALLButtonS.ALLButton ALLButton4 
                  Height          =   345
                  Left            =   240
                  TabIndex        =   55
                  Top             =   120
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   609
                  BTYPE           =   2
                  TX              =   "ÿ»«⁄Â  «·ﬁÌœ"
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
                  BCOL            =   14871017
                  BCOLO           =   14871017
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmEmpSalary5.frx":1315
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSComCtl2.DTPicker DTP_Date 
                  Height          =   285
                  Left            =   4680
                  TabIndex        =   62
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   120389635
                  CurrentDate     =   37140
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   69
                  Top             =   -480
                  Visible         =   0   'False
                  Width           =   2925
                  _ExtentX        =   5159
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   114
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   120389635
                  CurrentDate     =   37140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· «—ÌŒ"
                  Height          =   225
                  Index           =   11
                  Left            =   5850
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Tag             =   "53"
                  Top             =   225
                  Width           =   780
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   255
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3615
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   135
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.CheckBox Check16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ÊﬁÌ⁄"
               Height          =   210
               Left            =   -90
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox Check15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ì"
               Height          =   210
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.CheckBox Check14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì 2"
               Height          =   210
               Left            =   2445
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.CheckBox Check13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ã“«¡« "
               Height          =   210
               Left            =   3435
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CheckBox Check12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”·›"
               Height          =   210
               Left            =   4620
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CheckBox Check11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì1"
               Height          =   210
               Left            =   5340
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ê·« "
               Height          =   210
               Left            =   6150
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox Check9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬂ«›√ "
               Height          =   210
               Left            =   7140
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox Check8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷«›Ì"
               Height          =   225
               Left            =   9045
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   75
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox Check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ—Ï"
               Height          =   210
               Left            =   8865
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ⁄«„"
               Height          =   210
               Left            =   10125
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " „Ê«’·« "
               Height          =   210
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ· ”ﬂ‰"
               Height          =   210
               Left            =   11850
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—« » «”«”Ì"
               Height          =   210
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   315
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   165
               Left            =   14565
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   735
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„ÊŸ›"
               Height          =   195
               Left            =   16185
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   765
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   5895
               Left            =   0
               TabIndex        =   60
               Top             =   720
               Width           =   18180
               _cx             =   32067
               _cy             =   10398
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
               Cols            =   83
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEmpSalary5.frx":1331
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
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„·ÕÊŸ… : «÷€ÿ ⁄·Ï «”„ «·„ÊŸ› ·„‘«Âœ… „·›…"
               ForeColor       =   &H000000FF&
               Height          =   465
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   0
               Width           =   4800
            End
            Begin VB.Image ImgFavorites 
               Height          =   360
               Left            =   6150
               Picture         =   "FrmEmpSalary5.frx":1E22
               Stretch         =   -1  'True
               Top             =   120
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6720
            Index           =   2
            Left            =   -19515
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin MSDataListLib.DataCombo DCmboEmp 
               Height          =   315
               Left            =   6060
               TabIndex        =   6
               Top             =   90
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               _Version        =   393216
               Text            =   "DataCombo1"
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
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   90
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6720
            Index           =   4
            Left            =   -19215
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   255
               Left            =   0
               TabIndex        =   112
               Top             =   0
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«Œ›«¡ «·«”„ ⁄—»Ì"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.CommandButton Command4 
               Caption         =   " ’œÌ— ‰„Ê–Ã «·œ›⁄"
               Height          =   330
               Left            =   16725
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   6390
               Width           =   1455
            End
            Begin VB.CommandButton Command3 
               Caption         =   " ’œÌ— ‰„Ê–Ã «·»‰ﬂ"
               Height          =   330
               Left            =   7965
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   6390
               Width           =   1440
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   75
               Width           =   1080
            End
            Begin VB.ComboBox CboPaymentType 
               Height          =   315
               ItemData        =   "FrmEmpSalary5.frx":5A8A
               Left            =   11940
               List            =   "FrmEmpSalary5.frx":5A9D
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   120
               Width           =   1260
            End
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   225
               Left            =   16545
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   120
               Width           =   1185
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   8925
               TabIndex        =   35
               Top             =   120
               Width           =   2355
               _ExtentX        =   4154
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
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   8925
               TabIndex        =   36
               Top             =   120
               Width           =   2355
               _ExtentX        =   4154
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
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   300
               Left            =   4875
               TabIndex        =   42
               Top             =   75
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   529
               _Version        =   393216
               Format          =   120389633
               CurrentDate     =   39614
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   435
               Left            =   1530
               TabIndex        =   46
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   767
               BTYPE           =   2
               TX              =   "”œ«œ «·—« »"
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmEmpSalary5.frx":5AC4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid1 
               Height          =   5790
               Left            =   9585
               TabIndex        =   61
               Top             =   525
               Width           =   8595
               _cx             =   15161
               _cy             =   10213
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
               Cols            =   64
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary5.frx":5AE0
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
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   255
               Left            =   14295
               TabIndex        =   64
               Top             =   120
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   450
               _Version        =   393216
               Format          =   120389633
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DCAccount 
               Height          =   315
               Left            =   8925
               TabIndex        =   66
               Top             =   120
               Width           =   2355
               _ExtentX        =   4154
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
            Begin VSFlex8Ctl.VSFlexGrid Grid2 
               Height          =   5790
               Left            =   0
               TabIndex        =   90
               Top             =   525
               Width           =   9495
               _cx             =   16748
               _cy             =   10213
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
               Cols            =   69
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary5.frx":62A8
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
               RightToLeft     =   0   'False
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
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„·ÕÊŸ… : «÷€ÿ ⁄·Ï «”„ «·„ÊŸ› ·„‘«Âœ… „·›…"
               ForeColor       =   &H000000FF&
               Height          =   210
               Left            =   10575
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   6390
               Width           =   4080
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   14
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   120
               Width           =   1545
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               Height          =   225
               Index           =   13
               Left            =   3855
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   120
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·”œ«œ"
               Height          =   225
               Index           =   12
               Left            =   15465
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   135
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
               Height          =   225
               Index           =   10
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   120
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   225
               Index           =   9
               Left            =   8010
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   120
               Width           =   825
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ“Ì‰…"
               Height          =   225
               Index           =   8
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   120
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   225
               Index           =   7
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   120
               Width           =   1005
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6720
            Left            =   45
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin VB.CommandButton Command5 
               Caption         =   "ÿ»«⁄…  ›«’Ì· «·„‘«—Ì⁄"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   6840
               Width           =   1860
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   6540
               Left            =   510
               TabIndex        =   99
               Top             =   0
               Width           =   17295
               _cx             =   30506
               _cy             =   11536
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
               FormatString    =   $"FrmEmpSalary5.frx":6AF8
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6720
            Left            =   19005
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   45
            Width           =   18270
            _cx             =   32226
            _cy             =   11853
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
            Begin VB.CommandButton Command6 
               Caption         =   "ÿ»«⁄…  ›«’Ì· «·„‘«—Ì⁄"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   6840
               Visible         =   0   'False
               Width           =   1860
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
               Height          =   6540
               Left            =   120
               TabIndex        =   103
               Top             =   120
               Width           =   18015
               _cx             =   31776
               _cy             =   11536
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
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary5.frx":6D0C
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1965
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7245
         Width           =   18345
         _cx             =   32359
         _cy             =   3466
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
         Begin VB.CommandButton Command7 
            BackColor       =   &H008080FF&
            Caption         =   "«ﬁ›«· «·„”Ì—"
            Height          =   345
            Left            =   5280
            MaskColor       =   &H00000040&
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   840
            Width           =   1545
         End
         Begin VB.ComboBox cboPayType 
            Height          =   315
            ItemData        =   "FrmEmpSalary5.frx":6FAB
            Left            =   9360
            List            =   "FrmEmpSalary5.frx":6FAD
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1560
            Width           =   3270
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1215
            Index           =   3
            Left            =   6960
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   480
            Width           =   2055
            _cx             =   3625
            _cy             =   2143
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
            Caption         =   "≈Œ Ì«— «· «—ÌŒ"
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
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   240
               Width           =   1755
            End
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   540
               Width           =   1755
            End
            Begin ImpulseButton.ISButton CmdOk 
               Height          =   300
               Left            =   90
               TabIndex        =   88
               Top             =   840
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷  "
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
               ButtonImage     =   "FrmEmpSalary5.frx":6FAF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   45
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1800
               Width           =   1755
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   15
               Index           =   0
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   1875
               Width           =   1755
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   345
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   180
            Width           =   1665
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Õ–› «·ﬁÌœ"
            Enabled         =   0   'False
            Height          =   345
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1335
            Width           =   1545
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "FrmEmpSalary5.frx":7349
            Left            =   120
            List            =   "FrmEmpSalary5.frx":7386
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   480
            Width           =   3075
         End
         Begin MSDataListLib.DataCombo Dcemp 
            Height          =   315
            Left            =   9360
            TabIndex        =   9
            Top             =   1215
            Width           =   2385
            _ExtentX        =   4207
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   405
            Left            =   30
            TabIndex        =   2
            Top             =   1005
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   714
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
            ButtonImage     =   "FrmEmpSalary5.frx":7464
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   345
            Left            =   2130
            TabIndex        =   8
            Top             =   180
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·‘«‘…"
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
            ButtonImage     =   "FrmEmpSalary5.frx":77FE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcproject 
            Height          =   315
            Left            =   9360
            TabIndex        =   11
            Top             =   480
            Width           =   3270
            _ExtentX        =   5768
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   525
            Left            =   8820
            TabIndex        =   17
            Top             =   -510
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 2"
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
            ButtonImage     =   "FrmEmpSalary5.frx":7B98
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   405
            Left            =   9870
            TabIndex        =   18
            Top             =   -495
            Visible         =   0   'False
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 3 "
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
            ButtonImage     =   "FrmEmpSalary5.frx":7F32
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   405
            Left            =   0
            TabIndex        =   32
            Top             =   -75
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   714
            BTYPE           =   2
            TX              =   " ⁄œÌ· «·‘«‘…"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmEmpSalary5.frx":82CC
            PICN            =   "FrmEmpSalary5.frx":82E8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   375
            Left            =   105
            TabIndex        =   48
            Top             =   1515
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 3 "
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
            ButtonImage     =   "FrmEmpSalary5.frx":8794
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   1020
            TabIndex        =   49
            Top             =   525
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 3 "
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
            ButtonImage     =   "FrmEmpSalary5.frx":8B2E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   360
            Left            =   1755
            TabIndex        =   50
            Top             =   1305
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷  "
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
            ButtonImage     =   "FrmEmpSalary5.frx":8EC8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   345
            Left            =   5280
            TabIndex        =   72
            Top             =   1320
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            BTYPE           =   2
            TX              =   "«‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
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
            BCOL            =   14871017
            BCOLO           =   14871017
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmEmpSalary5.frx":9262
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo Dcdep 
            Height          =   315
            Left            =   14040
            TabIndex        =   79
            Top             =   1080
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCGroupID 
            Height          =   315
            Left            =   14040
            TabIndex        =   81
            Top             =   720
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcempcontract 
            Height          =   315
            Left            =   9360
            TabIndex        =   83
            Top             =   840
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcBranch1 
            Height          =   315
            Left            =   14040
            TabIndex        =   87
            Top             =   360
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   11745
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1215
            Width           =   885
         End
         Begin MSComDlg.CommonDialog cd 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo DcbTeam 
            Height          =   315
            Left            =   9360
            TabIndex        =   104
            Top             =   120
            Width           =   3270
            _ExtentX        =   5768
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
         Begin MSDataListLib.DataCombo DcbHemiaSalary 
            Height          =   315
            Left            =   5520
            TabIndex        =   109
            Top             =   120
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSDataListLib.DataCombo DcbDepartment2 
            Height          =   315
            Left            =   14040
            TabIndex        =   110
            Top             =   1440
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton cmdReCreateJL 
            Height          =   300
            Left            =   3480
            TabIndex        =   115
            Top             =   870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«⁄«œ… «‰‘«¡ «·ﬁÌœ"
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
            ButtonImage     =   "FrmEmpSalary5.frx":927E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ”„"
            Height          =   240
            Index           =   21
            Left            =   16920
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÊœ Õ„«Ì… «·«ÃÊ—"
            Height          =   240
            Index           =   20
            Left            =   8160
            TabIndex        =   108
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”œ«œ"
            Height          =   360
            Index           =   19
            Left            =   12600
            TabIndex        =   107
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—Ìﬁ"
            Height          =   240
            Index           =   18
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
            Height          =   240
            Index           =   17
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   240
            Index           =   16
            Left            =   16920
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   375
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Êﬁ⁄"
            Height          =   240
            Index           =   15
            Left            =   16920
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "«Œ — «·„Õœœ«  À„ «÷€ÿ Enter "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   13980
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   75
            Width           =   3765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—…"
            Height          =   240
            Index           =   4
            Left            =   16920
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1095
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ ‰„Ê–Ã"
            Height          =   225
            Index           =   6
            Left            =   3105
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„‘—Ê⁄"
            Height          =   240
            Index           =   5
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   495
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ÊŸ› „Õœœ"
            DataField       =   "Õœœ"
            Height          =   330
            Index           =   3
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1230
            Width           =   1020
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   16
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
      ButtonImage     =   "FrmEmpSalary5.frx":9618
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalary5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cProgress As ClsProgress
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim showinMosirVac(40) As Boolean
Dim culc30orRminder(40) As Integer
Dim FixedOrChanged(40) As Integer
Dim MofrdAbcen(40) As Boolean
Dim AddOrDiscount(40) As Integer
Dim ViewComp(40) As Boolean
Dim showMofradAll(40) As Boolean
Dim Account_code(40) As String
Dim Account_code1(40) As String
Dim empDes As String
Dim ZmamAccount(40) As String

Dim AdvPaymentdAccount(40) As String
Dim rsEmpID As ADODB.Recordset
Dim componentname(40) As String
Dim firstrun As Boolean
Dim rsBranch As New ADODB.Recordset
Dim RsDepartment As New ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Private Sub Coloring()
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 80) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 80) = vbWhite
            End If

        Next i

    End With

    With GRID1

        For i = .FixedRows To .rows - 2
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 62) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 62) = vbWhite
            End If

        Next i

    End With
 
     With Grid2

        For i = .FixedRows To .rows - 2
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 62) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 62) = vbWhite
            End If
  '              If val(GRID2.TextMatrix(i, GRID2.ColIndex("ID"))) = 0 Then
  '              GRID2.RemoveItem i
  '                GRID2.TextMatrix(i, Grid1.ColIndex("Ser")) = i
  '              End If
        Next i

    End With
    
End Sub
Sub RenameCompoPrint()
    
     If Dir(App.path & "\SalaryPrint.txt") = "" Then
   '  MsgBox "„·›  ﬁ«—Ì— «·„”Ì— €Ì— „ÊÃÊœ ", vbCritical
     Exit Sub
    End If
    Open App.path & "\SalaryPrint.txt" For Input As #1
    Combo1.Clear
Dim a As String
Dim VarSet() As String
    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
           VarSet = Split(a, "*", , vbTextCompare)
           

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                Combo1.AddItem (VarSet(0))
                            
            End If
          Else
                    VarSet = Split(a, "*", , vbTextCompare)

            If Not IsNull(VarSet(1)) Or VarSet(1) <> Empty Then
            
                Combo1.AddItem (VarSet(1))
                            
            End If
            
          End If
        End If
    
    Loop

    Close #1
End Sub
Private Sub GetMySetting()

    Dim StrSetting As String
    Dim StrShowSet As String
    Dim frmname As String

    Dim VarCols As Variant
    Dim VarColSet As Variant
    Dim i As Integer
    On Error Resume Next
    frmname = Me.Name
    StrSetting = GetSetting(SystemOptions.SysRegsAppPath, "Interface SettingEmpSalary" & "\" & user_id, frmname, "")

    If StrSetting = "" Then
        Exit Sub
    End If

    VarCols = Split(StrSetting, ";", , vbTextCompare)

    If UBound(VarCols) > 0 Then

        For i = 0 To UBound(VarCols)
            VarColSet = Empty
            VarColSet = Split(CStr(VarCols(i)), "-", , vbTextCompare)

            With Grid
                .ColPosition(.ColIndex(CStr(VarColSet(0)))) = CLng(VarColSet(1))
            End With

            With GRID1
                .ColPosition(.ColIndex(CStr(VarColSet(0)))) = CLng(VarColSet(1))
            End With
        
        Next i

    End If

    StrShowSet = GetSetting(SystemOptions.SysRegsAppPath, "Cols SettingEmpSalary" & "\" & user_id, frmname, "")

    If StrShowSet = "" Then
        Exit Sub
    End If

    VarCols = Split(StrShowSet, ";", , vbTextCompare)

    If UBound(VarCols) > 0 Then

        For i = 0 To UBound(VarCols)
            VarColSet = Empty
            VarColSet = Split(CStr(VarCols(i)), "-", , vbTextCompare)

            With Grid
                .ColHidden(.ColIndex(CStr(VarColSet(0)))) = CBool(VarColSet(1))
            End With
        
            With GRID1
                .ColHidden(.ColIndex(CStr(VarColSet(0)))) = CBool(VarColSet(1))
            End With
        
        Next i

    End If

    StrSetting = ""
 
End Sub
 
Sub SaveMySetting()
    Dim i As Integer
    Dim StrTemp As String
    Dim StrShow As String
    Dim frmname As String
    frmname = Me.Name
 
    For i = 0 To Grid.Cols - 1
        StrTemp = StrTemp & Grid.ColKey(i) & "-" & i & ";"
        StrShow = StrShow & Grid.ColKey(i) & "-" & Grid.ColHidden(i) & ";"
    Next i

    StrTemp = Trim(StrTemp)
    StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
    StrShow = Trim(StrShow)
    StrShow = mId(StrShow, 1, Len(StrShow) - 1)
    SaveSetting SystemOptions.SysRegsAppPath, "Interface SettingEmpSalary" & "\" & user_id, frmname, StrTemp
    SaveSetting SystemOptions.SysRegsAppPath, "Cols SettingEmpSalary" & "\" & user_id, frmname, StrShow

    '-----------------------------------------

End Sub

Private Function save_cost_center(cost_center_id As String, _
                                  opr_type As String, _
                                  record_date As Date, _
                                  value As Double, _
                                  kedno As String, _
                                  account_no As String, _
                                  account_name As String, _
                                  line_no As Double)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = get_EMPLOYEE_COST_CENTER_NAME(cost_center_id, "ACCOUNT_NAME")
    rs("value").value = value
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = kedno
    rs("kedno").value = kedno
        
    rs("opr_type").value = opr_type
    rs("account_name").value = account_name
    rs("account_no").value = account_no
    rs("line_no").value = line_no
    rs("record_date").value = record_date
    rs.update
    rs.Close

End Function

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 4050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()


    FrmShowCol1.show
End Sub
Sub UpdateLockSalary(year As String, Month As String)
Dim sql As String

 sql = " update notes set LockSalary =1  where NoteSerial=" & TxtNoteSerial.text & " and  salary=" & val(year) & Month & " "
 Cn.Execute sql
 
End Sub
   
Function GetNotesSerials(year As String, Month As String, NoteType As Integer) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    If val(dcBranch1.BoundText) <> 0 Then
    Current_branch = val(dcBranch1.BoundText)
    End If
    
    sql = "Select NoteSerial from notes where salary=" & val(year) & Month & " and  NoteType=" & NoteType & " and branch_no=" & Current_branch
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        GetNotesSerials = ""
    Else
        GetNotesSerials = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    End If
 
End Function
Function check_Lock_Salary(year As String, Month As String, Optional branch_no As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    branch_no = Current_branch
    sql = "Select * from notes where LockSalary=1 and salary=" & year & Month
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_Lock_Salary = False
    Else
        check_Lock_Salary = True
    End If
 
End Function

Function check_previous_dev(year As String, Month As String, Optional branch_no As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    branch_no = Current_branch
    sql = "Select * from notes where salary=" & year & Month & "and branch_no=" & branch_no
 
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

Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\employee_account_error.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡ «·„ÊŸ›Ì‰ «·–Ì‰ ·œÌÂ„ „‘«ﬂ·  "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Function check_employee_accounts() As Boolean
    Dim Employee_account As String
    Dim error_string As String
    error_string = ""
    check_employee_accounts = True
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
                   If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                   error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡    ÕœÌœ «·›—⁄ «· «»⁄ ·Â"
        
                check_employee_accounts = False
                   End If
                   
                   
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")

            If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» –„ …"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» –„ … ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» «·«ÃÊ— «·„” Õﬁ…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«»   «·„œ›Ê⁄«  «·„ﬁœ„…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«»    «·„œ›Ê⁄«  «·„ﬁœ„… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            '     If Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) = 0 Then
            '     error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & " ·„ Ì „  ÕœÌœ —« » «”«”Ì ·Â  " & vbCrLf
            '
            '    check_employee_accounts = False
            '
            '     End If
            If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
        Next i

    End With

    Dim X As Integer
    Dim StrLogFileName As String

    If error_string <> "" Then
        X = MsgBox("Â·  —Ìœ › Õ «·„·› ··„—«Ã⁄Â", vbCritical + vbYesNo, "ÌÊÃœ Œÿ√ ›Ì Õ”«»«  «·„ÊŸ›Ì‰  Ì„ﬂ‰ „—«Ã⁄ … ›Ì „·› «·«Œÿ«¡")

        If X = vbYes Then
            StrLogFileName = App.path & "\employee_account_error.txt"
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        End If
    End If

End Function

Function Create_dev() As Boolean
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
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        
    If check_employee_accounts = False Then
        Exit Function
    End If

    If check_previous_dev(CboYear.text, CmbMonth.ListIndex + 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Else
            MsgBox "JV Alraedy Created To this Month", vbCritical
        End If

        Create_dev = False
        Exit Function
          
    End If
        
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
        
'    If Account_Code_dynamic = "NO branch" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        Else
'            MsgBox "No Branch Created", vbCritical
'        End If
'
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            Else
'                MsgBox "The Salary Account in this Branch is not specific", vbCritical
'            End If
'
'            GoTo ErrTrap
'
'        End If
'    End If

'    SalaryAccount = Account_Code_dynamic
'    Account_Code_dynamic = get_account_code_branch(53, my_branch)
'    DiscountAccount = Account_Code_dynamic

'    If Account_Code_dynamic = "NO branch" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        Else
'            MsgBox "No Branch Created", vbCritical
'        End If
'
    '    GoTo ErrTrap
'    Else

    '    If Account_Code_dynamic = "NO account" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '            MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·Œ’„      ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
    '        Else
    '            MsgBox "The Salary Account in this Branch is not specific", vbCritical
    '        End If
'
'            GoTo ErrTrap
'
'        End If
'    End If
        
'    Account_Code_dynamic = get_account_code_branch(54, my_branch)
'    BonusAccount = Account_Code_dynamic

'    If Account_Code_dynamic = "NO branch" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        Else
'            MsgBox "No Branch Created", vbCritical
'        End If

'        GoTo ErrTrap
    'Else

'        If Account_Code_dynamic = "NO account" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„ﬂ«›√…   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            Else
'                MsgBox "The Salary Account in this Branch is not specific", vbCritical
'            End If
'
   '         GoTo ErrTrap
         
'        End If
'    End If
        
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
    Else
        Msg = " Salary Allocation JL Month: " & CmbMonth.text & "   Year: " & CboYear.text
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress
    DoEvents
 
    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"


    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
 
    'notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    my_branch = branch_id

    If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
            MsgBox " Can not start a new JL, you exceed the limit ": GoTo ErrTrap
        End If

    Else
                       
        If Notes_coding(val(my_branch), DTP_Date.value) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                MsgBox "   Can not Create a new JV , you Select Manual Numbering in JL Voucher Coding ": GoTo ErrTrap
            End If

        Else
            notes_serial = Notes_coding(val(my_branch), DTP_Date.value)
        End If
    End If
               
    If SystemOptions.UserInterface = ArabicInterface Then
        ALLButton2.Caption = "Ì—ÃÏ «·«‰ Ÿ«— Õ Ï «·«‰ Â«¡"
    Else
        ALLButton2.Caption = "Wait this process may take several minutre"
    End If

    rs.AddNew
    rs("branch_no").value = 1
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = net_value ' Null
    '   rs("note_value_by_characters").value = WriteNo(Format(net_value, "0.00"), 0, True, ".")
    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 66
    rs("NoteDate").value = DTP_Date.value
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("sanad_year").value = year(DTP_Date.value)
    rs("sanad_month").value = Month(DTP_Date.value)
     rs("branch_no").value = Current_branch
    
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1

    '     If .TextMatrix(i, .ColIndex("project")) = "" Or Val(.TextMatrix(i, .ColIndex("project"))) = 0 Then
    '
    '              If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, _
    '                 Msg, Val(notes_id), , , , Date, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , Val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
    '             GoTo ErrTrap
    '             End If
    '      Else
    '               Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
    '               If Account_Code_dynamic1 <> "" Then
    '                      If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, _
    '                        Msg, Val(notes_id), , , , Date, user_id, , , , , , , , , setfoxy_Line, , Val(.TextMatrix(i, .ColIndex("project"))), , , , , , , Val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
    '                     GoTo ErrTrap
    '                     End If
    '              End If
    '      End If
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
    Dim BranchID As Integer
    BranchID = 1

    With Grid

        For j = 1 To 40
            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„
                If ZmamAccount(j) <> True Then
                    If val(.TextMatrix(.rows - 1, .ColIndex(ColumnName))) > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), val(.TextMatrix(.rows - 1, .ColIndex(ColumnName))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                End If
                             
            End If
    
        Next j
       
        '«·„ﬂ«›√ 
        If val(.TextMatrix(.rows - 1, .ColIndex("Mokafea"))) > 0 Then
        '    If ModAccounts.AddNewDev(LngDevID, line_no, BonusAccount, val(.TextMatrix(.Rows - 1, .ColIndex("Mokafea"))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
        '        GoTo ErrTrap
        '    End If
'
'            line_no = line_no + 1
        End If
                                    
        '«·Œ’Ê„« 
        If val(.TextMatrix(.rows - 1, .ColIndex("TotalDiscount"))) > 0 Then
'            If ModAccounts.AddNewDev(LngDevID, line_no, DiscountAccount, val(.TextMatrix(.Rows - 1, .ColIndex("TotalDiscount"))), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 1
        End If
          
        '      «·Œ’Ê„« 
        For j = 1 To 40
            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
                If ZmamAccount(j) <> True Then
                    If val(.TextMatrix(.rows - 1, .ColIndex(ColumnName))) > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), val(.TextMatrix(.rows - 1, .ColIndex(ColumnName))), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                End If
            End If
    
        Next j

        For i = .FixedRows To .rows - 2
    
            If .TextMatrix(i, .ColIndex("EmpTotalNet")) > 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
            For j = 1 To 40 '  „« ÌŒ’ –„… «·„ÊŸ›
                ColumnName = "Comp" & j

                If ViewComp(j) = True And ZmamAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
 
 
 '*********************************„œ›Ê⁄«  „ﬁœ„…*************************************
 

            
 '**********************************************************************
        Next i

    End With

    Create_dev = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
    Create_dev = False
    'If SystemOptions.UserInterface = ArabicInterface Then
    'MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    'Else
    'MsgBox "Error During Saving", vbExclamation
    'End If
End Function

Function CheckPayRollHaveBranches() As Double
    Dim i As Integer
    Dim SUM As Integer
    SUM = 0

    With Grid

        For i = .FixedRows To .rows - 2
            SUM = SUM + val(.TextMatrix(i, .ColIndex("BranchId")))
        Next i

        CheckPayRollHaveBranches = (.rows - 2) / SUM
    End With

End Function

Function GetComponentValuePerBranch2(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With GRID1

        For i = .FixedRows To .rows - 2
    
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If

        Next i

    End With

    GetComponentValuePerBranch2 = SUM
End Function
Function GetComponentValuePerEmpID(EmpID As Double, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer
    With Grid
        For i = .FixedRows To .rows - 2
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("Emp_ID"))) = EmpID Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If
        Next i

    End With

    GetComponentValuePerEmpID = SUM
End Function

Function GetComponentValuePerBranch(BramchId As Integer, componentname As String, Optional DeparmentID As Double = 0) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
         If SystemOptions.SalaryJLByManagement = True And DeparmentID <> 0 Then
                     If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId And val(.TextMatrix(i, .ColIndex("dep"))) = DeparmentID Then
                         If val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                            SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
                        End If
            End If
         Else
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                If val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                    SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
                End If
            End If
         End If

        Next i

    End With

    GetComponentValuePerBranch = SUM
End Function
Function Create_dev3() As Boolean
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        Dim mofradname  As String
        Dim DepartmentID As Double
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    Dim rsDummy As ADODB.Recordset
    Dim s As String
Dim mMonth As Integer
mMonth = val(CmbMonth.ListIndex) + 1
    
    If check_employee_accounts = False Then
        Exit Function
    End If

    If check_previous_dev(CboYear.text, CmbMonth.ListIndex + 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Else
            MsgBox "JV Alraedy Created To this Month", vbCritical
        End If

        Create_dev3 = False
        Exit Function
          
    End If
        
 
 
        
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
    Else
        Msg = " Salary Allocation JL Month: " & CmbMonth.text & "   Year: " & CboYear.text
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress
    DoEvents
 
    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
 
    'notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    my_branch = Current_branch

    If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
            MsgBox " Can not start a new JL, you exceed the limit ": GoTo ErrTrap
        End If

    Else
                       
        If Notes_coding(val(my_branch), DTP_Date.value) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                MsgBox "   Can not Create a new JV , you Select Manual Numbering in JV Voucher Coding ": GoTo ErrTrap
            End If

        Else
            notes_serial = Notes_coding(val(my_branch), DTP_Date.value)
        End If
    End If
               
    If SystemOptions.UserInterface = ArabicInterface Then
        ALLButton2.Caption = "Ì—ÃÏ «·«‰ Ÿ«— Õ Ï «·«‰ Â«¡"
    Else
        ALLButton2.Caption = "Wait this process may take several minutre"
    End If

    rs.AddNew
    rs("branch_no").value = Current_branch
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = net_value ' Null

    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 66
    rs("NoteDate").value = DTP_Date.value
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("sanad_year").value = year(DTP_Date.value)
    rs("sanad_month").value = Month(DTP_Date.value)
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
    Dim BranchID As Integer
    Dim DepartmentID1 As Double
    Dim DepartmentName  As String
    Dim CValue As Double
    Dim Branch As Integer
    Dim Dept As Integer
    Dim ProjectID As Integer
    
    BranchID = 1

    With Grid

        For j = 1 To 40

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If
        

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                   
                    For Branch = 1 To rsBranch.RecordCount
                    BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                     If SystemOptions.SalaryJLByManagement = True Then
                             If RsDepartment.RecordCount > 0 Then
                             RsDepartment.MoveFirst
                             End If
                      For Dept = 1 To RsDepartment.RecordCount
                         DepartmentID1 = IIf(IsNull(RsDepartment("DeparmentID").value), 0, (RsDepartment("DeparmentID").value))
                        If SystemOptions.UserInterface = ArabicInterface Then
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentName").value), "", (RsDepartment("DepartmentName").value))
                         Else
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentNameE").value), "", (RsDepartment("DepartmentNameE").value))
                         End If
                          CValue = GetComponentValuePerBranch(BranchID, ColumnName, DepartmentID1)
                      If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)) & "   " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID1) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If
                        RsDepartment.MoveNext
                        Next Dept
                         Else
                         CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                              
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If
                        End If
                        rsBranch.MoveNext
                    Next Branch
                             
                End If
                             
            End If
    
        Next j
       
        '«·„ﬂ«›√ 
        If rsBranch.RecordCount > 0 Then
            rsBranch.MoveFirst
        End If
            
        For Branch = 1 To rsBranch.RecordCount
                                 
            BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                         
            CValue = GetComponentValuePerBranch(BranchID, "Mokafea")
                               
            If CValue > 0 Then
            
                If CValue > 0 Then
          '          If ModAccounts.AddNewDev(LngDevID, line_no, BonusAccount, CValue, 0, Msg & "   „ﬂ«›√ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
          '              GoTo ErrTrap
          '          End If

                    line_no = line_no + 1
                End If
                                    
            End If

            rsBranch.MoveNext
        Next Branch
                                    
        '«·Œ’Ê„« 
        If rsBranch.RecordCount > 0 Then
            rsBranch.MoveFirst
        End If
            
        For Branch = 1 To rsBranch.RecordCount
                                 
            BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
            CValue = GetComponentValuePerBranch(BranchID, "TotalDiscount")
                               
            If CValue > 0 Then
    
                If CValue > 0 Then
                
               ' If SystemOptions.ProjectEmployeeGV = False Then
          '          If ModAccounts.AddNewDev(LngDevID, line_no, DiscountAccount, CValue, 1, Msg & "  Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
          '              GoTo ErrTrap
          '          End If
               '  End If
                 
                    line_no = line_no + 1
                End If
                                    
            End If

            rsBranch.MoveNext
        Next Branch
                                    
        For j = 1 To 40 ' Œ’Ê„« 

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                                                   
                    For Branch = 1 To rsBranch.RecordCount
                                 
                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               
                        If CValue > 0 Then
                   '      SystemOptions.ProjectEmployeeGV = True
 If SystemOptions.ProjectDiscountPolicy = 1 Then
 'Dim CurrentAccount As String
' CurrentAccount = Account_Code(j)
                           '  If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                     '   If Account_code1(j) <> "" Then
                                     '   CurrentAccount = Account_code1(j)
                                     '   End If
                            
                             
                           '  End If
                             
          '                  If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(J), CValue, 1, Msg & "   Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
          '                      GoTo ErrTrap
          '                  End If
                            
          '                  Else
          '
          '                           If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code(J), CValue, 1, Msg & "   Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
          '                      GoTo ErrTrap
          '                  End If
                            
                            
                            
End If
                            line_no = line_no + 1
                        End If
                                    
                        rsBranch.MoveNext
                    Next Branch
                                    
                End If
            End If
    
        Next j

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex("total1"))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                
                If val(.TextMatrix(i, .ColIndex("VacDay"))) = daysInMonth(DTP_Date.value) Then
                    If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("total1")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                        GoTo ErrTrap
                    End If
                Else
                    If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("total1")) - val(.TextMatrix(i, .ColIndex("TotalVacValue"))), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                        GoTo ErrTrap
                    End If
                End If

'
                
                
                
                

                line_no = line_no + 1
            End If
     
     
     
     
     
            If val(.TextMatrix(i, .ColIndex("TotalVacValue"))) > 0 Then '«·«Ã«“«  «·„” Õﬁ… œ«∆‰
                
                
                s = "Select Account_Code from tblVacancy where IsNull(Account_Code,'') <> '' "
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                
                If Not rsDummy.EOF Then
                    StrAccountCode = Trim(rsDummy!Account_code & "")
                End If
                 'If val(.TextMatrix(i, .ColIndex("VacDay"))) = daysInMonth(DTP_Date.value) Then
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalVacValue")), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If
                
                
                
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code2") '«·«Ã«“«  «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                
                line_no = line_no + 1
                
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalVacValue")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If
                
            '    End If
                
                
                

                line_no = line_no + 1
            End If
             '    If .TextMatrix(i, .ColIndex("EmpTotalNet")) < 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                
             '   Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
             '   StrAccountCode = Employee_account
        '
        '        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(.TextMatrix(i, .ColIndex("EmpTotalNet"))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
        '            GoTo ErrTrap
        '        End If
'
'                line_no = line_no + 1

'            End If
            
            
            
                      
        '      «·Œ’Ê„« 
        For j = 1 To 40
            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = -1 And (ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True) Then
                If ZmamAccount(j) <> True Then
                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                    
                               Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        '
                       
                     If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex(ColumnName))), 0, Msg & "  " & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                            GoTo ErrTrap
                        End If
                            line_no = line_no + 1
                        
                        If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), val(.TextMatrix(i, .ColIndex(ColumnName))), 1, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                            GoTo ErrTrap
                        End If
                        
                        
                        

                        line_no = line_no + 1
                    End If
                End If
            End If
    
        Next j



            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And ZmamAccount(j) = True Then
                     
                                 
              
        '
                       
             
                            
                            
                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                               Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                                StrAccountCode = Employee_account
                                                 If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                            GoTo ErrTrap
                                        End If
        
                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                                                        
                                  
                                line_no = line_no + 1
                                
                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                                            GoTo ErrTrap
                                        End If
        
                                line_no = line_no + 1
                            End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then  '«·”·› œ«∆‰
            
            
            
                                      Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 0, Msg & " ”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                                            GoTo ErrTrap
                                        End If
                                        
                         line_no = line_no + 1
                                        
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
 
            ' ÿ—› «·«Ã«“«  «·„—÷Ì…
              If .TextMatrix(i, .ColIndex("VoCation3")) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then  '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")  '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                
                
                
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("VoCation3")), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                
                StrAccountCode = get_account_code_branch(204, my_branch)
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("VoCation3")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
            End If
     
 
 
 
 
 
 
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                          Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                                StrAccountCode = Employee_account
                                 
                     If AddOrDiscount(j) = 0 Then
                      
                                                           
                                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                                                                GoTo ErrTrap
                                                            End If
                        
                                                line_no = line_no + 1
                                            End If
                 
                 Else
                 
                 
                 
                 
                                           If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then
                                           
                                           
                                                       
                                      Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
                                                 If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & " ”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                                            GoTo ErrTrap
                                        End If
                                        
                                        line_no = line_no + 1
                                        
                                             Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                                StrAccountCode = Employee_account
                                     
                                     
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep")))) = False Then
                                                                GoTo ErrTrap
                                                            End If
                        
                                                line_no = line_no + 1
                                            End If
                 
                 
                 
                 End If
                 
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With
'**************************************************************************

'«· √„Ì‰« 


    rs.Close
    
       Dim sql As String
       sql = " "
' «„Ì‰«  ›Ì «·„”Ì—
Dim Nationality  As String
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim emp_Name As String
    GetInsuranceAccount mofradAccount, mofradAccount1
 With Grid
                         For i = .FixedRows To .rows - 2
                
                            If val(.TextMatrix(i, .ColIndex("ToalInsurance"))) > 0 And val(.TextMatrix(i, .ColIndex("VacDay"))) <> daysInMonth(DTP_Date.value) Then  '
                            Emp_id = val(.TextMatrix(i, .ColIndex("Emp_ID")))
                   emp_Name = (.TextMatrix(i, .ColIndex("emp_Name")))
                                Nationality = (.TextMatrix(i, .ColIndex("Nationality")))
                                    If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
                                                mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                                     End If
                     mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")  '«·«ÃÊ— «·„” Õﬁ…
                                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) <> 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰

                                If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 0, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›    " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                '
                                line_no = line_no + 1
                                End If
                                
                                      If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 1, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›   " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                
                                line_no = line_no + 1
                                
                                
                            End If
                 
                Next i

End With

Dim EmpList As String

EmpList = ""
For i = 1 To Grid.rows - 1
    If Trim$(Grid.TextMatrix(i, Grid.ColIndex("Emp_id"))) <> "" Then
        If EmpList <> "" Then EmpList = EmpList & ","
        EmpList = EmpList & Grid.TextMatrix(i, Grid.ColIndex("Emp_id"))
    End If
Next i

If EmpList <> "" Then
    EmpList = "(" & EmpList & ")"
Else
    EmpList = "()"
End If


'  √„Ì‰«  «Ã«“« 
sql = "  SELECT      dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, " & CHR(13)
sql = sql & "                       dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID,SUM(dbo.EmpSalaryComponent.[Value]) AS value, dbo.mofrad.Account_Code, dbo.mofrad.Account_code1" & CHR(13)
sql = sql & "  FROM         dbo.mofrad INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN" & CHR(13)
sql = sql & "                         dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID" & CHR(13)
sql = sql & "    WHERE     (dbo.mofrad.Insurances = 1) AND (dbo.EmpSalaryComponent.emp_ID IN " & EmpList & ") and (dbo.EmpSalaryComponent.emp_ID IN" & CHR(13)
sql = sql & "                             (SELECT     Emp_ID" & CHR(13)
sql = sql & "                                From TblEmployee" & CHR(13)
sql = sql & "                                WHERE     dbo.TblEmployee.jopstatusid IN" & CHR(13)
sql = sql & "                                                          (SELECT     id" & CHR(13)
sql = sql & "                                                             From dbo.jopstatus" & CHR(13)
sql = sql & "                                                             WHERE     Insurances = 1 and id<>1  ) AND dbo.TblEmployee.BignDateWork <" & SQLDate(DTP_Date.value, True) & ")) AND" & CHR(13)
sql = sql & "                         (year(dbo.EmpSalaryComponent.EntIncresDataM)<year( " & SQLDate(DTP_Date.value, True) & ") OR" & CHR(13)
                      sql = sql & "   dbo.EmpSalaryComponent.EntIncresDataM IS NULL) AND (dbo.mofrad.Insurances = 1)" & CHR(13)
sql = sql & "   GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID," & CHR(13)
sql = sql & "                         dbo.MOFRAD.Account_Code , dbo.MOFRAD.Account_code1, dbo.TblEmployee.fullcode,dbo.TblEmployee.Nationality" & CHR(13)
sql = sql & "    ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)





' = 0
'ResidentVal1 = 0
Dim CitizenVal1 As Double
Dim ResidentVal1 As Double
Dim Balance As Double
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     'mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     'mofradAccount1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
     
        GetInsuranceAccount mofradAccount, mofradAccount1, CitizenVal1, ResidentVal1
     Nationality = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
       Emp_id = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
      
     If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
     
        mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                 
  Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * CitizenVal1 / 100
  
  Else
 Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * ResidentVal1 / 100
     End If
     
     
     
      
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     mofradname = IIf(IsNull(rs("AccountName").value), "", rs("AccountName").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
   
      emp_Name = IIf(IsNull(rs("emp_Name").value), "", rs("emp_Name").value)
                             If mofradAccount <> "" And mofradAccount1 <> "" And Balance > 0 Then
                                   
                                  
                                   If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, Balance, 0, Msg & mofradname & "   √Ì‰« -€Ì— „ÊÃÊœ… »«·„”Ì— " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                        
                                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  √Ì‰«  -€Ì— „ÊÃÊœ… »«·„”Ì—" & "  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                             
                                             
                                             
                             End If
     rs.MoveNext
     Next i
    End If

    
    '*************************************************************************************

  If SystemOptions.ProjectEmployeeGV = True Then
rs.Close
   ' Dim sql As String
    
   ' Dim Balance As Double
'Dim mofradAccount As String
'Dim mofradAccount1 As String

Dim Salary_account As String
 Dim Project_name As String
 'Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(DTP_Date.value, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(DTP_Date.value, True) & "))"

sql = sql & " AND (dbo.TblChangedComponentRegisterDetails.emp_ID IN " & EmpList & ")"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then   '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    '
                     '            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                     '       GoTo ErrTrap
                     '   End If
        
        
    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
'                                line_no = line_no + 1
'                                                             If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
'                            GoTo ErrTrap
'                        End If
'
'
'                        line_no = line_no + 1
'
'                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
'                            GoTo ErrTrap
'                        End If
        
        
            line_no = line_no + 1
        
        
         
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
    
  
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 'Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(DTP_Date, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & " ))"
sql = sql & " AND (dbo.TblChangedComponentRegisterDetails.emp_ID IN " & EmpList & ")"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then   '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
   If empDes <> "" Then
    sql = " SELECT    distinct TOP 100 PERCENT mofrdat.mofrad_code, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
sql = sql & "                      dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
sql = sql & "                      dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, "
sql = sql & "                       dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
sql = sql & "                      dbo.TblEmployee.ContractID ,ProJectMofrdSalar.ProjID ProjectID,mofrad.AddOrDiscount,mofrad.Account_code, dbo.TblEmployee.GroupID, dbo.Projects.Salary_account , dbo.ProJectMofrdSalar.TypeSalary,"

sql = sql & "                        NoDay = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & " THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END,"
sql = sql & "                             Total = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END * Valuee"
sql = sql & " FROM         dbo.ProJectMofrdSalar LEFT OUTER JOIN"
sql = sql & "                      dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID"
 
sql = sql & " LEFT OUTER JOIN dbo.mofrad "
      
sql = sql & "            ON  dbo.mofrad.id = dbo.mofrdat.mofrad_type"
'sql = sql & "  WHERE     (" & SQLDate(DTP_Date.value, True) & "BETWEEN dbo.ProJectMofrdSalar.fromDate AND dbo.ProJectMofrdSalar.toDate) " & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

'sql = sql & "  Where (dbo.ProJectMofrdSalar.YearID = " & val(CboYear.ListIndex) & ") And (dbo.ProJectMofrdSalar.MonthID = " & val(CmbMonth.ListIndex) & ") and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  Where ( year(dbo.ProJectMofrdSalar.fromDate ) = " & val(CboYear.text) & ") "
'And ( month(dbo.ProJectMofrdSalar.fromDate) = " & val(CmbMonth.ListIndex) + 1 & ")

sql = sql & "  AND ("
sql = sql & "                 ("
sql = sql & "                     Month (dbo.ProJectMofrdSalar.FromDate) <= " & mMonth & " "
sql = sql & "                     AND ("
sql = sql & "                             ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                             OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                         )"
sql = sql & "                 )"
sql = sql & "                 OR MONTH(dbo.ProJectMofrdSalar.fromDate) = " & mMonth & " "
sql = sql & "             )"

sql = sql & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  ORDER BY dbo.ProJectMofrdSalar.ProjID ,dbo.ProJectMofrdSalar.EmpID"
    
    
    
If empDes <> "" Then





sql = " DECLARE @SelectedMonth INT "
sql = sql & " SET @SelectedMonth = " & mMonth & "; "
sql = sql & " WITH DateRanges AS ( "
sql = sql & " SELECT "
sql = sql & " ProJectMofrdSalar.id, mofrdat.mofrad_code, dbo.ProJectMofrdSalar.EmpId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
sql = sql & " dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID, "
sql = sql & " dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, mofrad.AddOrDiscount, mofrad.Account_code, "
sql = sql & " dbo.ProJectMofrdSalar.Valuee, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, "
sql = sql & " dbo.TblEmployee.BranchId, dbo.TblEmployee.ContractID, dbo.ProJectMofrdSalar.ProjID AS ProjectID, "
sql = sql & " dbo.TblEmployee.GroupID, dbo.projects.Salary_account, dbo.ProJectMofrdSalar.TypeSalary, "
sql = sql & " ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate "
sql = sql & " FROM dbo.ProJectMofrdSalar "
sql = sql & " LEFT JOIN dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code "
sql = sql & " LEFT JOIN dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id "
sql = sql & " LEFT JOIN dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpId = dbo.TblEmployee.Emp_ID "
sql = sql & " LEFT JOIN dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.ID "
sql = sql & " WHERE (dbo.ProJectMofrdSalar.EmpID IN (" & empDes & ")) "
sql = sql & " AND (YEAR(dbo.ProJectMofrdSalar.FromDate) = " & val(CboYear.text) & ") "
sql = sql & " AND ((MONTH(dbo.ProJectMofrdSalar.FromDate) <= @SelectedMonth "
sql = sql & " AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) >= @SelectedMonth)) "
sql = sql & " OR (MONTH(dbo.ProJectMofrdSalar.FromDate) < @SelectedMonth "
sql = sql & " AND MONTH(dbo.ProJectMofrdSalar.ToDate) = @SelectedMonth)) "
sql = sql & " ) "

sql = sql & " SELECT "
sql = sql & " ProJectMofrdSalar.id, mofrdat.mofrad_code, dbo.ProJectMofrdSalar.EmpId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
sql = sql & " dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrDID, "
sql = sql & " dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, mofrad.AddOrDiscount, mofrad.Account_code, "
sql = sql & " dbo.ProJectMofrdSalar.Valuee, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, "
sql = sql & " dbo.TblEmployee.BranchId, dbo.TblEmployee.ContractID, dbo.ProJectMofrdSalar.ProjID AS ProjectID, "
sql = sql & " dbo.TblEmployee.GroupID, dbo.projects.Salary_account, dbo.ProJectMofrdSalar.TypeSalary, "
sql = sql & " ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate, "
sql = sql & " CASE "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) > @SelectedMonth) "
sql = sql & " THEN (dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth) - DAY(ProJectMofrdSalar.FromDate) + 1) "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth "
sql = sql & " THEN (DATEDIFF(DAY, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate) + 1) "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) > @SelectedMonth) "
sql = sql & " THEN dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth) "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth "
sql = sql & " THEN DAY(ProJectMofrdSalar.ToDate) "
sql = sql & " ELSE 0 "
sql = sql & " END AS NoDay, "
sql = sql & " CASE "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) > @SelectedMonth) "
sql = sql & " THEN (dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth) - DAY(ProJectMofrdSalar.FromDate) + 1) * ProJectMofrdSalar.Valuee "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth "
sql = sql & " THEN (DATEDIFF(DAY, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate) + 1) * ProJectMofrdSalar.Valuee "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) > @SelectedMonth) "
sql = sql & " THEN dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth) * ProJectMofrdSalar.Valuee "
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth "
sql = sql & " THEN DAY(ProJectMofrdSalar.ToDate) * ProJectMofrdSalar.Valuee "
sql = sql & " ELSE 0 "
sql = sql & " END AS Total "
sql = sql & " FROM dbo.ProJectMofrdSalar "
sql = sql & " LEFT JOIN dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code "
sql = sql & " LEFT JOIN dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id "
sql = sql & " LEFT JOIN dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpId = dbo.TblEmployee.Emp_ID "
sql = sql & " LEFT JOIN dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.ID "
sql = sql & " JOIN DateRanges ON dbo.ProJectMofrdSalar.FromDate <= DateRanges.ToDate "
sql = sql & " AND dbo.ProJectMofrdSalar.ToDate >= DateRanges.FromDate "
sql = sql & " AND DateRanges.EmpID = ProJectMofrdSalar.EmpID "

sql = sql & " UNION "

sql = sql & " SELECT "
sql = sql & " 0, mofrad.ID, p.Emp_id, TblEmployee.Emp_Name, TblEmployee.Fullcode, TblEmployee.Emp_Namee, "
sql = sql & " p.projectid, pr.Project_name, pr.Project_nameE, PP.ComponentID, "
sql = sql & " mofrad.name, mofrad.namee, mofrad.AddOrDiscount, mofrad.Account_code, "
sql = sql & " p.value, TblEmployee.SalaryType, TblEmployee.DepartmentID, TblEmployee.BranchId, TblEmployee.ContractID, "
sql = sql & " p.projectid, TblEmployee.GroupID, pr.Salary_account, 0, "
sql = sql & " GETDATE(), GETDATE(), p.NoofDays, p.NoofDays * p.value "
sql = sql & " FROM dbo.TblChangedComponentRegisterDetails p "
sql = sql & " INNER JOIN TblChangedComponentRegister PP ON PP.ChangedComponentid = p.ChangedComponentid "
sql = sql & " LEFT JOIN dbo.projects pr ON p.projectid = pr.id "
sql = sql & " LEFT JOIN dbo.TblEmployee ON p.Emp_id = TblEmployee.Emp_ID "
sql = sql & " LEFT JOIN dbo.mofrad ON PP.ComponentID = dbo.mofrad.ID "
sql = sql & " WHERE ISNULL(p.projectid, 0) <> 0 "
sql = sql & " AND (PP.Actualmonth = " & mMonth & ") "
sql = sql & " AND (p.Emp_id IN (" & empDes & ")) "
sql = sql & " AND (PP.Actualyear = " & val(CboYear.text) & ") "
sql = sql & " AND  mofrad.AddOrDiscount = 0"
sql = sql & " ORDER BY Total "


End If

   
  
 Dim mEmpRow As Long
 Dim mValue As Double
 'cccccccccccccccccc
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
             ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
            
            
            mEmpRow = Grid.FindRow(val(rs("EmpID").value & ""), Grid.FixedRows, Grid.ColIndex("Emp_ID"), False, True)

'.TextMatrix(i, .ColIndex("Emp_ID"))


mValue = val(rs!valuee & "") * (val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")))) / val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")))
'(val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")) - val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("AbcentDay"))
'/ val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays"))

'x = val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays"))) - val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("AbcentDay")))

'grid.TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(.TextMatrix(i, .ColIndex("AbcentDay"))) - val(val(.TextMatrix(i, .ColIndex("vacDay"))))
Balance = val(rs("NoDay") & "") * mValue 'IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)




             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    End If
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(DTP_Date, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If



' project gv

    Create_dev3 = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
    Create_dev3 = False
    'If SystemOptions.UserInterface = ArabicInterface Then
    'MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    'Else
    'MsgBox "Error During Saving", vbExclamation
    'End If
End Function
Function GetCarIDByEmpID(Optional Emp_id As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     fixedAssetid"
sql = sql & " From dbo.TblCarsData"
sql = sql & " Where (Emp_id = " & Emp_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarIDByEmpID = IIf(IsNull(rs2("fixedAssetid").value), 0, rs2("fixedAssetid").value)
Else
GetCarIDByEmpID = 0
End If
End Function
Function Create_dev4() As Boolean
    Dim CarID As Double
    Dim FixedID As Integer
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
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    Dim DepartmentID As Double
    If check_employee_accounts = False Then
        Exit Function
    End If

    If check_previous_dev(CboYear.text, CmbMonth.ListIndex + 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Else
            MsgBox "JV Alraedy Created To this Month", vbCritical
        End If

        Create_dev4 = False
        Exit Function
          
    End If
        
 
 
        
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
    Else
        Msg = " Salary Allocation JL Month: " & CmbMonth.text & "   Year: " & CboYear.text
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress
    DoEvents
    Dim StrSQL As String
 Set rsEmpID = New ADODB.Recordset
  StrSQL = "select * from TblEmployee "
  rsEmpID.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
 
    my_branch = Current_branch

    If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
            MsgBox " Can not start a new JL, you exceed the limit ": GoTo ErrTrap
        End If

    Else
                       
        If Notes_coding(val(my_branch), DTP_Date.value) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                MsgBox "   Can not Create a new JV , you Select Manual Numbering in JV Voucher Coding ": GoTo ErrTrap
            End If

        Else
            notes_serial = Notes_coding(val(my_branch), DTP_Date.value)
        End If
    End If
               
    If SystemOptions.UserInterface = ArabicInterface Then
        ALLButton2.Caption = "Ì—ÃÏ «·«‰ Ÿ«— Õ Ï «·«‰ Â«¡"
    Else
        ALLButton2.Caption = "Wait this process may take several minutre"
    End If

    rs.AddNew
    
    rs("branch_no").value = Current_branch
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = net_value ' Null

    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 66
    rs("NoteDate").value = DTP_Date.value
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("sanad_year").value = year(DTP_Date.value)
    rs("sanad_month").value = Month(DTP_Date.value)
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
    Dim BranchID As Integer
    Dim EmpID As Double
    Dim CValue As Double
    Dim EmpID1 As Integer
    Dim ProjectID As Integer
    
    BranchID = 1

    With Grid

        For j = 1 To 40
            If rsEmpID.RecordCount > 0 Then
                rsEmpID.MoveFirst
            End If

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                    For EmpID1 = 1 To rsEmpID.RecordCount
                                 
                        EmpID = IIf(IsNull(rsEmpID("Emp_ID").value), 0, (rsEmpID("Emp_ID").value))
                        BranchID = IIf(IsNull(rsEmpID("BranchId").value), 1, (rsEmpID("BranchId").value))
                        CValue = GetComponentValuePerEmpID(EmpID, ColumnName)
                      '   CarID = GetCarIDByEmpID(EmpID)
                         FixedID = GetCarIDByEmpID(EmpID)
                         CarID = FixedID
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , FixedID, , , BranchID, CarID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                      End If
                        rsEmpID.MoveNext
                    Next EmpID1
                             
                End If
                             
            End If
    
        Next j
       
                          
        For j = 1 To 40 ' Œ’Ê„« 

            If rsEmpID.RecordCount > 0 Then
                rsEmpID.MoveFirst
            End If
            ColumnName = "Comp" & j
            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                                                   
                    For EmpID1 = 1 To rsEmpID.RecordCount
                                 
                        EmpID = IIf(IsNull(rsEmpID("Emp_ID").value), 0, (rsEmpID("Emp_ID").value))
                        BranchID = IIf(IsNull(rsEmpID("BranchId").value), 1, (rsEmpID("BranchId").value))
                        CValue = GetComponentValuePerEmpID(EmpID, ColumnName)
                        'CarID = GetCarIDByEmpID(EmpID)
                        FixedID = GetCarIDByEmpID(EmpID)
                        CarID = FixedID
                         If CValue > 0 Then
                     If SystemOptions.ProjectDiscountPolicy = 1 Then
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , FixedID, , , BranchID, CarID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            Else
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , FixedID, , , BranchID, CarID) = False Then
                                GoTo ErrTrap
                            End If
                       End If
                            line_no = line_no + 1
                        End If
              
                        rsEmpID.MoveNext
                    Next EmpID1
                                    
                End If
            End If
    
        Next j


        For i = .FixedRows To .rows - 2
    
            If .TextMatrix(i, .ColIndex("EmpTotalNet")) > 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
     
                 If .TextMatrix(i, .ColIndex("EmpTotalNet")) < 0 Then '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("ToalInsurance")))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
            
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And ZmamAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
 
 
 
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                    StrAccountCode = Employee_account
                                 If AddOrDiscount(j) = 0 Then
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                                                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If

                        
                        
                        End If
                        
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With

  If SystemOptions.ProjectEmployeeGV = True Then
rs.Close
    Dim sql As String
    
    Dim Balance As Double
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount ,dbo.TblEmployee.DepartmentID"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(DTP_Date.value, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(DTP_Date.value, True) & "))"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount,dbo.TblEmployee.DepartmentID"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
    ' DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    '
                     '            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                     '       GoTo ErrTrap
                     '   End If
        
        
    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
                                line_no = line_no + 1
                                                             If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
                        
                        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
            line_no = line_no + 1
        
        
         
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
    
  
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code ,dbo.TblEmployee.DepartmentID"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(DTP_Date, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & " ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 

  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
   '  DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                         If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close

    
    
   ' Õ„Ì· «·„’—Ê›«  ⁄·Ï «·„‘«—Ì⁄
    
       sql = "SELECT      SUM(ROUND(dbo.EmpSalaryComponent.[Value] * dbo.opr_employee_details.[interval] / 30, " & SystemOptions.EmpSalaryDigts & ")) AS Total, dbo.mofrad.Account_Code, "
sql = sql & " dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years AS [year],"
sql = sql & " dbo.opr_Employee.Months, SUM(dbo.opr_employee_details.[interval]) AS Intervals, dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name,"
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId ,dbo.TblEmployee.DepartmentID"
sql = sql & " FROM         dbo.opr_employee_details INNER JOIN"
sql = sql & " dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id INNER JOIN"
sql = sql & " dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.opr_employee_details.Emp_id = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
sql = sql & " GROUP BY dbo.mofrad.Account_Code, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years, dbo.opr_Employee.Months,"
sql = sql & " dbo.MOFRAD.AddOrDiscount , dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name, dbo.Projects.Project_name, dbo.TblEmployee.BranchId"
sql = sql & " HAVING      (dbo.EmpSalaryComponent.EntIncresDataM IS NULL  OR"
sql = sql & "  dbo.EmpSalaryComponent.EntIncresDataM >= " & SQLDate(DTP_Date, True) & " )"
sql = sql & "   AND (dbo.opr_Employee.Months = " & CmbMonth.ListIndex & ") AND (2006 + dbo.opr_Employee.Years = " & CboYear.text & ")"


sql = sql & " ORDER BY dbo.opr_employee_details.ProjectID"

 
   
  
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     
             ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3 ,dbo.TblEmployee.DepartmentID"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(DTP_Date, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
      DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If


'«· √„Ì‰« 


    rs.Close
    
       
       sql = " "
' «„Ì‰«  ›Ì «·„”Ì—
Dim Nationality  As String

    GetInsuranceAccount mofradAccount, mofradAccount1
 With Grid
                         For i = .FixedRows To .rows - 2
                
                            If val(.TextMatrix(i, .ColIndex("ToalInsurance"))) > 0 Then '
                            Emp_id = val(.TextMatrix(i, .ColIndex("Emp_ID")))
                   emp_Name = (.TextMatrix(i, .ColIndex("emp_Name")))
                                Nationality = (.TextMatrix(i, .ColIndex("Nationality")))
                                    If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
                                                mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                                     End If
                     mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")  '«·«ÃÊ— «·„” Õﬁ…
                                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) <> 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰

                                If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 0, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›    " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                '
                                line_no = line_no + 1
                                End If
                                
                                      If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 1, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›   " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                
                                line_no = line_no + 1
                                
                                
                            End If
                 
                Next i

End With

'  √„Ì‰«  «Ã«“« 
sql = "  SELECT      dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, " & CHR(13)
sql = sql & "                       dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID,SUM(dbo.EmpSalaryComponent.[Value]) AS value, dbo.mofrad.Account_Code, dbo.mofrad.Account_code1" & CHR(13)
sql = sql & "  FROM         dbo.mofrad INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN" & CHR(13)
sql = sql & "                         dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID" & CHR(13)
sql = sql & "    WHERE     (dbo.mofrad.Insurances = 1) AND (dbo.EmpSalaryComponent.emp_ID IN" & CHR(13)
sql = sql & "                             (SELECT     Emp_ID" & CHR(13)
sql = sql & "                                From TblEmployee" & CHR(13)
sql = sql & "                                WHERE     dbo.TblEmployee.jopstatusid IN" & CHR(13)
sql = sql & "                                                          (SELECT     id" & CHR(13)
sql = sql & "                                                             From dbo.jopstatus" & CHR(13)
sql = sql & "                                                             WHERE     Insurances = 1 and id<>1  ) AND dbo.TblEmployee.BignDateWork <" & SQLDate(DTP_Date.value, True) & ")) AND" & CHR(13)
sql = sql & "                         (year(dbo.EmpSalaryComponent.EntIncresDataM)<year( " & SQLDate(DTP_Date.value, True) & ") OR" & CHR(13)
                      sql = sql & "   dbo.EmpSalaryComponent.EntIncresDataM IS NULL) AND (dbo.mofrad.Insurances = 1)" & CHR(13)
sql = sql & "   GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID," & CHR(13)
sql = sql & "                         dbo.MOFRAD.Account_Code , dbo.MOFRAD.Account_code1, dbo.TblEmployee.fullcode,dbo.TblEmployee.Nationality" & CHR(13)
sql = sql & "    ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)





' = 0
'ResidentVal1 = 0
Dim CitizenVal1 As Double
Dim ResidentVal1 As Double
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     'mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     'mofradAccount1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
     
        GetInsuranceAccount mofradAccount, mofradAccount1, CitizenVal1, ResidentVal1
     Nationality = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
       Emp_id = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
      
     If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
     
        mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                 
  Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * CitizenVal1 / 100
  
  Else
 Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * ResidentVal1 / 100
     End If
     
     
     
      
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     mofradname = IIf(IsNull(rs("AccountName").value), "", rs("AccountName").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
   
      emp_Name = IIf(IsNull(rs("emp_Name").value), "", rs("emp_Name").value)
                             If mofradAccount <> "" And mofradAccount1 <> "" And Balance > 0 Then
                                   
                                  
                                   If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, Balance, 0, Msg & mofradname & "   √Ì‰« -€Ì— „ÊÃÊœ… »«·„”Ì— " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                        
                                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  √Ì‰«  -€Ì— „ÊÃÊœ… »«·„”Ì—" & "  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                             
                                             
                                             
                             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    


' project gv

    Create_dev4 = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
    Create_dev4 = False
    'If SystemOptions.UserInterface = ArabicInterface Then
    'MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    'Else
    'MsgBox "Error During Saving", vbExclamation
    'End If
End Function

Function Create_dev2() As Boolean
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
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        
    If check_employee_accounts = False Then
        Exit Function
    End If

    If check_previous_dev(CboYear.text, CmbMonth.ListIndex + 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Else
            MsgBox "JV Alraedy Created To this Month", vbCritical
        End If

        Create_dev2 = False
        Exit Function
          
    End If
        
 
 
        
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
    Else
        Msg = " Salary Allocation JL Month: " & CmbMonth.text & "   Year: " & CboYear.text
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress
    DoEvents
 
    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
 
    'notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    my_branch = Current_branch

    If Notes_coding(val(my_branch), DTP_Date.value) = "error" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
            MsgBox " Can not start a new JL, you exceed the limit ": GoTo ErrTrap
        End If

    Else
                       
        If Notes_coding(val(my_branch), DTP_Date.value) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                MsgBox "   Can not Create a new JV , you Select Manual Numbering in JV Voucher Coding ": GoTo ErrTrap
            End If

        Else
            notes_serial = Notes_coding(val(my_branch), DTP_Date.value)
        End If
    End If
               
    If SystemOptions.UserInterface = ArabicInterface Then
        ALLButton2.Caption = "Ì—ÃÏ «·«‰ Ÿ«— Õ Ï «·«‰ Â«¡"
    Else
        ALLButton2.Caption = "Wait this process may take several minutre"
    End If

    rs.AddNew
    
    rs("branch_no").value = Current_branch
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = net_value ' Null

    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 66
    rs("NoteDate").value = DTP_Date.value
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("sanad_year").value = year(DTP_Date.value)
    rs("sanad_month").value = Month(DTP_Date.value)
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
    Dim BranchID As Integer
    Dim DepartmentID As Double
    Dim CValue As Double
    Dim Branch As Integer
    Dim Dept As Integer
    Dim ProjectID As Integer
    
    BranchID = 1

    With Grid

        For j = 1 To 40

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                Dim DepartmentName     As String
                    For Branch = 1 To rsBranch.RecordCount
                                 
                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                    If SystemOptions.SalaryJLByManagement = True Then
                             If RsDepartment.RecordCount > 0 Then
                             RsDepartment.MoveFirst
                             End If
                      For Dept = 1 To RsDepartment.RecordCount
                                If SystemOptions.UserInterface = ArabicInterface Then
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentName").value), "", (RsDepartment("DepartmentName").value))
                         Else
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentNameE").value), "", (RsDepartment("DepartmentNameE").value))
                         End If
                         
                         DepartmentID = IIf(IsNull(RsDepartment("DeparmentID").value), 0, (RsDepartment("DeparmentID").value))
                         CValue = GetComponentValuePerBranch(BranchID, ColumnName, DepartmentID)
                               
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)) & "  " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If
                        RsDepartment.MoveNext
                      Next Dept
                         
                     Else
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If
                    End If
                        rsBranch.MoveNext
                    Next Branch
                             
                End If
                             
            End If
    
        Next j
       
        '«·„ﬂ«›√ 
        If rsBranch.RecordCount > 0 Then
            rsBranch.MoveFirst
        End If
            
        For Branch = 1 To rsBranch.RecordCount
                                 
            BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
            CValue = GetComponentValuePerBranch(BranchID, "Mokafea")
                               
            If CValue > 0 Then
            
                If CValue > 0 Then
              '      If ModAccounts.AddNewDev(LngDevID, line_no, BonusAccount, CValue, 0, Msg & "   „ﬂ«›√ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
              '          GoTo ErrTrap
              '      End If

                    line_no = line_no + 1
                End If
                                    
            End If

            rsBranch.MoveNext
        Next Branch
                                    
        '«·Œ’Ê„« 
        If rsBranch.RecordCount > 0 Then
            rsBranch.MoveFirst
        End If
            
        For Branch = 1 To rsBranch.RecordCount
                                 
            BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
            CValue = GetComponentValuePerBranch(BranchID, "TotalDiscount")
                               
            If CValue > 0 Then
    
                If CValue > 0 Then
                
               ' If SystemOptions.ProjectEmployeeGV = False Then
              '      If ModAccounts.AddNewDev(LngDevID, line_no, DiscountAccount, CValue, 1, Msg & "  Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
              '          GoTo ErrTrap
              '      End If
               '  End If
                 
                    line_no = line_no + 1
                End If
                                    
            End If

            rsBranch.MoveNext
        Next Branch
                                    
        For j = 1 To 40 ' Œ’Ê„« 

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                                                   
                    For Branch = 1 To rsBranch.RecordCount
                                 
                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                               If SystemOptions.SalaryJLByManagement = True Then
                             If RsDepartment.RecordCount > 0 Then
                             RsDepartment.MoveFirst
                             End If
                      For Dept = 1 To RsDepartment.RecordCount
                                      If SystemOptions.UserInterface = ArabicInterface Then
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentName").value), "", (RsDepartment("DepartmentName").value))
                         Else
                         DepartmentName = IIf(IsNull(RsDepartment("DepartmentNameE").value), "", (RsDepartment("DepartmentNameE").value))
                         End If
                         DepartmentID = IIf(IsNull(RsDepartment("DeparmentID").value), 0, (RsDepartment("DeparmentID").value))
                         CValue = GetComponentValuePerBranch(BranchID, ColumnName, DepartmentID)
                                                 If CValue > 0 Then
            
 If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            Else
                            
                                     If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            
                            
End If
                            line_no = line_no + 1
                        End If
                      RsDepartment.MoveNext
                      Next Dept
                        Else
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               
                        If CValue > 0 Then
                   '      SystemOptions.ProjectEmployeeGV = True
 If SystemOptions.ProjectDiscountPolicy = 1 Then
 'Dim CurrentAccount As String
' CurrentAccount = Account_Code(j)
                           '  If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                     '   If Account_code1(j) <> "" Then
                                     '   CurrentAccount = Account_code1(j)
                                     '   End If
                            
                             
                           '  End If
                             
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            Else
                            
                                     If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & "   Œ’Ê„«  " & "  " & DepartmentName, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            
                            
End If
                            line_no = line_no + 1
                        End If
                      End If
                        rsBranch.MoveNext
                    Next Branch
                                    
                End If
            End If
    
        Next j

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) > 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
     
                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) < 0 Then '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("ToalInsurance")))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
            
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And ZmamAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
 
 
 
 
 
 
 
 
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                    StrAccountCode = Employee_account
                                 If AddOrDiscount(j) = 0 Then
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                                                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If

                        
                        
                        End If
                        
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With

  If SystemOptions.ProjectEmployeeGV = True Then
rs.Close
    Dim sql As String
    
    Dim Balance As Double
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount ,dbo.TblEmployee.DepartmentID"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(DTP_Date.value, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(DTP_Date.value, True) & "))"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    '
                     '            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                     '       GoTo ErrTrap
                     '   End If
        
        
    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
                                line_no = line_no + 1
                                                             If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
                        
                        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
        
            line_no = line_no + 1
        
        
         
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
    
  
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code ,dbo.TblEmployee.DepartmentID"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(DTP_Date, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & " ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 

  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                         If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close

    
    
   ' Õ„Ì· «·„’—Ê›«  ⁄·Ï «·„‘«—Ì⁄
    
       sql = "SELECT      SUM(ROUND(dbo.EmpSalaryComponent.[Value] * dbo.opr_employee_details.[interval] / 30, 2)) AS Total, dbo.mofrad.Account_Code, "
sql = sql & " dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years AS [year],"
sql = sql & " dbo.opr_Employee.Months, SUM(dbo.opr_employee_details.[interval]) AS Intervals, dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name,"
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId ,dbo.TblEmployee.DepartmentID"
sql = sql & " FROM         dbo.opr_employee_details INNER JOIN"
sql = sql & " dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id INNER JOIN"
sql = sql & " dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.opr_employee_details.Emp_id = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
sql = sql & " GROUP BY dbo.mofrad.Account_Code, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years, dbo.opr_Employee.Months,"
sql = sql & " dbo.MOFRAD.AddOrDiscount , dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name, dbo.Projects.Project_name, dbo.TblEmployee.BranchId"
sql = sql & " HAVING      (dbo.EmpSalaryComponent.EntIncresDataM IS NULL  OR"
sql = sql & "  dbo.EmpSalaryComponent.EntIncresDataM >= " & SQLDate(DTP_Date, True) & " )"
sql = sql & "   AND (dbo.opr_Employee.Months = " & CmbMonth.ListIndex & ") AND (2006 + dbo.opr_Employee.Years = " & CboYear.text & ")"


sql = sql & " ORDER BY dbo.opr_employee_details.ProjectID"

 
   
  
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     
             ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3 ,dbo.TblEmployee.DepartmentID"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(DTP_Date, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(DTP_Date, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     ProjectID = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
      DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If


'«· √„Ì‰« 


    rs.Close
    
       
       sql = " "
' «„Ì‰«  ›Ì «·„”Ì—
Dim Nationality  As String

    GetInsuranceAccount mofradAccount, mofradAccount1
 With Grid
                         For i = .FixedRows To .rows - 2
                
                            If val(.TextMatrix(i, .ColIndex("ToalInsurance"))) > 0 Then '
                            Emp_id = val(.TextMatrix(i, .ColIndex("Emp_ID")))
                   emp_Name = (.TextMatrix(i, .ColIndex("emp_Name")))
                                Nationality = (.TextMatrix(i, .ColIndex("Nationality")))
                                    If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
                                                mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                                     End If
                     mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")  '«·«ÃÊ— «·„” Õﬁ…
                                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) <> 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰

                                If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 0, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›    " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                '
                                line_no = line_no + 1
                                End If
                                
                                      If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, val(.TextMatrix(i, .ColIndex("ToalInsurance"))), 1, Msg & " √„Ì‰«  Õ’… «·„ÊŸ›   " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , val(.TextMatrix(i, .ColIndex("dep"))), val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                    GoTo ErrTrap
                                End If
                
                                line_no = line_no + 1
                                
                                
                            End If
                 
                Next i

End With

'  √„Ì‰«  «Ã«“« 
sql = "  SELECT      dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, " & CHR(13)
sql = sql & "                       dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID,SUM(dbo.EmpSalaryComponent.[Value]) AS value, dbo.mofrad.Account_Code, dbo.mofrad.Account_code1" & CHR(13)
sql = sql & "  FROM         dbo.mofrad INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN" & CHR(13)
sql = sql & "                         dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode INNER JOIN" & CHR(13)
                      sql = sql & "   dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID" & CHR(13)
sql = sql & "    WHERE     (dbo.mofrad.Insurances = 1) AND (dbo.EmpSalaryComponent.emp_ID IN" & CHR(13)
sql = sql & "                             (SELECT     Emp_ID" & CHR(13)
sql = sql & "                                From TblEmployee" & CHR(13)
sql = sql & "                                WHERE     dbo.TblEmployee.jopstatusid IN" & CHR(13)
sql = sql & "                                                          (SELECT     id" & CHR(13)
sql = sql & "                                                             From dbo.jopstatus" & CHR(13)
sql = sql & "                                                             WHERE     Insurances = 1 and id<>1  ) AND dbo.TblEmployee.BignDateWork <" & SQLDate(DTP_Date.value, True) & ")) AND" & CHR(13)
sql = sql & "                         (year(dbo.EmpSalaryComponent.EntIncresDataM)<year( " & SQLDate(DTP_Date.value, True) & ") OR" & CHR(13)
                      sql = sql & "   dbo.EmpSalaryComponent.EntIncresDataM IS NULL) AND (dbo.mofrad.Insurances = 1)" & CHR(13)
sql = sql & "   GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.DepartmentID," & CHR(13)
sql = sql & "                         dbo.MOFRAD.Account_Code , dbo.MOFRAD.Account_code1, dbo.TblEmployee.fullcode,dbo.TblEmployee.Nationality" & CHR(13)
sql = sql & "    ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)





' = 0
'ResidentVal1 = 0
Dim CitizenVal1 As Double
Dim ResidentVal1 As Double
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     'mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     'mofradAccount1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
     
        GetInsuranceAccount mofradAccount, mofradAccount1, CitizenVal1, ResidentVal1
     Nationality = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
       Emp_id = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
      
     If Nationality = "”⁄ÊœÌ" Or Nationality = "”⁄ÊœÏ" Or Nationality = "Saudi" Then      '”⁄ÊœÌ
     
        mofradAccount1 = get_EMPLOYEE_Account(CStr(Emp_id), "Account_Code1")   '«·«ÃÊ— «·„” Õﬁ…
                 
  Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * CitizenVal1 / 100
  
  Else
 Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value) * ResidentVal1 / 100
     End If
     
     
     
      
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     mofradname = IIf(IsNull(rs("AccountName").value), "", rs("AccountName").value)
     DepartmentID = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
   
      emp_Name = IIf(IsNull(rs("emp_Name").value), "", rs("emp_Name").value)
                             If mofradAccount <> "" And mofradAccount1 <> "" And Balance > 0 Then
                                   
                                  
                                   If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, Balance, 0, Msg & mofradname & "   √Ì‰« -€Ì— „ÊÃÊœ… »«·„”Ì— " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , ProjectID, , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                        
                                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  √Ì‰«  -€Ì— „ÊÃÊœ… »«·„”Ì—" & "  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , DepartmentID, Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                             
                                             
                                             
                             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    


' project gv

    Create_dev2 = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
    Create_dev2 = False
    'If SystemOptions.UserInterface = ArabicInterface Then
    'MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    'Else
    'MsgBox "Error During Saving", vbExclamation
    'End If
End Function

Function setfoxy_Line() As Double
    Dim last_line_id  As String
    last_line_id = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = last_line_id
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = last_line_id
 
    rs.update
    
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
        
    If check_previous_dev(CboYear.text, CmbMonth.text) Then
        MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Exit Function
    End If
        
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
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .rows - 2

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
    rs("m_year").value = CboYear.text
    rs("m_month").value = CmbMonth.text
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function CuurentLogdata(Optional Currentmode As String)

End Function
Sub UnUpdatePrePaid()
Dim sql As String
 sql = "UPDATE TblPripaidExpChiled"
 sql = sql & "    Set PerolPaed = Null "
 sql = sql & " where PeriolID=" & CboYear.text & CmbMonth.ListIndex + 1 & ""
 
  'If val(dcBranch1.BoundText) <> 0 Then
  '     sql = sql & "and TblEmpAdvance.emp_id in(select emp_id from tblemployee where branchid= " & val(dcBranch1.BoundText) & ")"
  '
  '  End If
    
    
                 Cn.Execute sql
End Sub
Sub UnUpdateAdvanc()
Dim i As Integer
Dim sql As String

 sql = " update dbo.TblEmpAdvanceDetails set     dbo.TblEmpAdvanceDetails.Payed=Null ,dbo.TblEmpAdvanceDetails.StutsID=666"
 sql = sql & "  FROM         dbo.TblEmpAdvance INNER JOIN"
 sql = sql & "                       dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
 sql = sql & "  Where ((Month(dbo.TblEmpAdvanceDetails.PartDate) = " & val(CmbMonth.ListIndex) + 1 & ") And (year(dbo.TblEmpAdvanceDetails.PartDate) = " & val(CboYear.text) & ")or (dbo.TblEmpAdvanceDetails.MothID2 = " & val(CmbMonth.ListIndex) + 1 & ") And (dbo.TblEmpAdvanceDetails.YearID2 = " & val(CboYear.text) & ") )and dbo.TblEmpAdvanceDetails.StutsID=555 and dbo.TblEmpAdvanceDetails.StutsID<>31 and dbo.TblEmpAdvanceDetails.StutsID<>12"
 
  If val(dcBranch1.BoundText) <> 0 Then
       sql = sql & "and TblEmpAdvance.emp_id in(select emp_id from tblemployee where branchid= " & val(dcBranch1.BoundText) & ")"
     
    End If
    
    
 Cn.Execute sql
End Sub
Sub UpdateAdvanc()
Dim i As Integer
Dim sql As String
With Grid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 And val(.TextMatrix(i, .ColIndex("TotalAdvance"))) <> 0 Then

 sql = " update dbo.TblEmpAdvanceDetails set     dbo.TblEmpAdvanceDetails.Payed=1,dbo.TblEmpAdvanceDetails.StutsID=555"
 sql = sql & "  FROM         dbo.TblEmpAdvance INNER JOIN"
 sql = sql & "                       dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
 sql = sql & "  Where  dbo.TblEmpAdvanceDetails.StutsID<>31 and dbo.TblEmpAdvanceDetails.StutsID<>12 and  ((Month(dbo.TblEmpAdvanceDetails.PartDate) = " & val(CmbMonth.ListIndex) + 1 & ") And (year(dbo.TblEmpAdvanceDetails.PartDate) = " & val(CboYear.text) & ")or (dbo.TblEmpAdvanceDetails.MothID2 = " & val(CmbMonth.ListIndex) + 1 & ") And (dbo.TblEmpAdvanceDetails.YearID2 = " & val(CboYear.text) & ")) And (dbo.TblEmpAdvance.Emp_id = " & val(.TextMatrix(i, .ColIndex("Emp_ID"))) & ")"
 
 'If val(dcBranch1.BoundText) <> 0 Then
 '      sql = sql & " And (dbo.TblEmpAdvance.BranchId=" & val(dcBranch1.BoundText)
     
 '   End If
    
 
 Cn.Execute sql
             End If
            Next i
End With
End Sub

Sub UpdatePrePaid()
Dim i As Integer
Dim sql As String
With Grid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 Then
 sql = "UPDATE TblPripaidExpChiled"
 sql = sql & "    Set PerolPaed = 1 ,PeriolID=" & CboYear.text & CmbMonth.ListIndex + 1 & ""
 sql = sql & " where id=" & val(.TextMatrix(i, .ColIndex("PrePaidID"))) & ""
                 Cn.Execute sql
             End If
            Next i
End With
End Sub
Sub UpdateInsurance()
Dim i As Integer
Dim sql As String
With Grid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 Then
 sql = "UPDATE TBLInsurancesJoin"
 sql = sql & "    Set TBLInsurancesJoin.payed = 1"
 sql = sql & " FROM TBLInsurancesJoin A"
 sql = sql & "   INNER JOIN TBLInsurances B ON"
 sql = sql & "       a.IDINS = b.IDINS"
 sql = sql & "   Where b.Monthe =" & val(CmbMonth.ListIndex) & " And b.SubYear = " & val(CboYear.ListIndex) & " And a.EmpCode =" & val(.TextMatrix(i, .ColIndex("Emp_ID"))) & " and A.BranchID =" & val(.TextMatrix(i, .ColIndex("BranchId"))) & ""
                 Cn.Execute sql
             End If
            Next i
End With
End Sub
Sub UpdateVoCationPayed()
Dim sql As String
 sql = "UPDATE TblRegsterSickleaveDet"
 sql = sql & "    Set payed = 1"
 sql = sql & "   Where MonthID =" & val(CmbMonth.ListIndex) & " And YearID = " & val(CboYear.text) & " "
             Cn.Execute sql
End Sub
Sub UnUpdateVoCationPayed()
Dim sql As String
 sql = "UPDATE TblRegsterSickleaveDet"
 sql = sql & "    Set payed =Null"
 sql = sql & "   Where MonthID =" & val(CmbMonth.ListIndex) & " And YearID = " & val(CboYear.text) & " "
             Cn.Execute sql
End Sub

Sub UpdateInsurancePayed()
Dim sql As String
 sql = "UPDATE TBLInsurancesJoin"
 sql = sql & "    Set TBLInsurancesJoin.payed = Null"
 sql = sql & " FROM TBLInsurancesJoin A"
 sql = sql & "   INNER JOIN TBLInsurances B ON"
 sql = sql & "       a.IDINS = b.IDINS"
 sql = sql & "   Where b.Monthe =" & val(CmbMonth.ListIndex) & " And b.SubYear = " & val(CboYear.ListIndex) & " "
             Cn.Execute sql
End Sub
Private Sub ALLButton2_Click()
'On Error Resume Next
    DcEmp.text = ""
    Dcdep.text = ""
    dcproject.text = ""

    'FillGridWithData
    DoEvents
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text
    DTP_Date.value = MonthLastDay(CDate(str))
      If ChekClodePeriod(DTP_Date.value) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ  €ÌÌ— «· «—ÌŒ  ·«‰ Â–Â «·› —… „€·ﬁ…"
    Else
    MsgBox "Please Change Date Becouse This is Period is Closed"
    End If
    Exit Sub
    End If
    If SystemOptions.LockSalary = True Then
    If check_Lock_Salary(CboYear.text, CmbMonth.ListIndex + 1) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·„”Ì— „ﬁ›· ·«Ì„ﬂ‰  ⁄œÌ·Â «Ê Õ–›Â", vbCritical
        Else
            MsgBox "JV Alraedy Locked", vbCritical
        End If

        Exit Sub
   End If
    End If
    If Grid.rows = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Õœœ ‘Â— «Ê·«", vbCritical
        Else
            MsgBox " Specify Month Firstly", vbCritical
        End If
        
        Exit Sub
    End If

    
    If detect_employee_work_type = 1 Then

If SystemOptions.SalaryJLByAnalyEqup = True Then
  
       If Create_dev4 = False Then
                      Exit Sub
                End If
Else
'SystemOptions.ProjectEmployeeGV = True
  If SystemOptions.ProjectEmployeeGV = True Then
        If Create_dev3 = False Then
                      Exit Sub
                End If
  Else
  
       If Create_dev2 = False Then
                      Exit Sub
                End If
  End If
  End If
'                If getNoOfBranches = 1 Then
'                    If Create_dev2 = False Then
'                        Exit Sub
'                    End If
'
'                Else
'
'                    If Create_dev2 = False Then
'                        Exit Sub
'                    End If


'                End If
        
        Else
           Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress
    DoEvents
 
        
    End If
    
    
    

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
    If SystemOptions.UserInterface = ArabicInterface Then
        If detect_employee_work_type = 1 Then
            MsgBox " „ «‰‘«¡   «·«” Õﬁ«ﬁ"
            Me.TxtNoteSerial.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 66)
            Me.TxtNoteSerial2.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 555)

        Else
            MsgBox " „ «‰‘«¡   ”‰œ «·—« »"
        End If
        
    Else
 
        If detect_employee_work_type = 1 Then
            MsgBox "JV  Create"
            Me.TxtNoteSerial.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 66)
            Me.TxtNoteSerial2.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 555)

        Else
            MsgBox "Salary Vchr Created"
        End If
       
    End If

    create_report_data

    DoEvents
    UpdatePrePaid
    UpdateInsurance
    UpdateVoCationPayed
    FillGridWithData2
    UpdateAdvanc
    'CmdOk_Click

    If SystemOptions.UserInterface = ArabicInterface Then
        If detect_employee_work_type = 1 Then
            ALLButton2.Caption = "«‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
            ALLButton2.Caption = "«‰‘«¡   ”‰œ «·—« »"
        End If

    Else
        ALLButton2.Caption = "Salary Allocation JL"
    End If

    LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & CHR(13) & " «·‘Â—     " & CmbMonth.text & CHR(13) & "  «·”‰…   " & CboYear.text & CHR(13) & " «· «—ÌŒ " & DTP_Date.value
                     
    LogTexte = ""
       AddToLogFile CInt(user_id), 66, Date, Time, LogTextA, LogTexte, Me.Name, "N", "", , val(TxtNoteSerial), ""
       
 
    FillTable
    
End Sub

Private Sub ALLButton3_Click()
    On Error Resume Next

    With GRID1

        If .rows = 3 And Not IsNumeric(.TextMatrix(1, .ColIndex("Emp_code"))) Then
            Exit Sub
        End If

    End With

    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim depit_side As String
    Dim credit_side As String
    Dim total_value As Double

    If Me.CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
        Else
            Msg = "Select Payment Method ...!!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPayMentType.SetFocus
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        If Trim(Me.DcboBox.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
            Else
                
                Msg = "Selet Box..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 2 Or Me.CboPayMentType.ListIndex = 3 Then

        If Me.DcboBankName.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ..!!"
            Else
                
                Msg = "Selet Bank..!!"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBankName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

        If Trim$(Me.TxtChequeNumber.text) = "" And Me.CboPayMentType.ListIndex = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            Else
                Msg = " Enter Cheque No....!!"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtChequeNumber.SetFocus
            Exit Sub
        End If

        '      If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
        '                                    If SystemOptions.UserInterface = ArabicInterface Then
        '                                      Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
        '                                  Else
        '                                  Msg = " Cheque Due Date Not Vaild...!!"
        '                                  End If
        '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '          DtpChequeDueDate.SetFocus
        '          SendKeys "{F4}"
        '          Exit Sub
        '      End If
    ElseIf Me.CboPayMentType.ListIndex = 4 Then

        If Me.DCAccount.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ≈Œ Ì«— «·Õ”«»..!!"
            Else
                
                Msg = "Selet Accounts..!!"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCAccount.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
            
    End If

    credit_side = ""
    depit_side = ""
    total_value = 0

    If Me.CboPayMentType.ListIndex = 0 Then

        credit_side = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

        If credit_side = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ Õ”«» «·Œ“Ì‰…": Exit Sub
            Else
                MsgBox "Error In Box Account": Exit Sub
            End If
        End If
                 
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 2 Or Me.CboPayMentType.ListIndex = 3 Then
    
        Dim rsbank As New ADODB.Recordset
        Set rsbank = New ADODB.Recordset
        rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
        If Not (rsbank.EOF Or rsbank.BOF) Then
            If rsbank!banks_Accounts = True Then
                  
                credit_side = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code")
            Else
                credit_side = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code")
            End If
                 
            If credit_side = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ Õ”«» «·»‰ﬂ": Exit Sub
                Else
                    MsgBox "Error In Bank Account": Exit Sub
                End If
            End If
        End If
        
    ElseIf Me.CboPayMentType.ListIndex = 4 Then

        If Me.DCAccount.BoundText <> "" Then
            credit_side = Me.DCAccount.BoundText
                
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ Õ”«»  ": Exit Sub
            Else
                MsgBox "Error In   Account": Exit Sub
            End If

        End If

    End If

    '«· √ﬂœ „‰ «Œ Ì«— „ÊŸ›Ì‰

    With GRID1

        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            
                GoTo SelectEmp
            End If

        Next i

    End With

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „  ÕœÌœ «Ì „ÊŸ› ··”œ«œ ·… :"
    Else
        MsgBox " there is No Employee Selected"
    End If

    Exit Sub

SelectEmp:

    With GRID1

        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            
                If get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_Code1") = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Â‰«ﬂ Œÿ√ ›Ì Õ”«»  «·«ÃÊ— «·„” Õﬁ… ···„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code"))
                    Else
                        MsgBox " Error In Employee Salary Allocation Account For Employee : " & .TextMatrix(i, .ColIndex("Emp_code"))
                    End If

                    Exit Sub
                End If
                   
            End If

        Next i

    End With
 
    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 
    cProgress.StartProgress

    DoEvents

    Dim StrSQL As String
    Dim notes_id As String
    Dim notes_serial As String
    Dim rs As New ADODB.Recordset
    Dim foxy_ked_NO As String
 
    StrSQL = "select * From Notes where NoteType=5 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
    foxy_ked_NO = CStr(new_id("foxy", "id", "", True))

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ﬁÌœ ”œ«œ —« » ⁄‰ ‘Â—   " & CmbMonth.text & "     ·”‰… " & CboYear.text
    Else
        Msg = "Salary Payment JL Month:    " & CmbMonth.text & "     Year " & CboYear.text
    End If

    With GRID1

        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                total_value = total_value + .TextMatrix(i, .ColIndex("EmpTotalNet"))
            End If

        Next i

    End With
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("branch_no").value = Current_branch
    ''
                 
    If Notes_coding(val(Current_branch), DTPicker1.value) = "error" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
            MsgBox " Can not start a new JL, you exceed the limit  ": Exit Sub
                      
        End If

    Else
                       
        If Notes_coding(val(Current_branch), DTPicker1.value) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  "
                                                
            Else
                MsgBox " Can not Create a new JL , you Select Manual Numbering in JL Voucher Coding   ": Exit Sub
            End If
                                                
            cProgress.FinishProgress
            cProgress.StopProgess
            Set cProgress = Nothing
                          
            Exit Sub
        Else
            notes_serial = Notes_coding(val(Current_branch), DTPicker1.value)
        End If
    End If

    rs("NoteSerial").value = notes_serial
                       
    'Rs("Note_Value").value = total_value
    rs("FOXY_NO").value = foxy_ked_NO
    
    rs("Note_Value").value = total_value ' Null
    rs("note_value_by_characters").value = WriteNo(Format(total_value, "0.00"), 0, True, ".")
    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 555
    rs("NoteDate").value = DTPicker1.value
    rs("UserID").value = user_id

    '
    If Me.CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = val(DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 0
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 1
    End If
    
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁ”œ
    rs("sanad_year").value = year(DTPicker1.value)
    rs("sanad_month").value = Month(DTPicker1.value)
        
    rs.update
    
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    total_value = 0
                
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double

    With GRID1

        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
                total_value = total_value + Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), SystemOptions.EmpSalaryDigts)
            
                depit_side = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_Code1")
                CURRENT_LINE = setfoxy_Line

                If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), SystemOptions.EmpSalaryDigts), 0, Msg, val(notes_id), , , , DTPicker1.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_id")))) = False Then
                        GoTo ErrTrap
                    End If
                End If
              
                If .TextMatrix(i, .ColIndex("cost_center_id")) <> "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        save_cost_center .TextMatrix(i, .ColIndex("cost_center_id")), "”‰œ ﬁÌœ ”œ«œ —« »", DTPicker1.value, .TextMatrix(i, .ColIndex("EmpTotalNet")), foxy_ked_NO, depit_side, .TextMatrix(i, .ColIndex("Emp_Name")), CURRENT_LINE
                    Else
                        save_cost_center .TextMatrix(i, .ColIndex("cost_center_id")), "Payment Salary JL", DTPicker1.value, .TextMatrix(i, .ColIndex("EmpTotalNet")), foxy_ked_NO, depit_side, .TextMatrix(i, .ColIndex("Emp_Name")), CURRENT_LINE
                    End If
                End If
            End If

        Next i

    End With
               
    If total_value > 0 Then
                        
        If getNoOfBranches = 1 Then
                                
            If ModAccounts.AddNewDev(LngDevID, i + 1, credit_side, total_value, 1, Msg, val(notes_id), , , , DTPicker1.value, user_id, 200, , , , , , , , setfoxy_Line, , , , , , , , , 1) = False Then
                GoTo ErrTrap
            End If
                                
        Else '›Ì Õ«·…  ⁄œ «·«›—Ê⁄
            Dim Branch As Integer
            Dim CValue  As Double

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            i = i + 1

            For Branch = 1 To rsBranch.RecordCount
                                                                         
                BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                                                        
                CValue = GetComponentValuePerBranch2(BranchID, "EmpTotalNet")
                                                                       
                If CValue > 0 Then
                                                    
                    If CValue > 0 Then
                        If ModAccounts.AddNewDev(LngDevID, i, credit_side, CValue, 1, Msg, val(notes_id), , , , DTPicker1.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If

                        i = i + 1
                    End If
                                                                            
                End If

                rsBranch.MoveNext
            Next Branch

        End If
                
    End If

    With GRID1

        For i = .FixedRows To .rows - 2
         
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                If Change_filed_value(val(.TextMatrix(i, .ColIndex("id"))), "id", "Payed", "emp_salary", 1) Then
                End If
            End If

        Next i

    End With

    Dim X As Integer
   
    FillGridWithData2

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox(" „ «‰‘«¡ ”‰œ «·”œ«œ —ﬁ„ «·ﬁÌœ ÂÊ " & CHR(13) & notes_serial & " Â·  —Ìœ ⁄—÷ «·ﬁÌœ ‰⁄„ «„ ·«", vbInformation + vbYesNo)

    Else
        X = MsgBox("   Voucher Created " & CHR(13) & notes_serial & "  Show GE", vbInformation + vbYesNo)
    End If

    If X = vbYes Then
        ShowGL_cc notes_serial, , 200
    End If
        
        
            LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & CHR(13) & " «·‘Â—     " & CmbMonth.text & CHR(13) & "  «·”‰…   " & CboYear.text & CHR(13) & " «· «—ÌŒ " & DTP_Date.value
                     
    LogTexte = ""
       AddToLogFile CInt(user_id), 555, Date, Time, LogTextA, LogTexte, Me.Name, "N", "", , val(TxtNoteSerial), ""
       
       
    '
ErrTrap:

    Exit Sub
    'Dim StrSQL As String
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Double_Entry_Vouchers_ID=" & LngDevID
    Cn.Execute StrSQL, , adExecuteNoRecords



 
 
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·ﬁÌœ ", vbCritical
    Else
        MsgBox "Error During Saving ", vbCritical
    End If

End Sub

Private Sub ALLButton4_Click()
    ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub ALLButton5_Click()
    ShowGL_cc Me.TxtNoteSerial2.text, , 200
End Sub
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim mMonth As Long
mMonth = val(CmbMonth.ListIndex) + 1
If empDes = "" Then: Exit Function
MySQL = " SELECT     TOP 100 PERCENT dbo.ProJectMofrdSalar.ID, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
MySQL = MySQL & "                       dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, dbo.ProJectMofrdSalar.Total, dbo.ProJectMofrdSalar.NoDay,"
MySQL = MySQL & "                       dbo.ProJectMofrdSalar.YearID, dbo.ProJectMofrdSalar.MonthID, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
MySQL = MySQL & "                       dbo.TblEmployee.ContractID, dbo.TblEmployee.GroupID, dbo.projects.Salary_account, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee3 , dbo.TblEmployee.Emp_Namee4 ,dbo.ProJectMofrdSalar.TypeSalary "
MySQL = MySQL & "  FROM         dbo.ProJectMofrdSalar LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "  Where (dbo.ProJectMofrdSalar.YearID = " & val(CboYear.ListIndex) & ") And (dbo.ProJectMofrdSalar.MonthID = " & val(CmbMonth.ListIndex) & ")  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"
MySQL = MySQL & "  ORDER BY dbo.ProJectMofrdSalar.ProjID, dbo.ProJectMofrdSalar.EmpID"

sql = " SELECT    distinct TOP 100 PERCENT mofrdat.mofrad_code, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
sql = sql & "                      dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
sql = sql & "                      dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, "
sql = sql & "                       dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
sql = sql & "                      dbo.TblEmployee.ContractID , dbo.TblEmployee.GroupID, dbo.Projects.Salary_account , dbo.ProJectMofrdSalar.TypeSalary,"

sql = sql & "                        NoDay = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & " THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END,"
sql = sql & "                             Total = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END * Valuee"
sql = sql & " FROM         dbo.ProJectMofrdSalar LEFT OUTER JOIN"
sql = sql & "                      dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID"
 

'sql = sql & "  WHERE     (" & SQLDate(DTP_Date.value, True) & "BETWEEN dbo.ProJectMofrdSalar.fromDate AND dbo.ProJectMofrdSalar.toDate) " & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

'sql = sql & "  Where (dbo.ProJectMofrdSalar.YearID = " & val(CboYear.ListIndex) & ") And (dbo.ProJectMofrdSalar.MonthID = " & val(CmbMonth.ListIndex) & ") and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  Where ( year(dbo.ProJectMofrdSalar.fromDate ) = " & val(CboYear.text) & ") "
'And ( month(dbo.ProJectMofrdSalar.fromDate) = " & val(CmbMonth.ListIndex) + 1 & ")

sql = sql & "  AND ("
sql = sql & "                 ("
sql = sql & "                     Month (dbo.ProJectMofrdSalar.FromDate) <= " & mMonth & " "
sql = sql & "                     AND ("
sql = sql & "                             ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                             OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                         )"
sql = sql & "                 )"
sql = sql & "                 OR MONTH(dbo.ProJectMofrdSalar.fromDate) = " & mMonth & " "
sql = sql & "             )"

sql = sql & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  ORDER BY dbo.ProJectMofrdSalar.ProjID ,dbo.ProJectMofrdSalar.EmpID"
 If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â·  —Ìœ ⁄„· Ã—Ê» ⁄·Ï „” ÊÏ «·„‘—Ê⁄"
Else
Msg = "Do you want to Group at the project level"
End If
 If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects1.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects1.rpt"
        End If
 Else
      If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjects.rpt"
        End If
 End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      Else
      Msg = "No Found Data"
      End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Sub RetriveProjectSalar()
Dim rs2 As ADODB.Recordset
Dim i As Integer
If empDes = "" Then: Exit Sub
Dim mMonth As Integer
mMonth = val(CmbMonth.ListIndex) + 1
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 2

End With

Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT    distinct TOP 100 PERCENT mofrdat.mofrad_code, dbo.ProJectMofrdSalar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
sql = sql & "                      dbo.TblEmployee.Emp_Namee, dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID,"
sql = sql & "                      dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, "
sql = sql & "                       dbo.TblEmployee.SalaryType, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.BranchId,"
sql = sql & "                      dbo.TblEmployee.ContractID , dbo.TblEmployee.GroupID, dbo.Projects.Salary_account , dbo.ProJectMofrdSalar.TypeSalary,"

sql = sql & "                        NoDay = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & " THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END,"
sql = sql & "                             Total = CASE"
sql = sql & "                                          WHEN ("
sql = sql & "                                                   ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                                                   OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                                               ) AND MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  THEN 30 - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "                                          WHEN ISNULL(ProJectMofrdSalar.ToDate, '') = '' AND MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  THEN 30"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) < " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DAY(ProJectMofrdSalar.ToDate)"
sql = sql & "                                          WHEN MONTH(ProJectMofrdSalar.FromDate) = " & mMonth & "  AND MONTH(ProJectMofrdSalar.ToDate) = " & mMonth & "  THEN DATEDIFF(d, ProJectMofrdSalar.FromDate, ProJectMofrdSalar.ToDate)"
sql = sql & "                                               + 1"
sql = sql & "                                     END * Valuee"
sql = sql & " FROM         dbo.ProJectMofrdSalar LEFT OUTER JOIN"
sql = sql & "                      dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpID = dbo.TblEmployee.Emp_ID"
 

'sql = sql & "  WHERE     (" & SQLDate(DTP_Date.value, True) & "BETWEEN dbo.ProJectMofrdSalar.fromDate AND dbo.ProJectMofrdSalar.toDate) " & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

'sql = sql & "  Where (dbo.ProJectMofrdSalar.YearID = " & val(CboYear.ListIndex) & ") And (dbo.ProJectMofrdSalar.MonthID = " & val(CmbMonth.ListIndex) & ") and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  Where ( year(dbo.ProJectMofrdSalar.fromDate ) = " & val(CboYear.text) & ") "
'And ( month(dbo.ProJectMofrdSalar.fromDate) = " & val(CmbMonth.ListIndex) + 1 & ")

sql = sql & "  AND ("
sql = sql & "                 ("
sql = sql & "                     Month (dbo.ProJectMofrdSalar.FromDate) <= " & mMonth & " "
sql = sql & "                     AND ("
sql = sql & "                             ISNULL(ProJectMofrdSalar.ToDate, '') = ''"
sql = sql & "                             OR MONTH(ProJectMofrdSalar.ToDate) > " & mMonth & " "
sql = sql & "                         )"
sql = sql & "                 )"
sql = sql & "                 OR MONTH(dbo.ProJectMofrdSalar.fromDate) = " & mMonth & " "
sql = sql & "             )"

sql = sql & "  and (dbo.ProJectMofrdSalar.EmpID in( " & empDes & "))"

sql = sql & "  ORDER BY dbo.ProJectMofrdSalar.ProjID ,dbo.ProJectMofrdSalar.EmpID"






  sql = " DECLARE @SelectedMonth INT"
sql = sql & " SET @SelectedMonth =" & mMonth

sql = sql & " ;WITH DateRanges AS ("
sql = sql & "    SELECT"
sql = sql & "      ProJectMofrdSalar.id,  mofrdat.mofrad_code"
sql = sql & "   ,dbo.ProJectMofrdSalar.EmpId"
sql = sql & "   ,dbo.TblEmployee.Emp_Name"
sql = sql & "   ,dbo.TblEmployee.Fullcode"
sql = sql & "   ,dbo.TblEmployee.Emp_Namee"
sql = sql & "   ,dbo.ProJectMofrdSalar.ProjID"
sql = sql & "   ,dbo.projects.Project_name"
sql = sql & "   ,dbo.projects.Project_nameE"
sql = sql & "   ,dbo.ProJectMofrdSalar.MofrdID"
sql = sql & "   ,dbo.mofrdat.mofrad_name"
sql = sql & "   ,dbo.mofrdat.mofrad_namee"
sql = sql & "   ,dbo.ProJectMofrdSalar.Valuee"
sql = sql & "   ,dbo.TblEmployee.SalaryType"
sql = sql & "   ,dbo.TblEmployee.DepartmentID"
sql = sql & "   ,dbo.TblEmployee.BranchId"
sql = sql & "   ,dbo.TblEmployee.ContractID"
sql = sql & "   ,dbo.TblEmployee.GroupID"
sql = sql & "   ,dbo.projects.Salary_account"
sql = sql & "   ,dbo.ProJectMofrdSalar.TypeSalary ,ProJectMofrdSalar.FromDate,ProJectMofrdSalar.ToDate"
sql = sql & " From dbo.ProJectMofrdSalar"
sql = sql & " LEFT OUTER JOIN dbo.mofrdat"
sql = sql & "    ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code"
sql = sql & " LEFT OUTER JOIN dbo.projects"
sql = sql & "    ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id"
sql = sql & " LEFT OUTER JOIN dbo.TblEmployee"
sql = sql & "    ON dbo.ProJectMofrdSalar.EmpId = dbo.TblEmployee.Emp_ID"

sql = sql & "    Where"
'sql = sql & "    dbo.ProJectMofrdSalar.EmpId IN (170) AND"
sql = sql & "   (dbo.ProJectMofrdSalar.EmpID in( " & empDes & ")) AND "
sql = sql & "       ( YEAR(ProJectMofrdSalar.FromDate) =  " & val(CboYear.text) & "   )"
sql = sql & "        AND ("
sql = sql & " (MONTH(ProJectMofrdSalar.FromDate) <= @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) >= @SelectedMonth))"
sql = sql & "            OR (MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth)"
sql = sql & "        )"
sql = sql & ")"

sql = sql & "SELECT"
   
    
  sql = sql & "    ProJectMofrdSalar.id,   mofrdat.mofrad_code"
sql = sql & "   ,dbo.ProJectMofrdSalar.EmpId"
sql = sql & "   ,dbo.TblEmployee.Emp_Name"
sql = sql & "   ,dbo.TblEmployee.Fullcode"
sql = sql & "   ,dbo.TblEmployee.Emp_Namee"
sql = sql & "   ,dbo.ProJectMofrdSalar.ProjID"
sql = sql & "   ,dbo.projects.Project_name"
sql = sql & "   ,dbo.projects.Project_nameE"
sql = sql & "   ,dbo.ProJectMofrdSalar.MofrdID"
sql = sql & "   ,dbo.mofrdat.mofrad_name"
sql = sql & "   ,dbo.mofrdat.mofrad_namee"
'   --,SUM(dbo.ProJectMofrdSalar.Valuee) VALUE
sql = sql & "   ,dbo.ProJectMofrdSalar.Valuee"
sql = sql & "   ,dbo.TblEmployee.SalaryType"
sql = sql & "   ,dbo.TblEmployee.DepartmentID"
sql = sql & "   ,dbo.TblEmployee.BranchId"
sql = sql & "   ,dbo.TblEmployee.ContractID"
sql = sql & "   ,dbo.TblEmployee.GroupID"
sql = sql & "   ,dbo.projects.Salary_account"
sql = sql & "   ,dbo.ProJectMofrdSalar.TypeSalary ,ProJectMofrdSalar.FromDate,ProJectMofrdSalar.ToDate"
sql = sql & "    ,NoDay ="
sql = sql & "        CASE"
sql = sql & " WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth THEN"

sql = sql & "                datediff(D,ProJectMofrdSalar.FromDate,ProJectMofrdSalar.ToDate) +1"
'sql = sql & "                dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth)  - (dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.ToDate), @SelectedMonth) - DAY(ProJectMofrdSalar.ToDate) )"
sql = sql & "            WHEN MONTH(ProJectMofrdSalar.FromDate) = @SelectedMonth THEN"
sql = sql & "                dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth) - DAY(ProJectMofrdSalar.FromDate) + 1"
sql = sql & "            WHEN MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth THEN"
sql = sql & "                day (ProJectMofrdSalar.ToDate)"
sql = sql & "            Else"
sql = sql & "                dbo.GetDaysInMonth2(YEAR(ProJectMofrdSalar.FromDate), @SelectedMonth)"
sql = sql & "        End"

'    -- »«ﬁÌ «·«⁄„œ… Â‰«

sql = sql & " From"
sql = sql & "     dbo.ProJectMofrdSalar"
sql = sql & "     LEFT OUTER JOIN dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code"
sql = sql & "     LEFT OUTER JOIN dbo.projects ON dbo.ProJectMofrdSalar.ProjID = dbo.projects.id"
sql = sql & "     LEFT OUTER JOIN dbo.TblEmployee ON dbo.ProJectMofrdSalar.EmpId = dbo.TblEmployee.Emp_ID"
sql = sql & "     JOIN DateRanges ON dbo.ProJectMofrdSalar.FromDate <= DateRanges.ToDate AND dbo.ProJectMofrdSalar.ToDate >= DateRanges.FromDate"
sql = sql & "     AND DateRanges.EmpID = ProJectMofrdSalar.EmpID"
sql = sql & " Where"


sql = sql & "   (dbo.ProJectMofrdSalar.EmpID in( " & empDes & ")) AND "
sql = sql & "       ( YEAR(ProJectMofrdSalar.FromDate) =  " & val(CboYear.text) & "   )"
sql = sql & "        AND ("
sql = sql & " (MONTH(ProJectMofrdSalar.FromDate) <= @SelectedMonth AND (ProJectMofrdSalar.ToDate IS NULL OR MONTH(ProJectMofrdSalar.ToDate) >= @SelectedMonth))"
sql = sql & "            OR (MONTH(ProJectMofrdSalar.FromDate) < @SelectedMonth AND MONTH(ProJectMofrdSalar.ToDate) = @SelectedMonth)"
sql = sql & "        )"
'sql = sql & ")"


sql = sql & "     Group By"
sql = sql & "          mofrdat.mofrad_code"
sql = sql & "    ,dbo.ProJectMofrdSalar.EmpId"
sql = sql & "    ,dbo.TblEmployee.Emp_Name"
sql = sql & "    ,dbo.TblEmployee.Fullcode"
sql = sql & "    ,dbo.TblEmployee.Emp_Namee"
sql = sql & "    ,dbo.ProJectMofrdSalar.ProjID"
sql = sql & "    ,dbo.projects.Project_name"
sql = sql & "    ,dbo.projects.Project_nameE"
sql = sql & "    ,dbo.ProJectMofrdSalar.MofrdID"
sql = sql & "    ,dbo.mofrdat.mofrad_name"
sql = sql & "    ,dbo.mofrdat.mofrad_namee"
sql = sql & "    ,dbo.ProJectMofrdSalar.Valuee,ProJectMofrdSalar.id"
sql = sql & "    ,dbo.TblEmployee.SalaryType"
sql = sql & "    ,dbo.TblEmployee.DepartmentID"
sql = sql & "    ,dbo.TblEmployee.BranchId"
sql = sql & "    ,dbo.TblEmployee.ContractID"
sql = sql & "    ,dbo.TblEmployee.GroupID"
sql = sql & "    ,dbo.projects.Salary_account"
sql = sql & "    ,dbo.ProJectMofrdSalar.TypeSalary ,ProJectMofrdSalar.FromDate,ProJectMofrdSalar.ToDate"
'sql = sql & " Order By"
'sql = sql & "     dbo.ProJectMofrdSalar.ProjID,"
'sql = sql & "     dbo.ProJectMofrdSalar.EmpID"




sql = sql & "   Union"

sql = sql & "   SELECT"
sql = sql & " id = 0,"
sql = sql & "       mofrad.ID AS mofrad_code,"
sql = sql & "       p.Emp_id,"
sql = sql & "       TblEmployee.Emp_Name,"
sql = sql & "       TblEmployee.Fullcode,"
sql = sql & "       TblEmployee.Emp_Namee,"
sql = sql & "       p.projectid AS ProjID,"
sql = sql & "       pr.Project_name,"
sql = sql & "       pr.Project_nameE,"
sql = sql & "       PP.ComponentID AS MofrdID,"
sql = sql & "       mofrad.name AS mofrad_name,"
sql = sql & "       mofrad.namee AS mofrad_namee,"
sql = sql & "       p.value AS Valuee,"
sql = sql & "       TblEmployee.SalaryType,"
sql = sql & "       TblEmployee.DepartmentID,"
sql = sql & "       TblEmployee.BranchId,"
sql = sql & "       TblEmployee.ContractID,"
sql = sql & "       TblEmployee.GroupID,"

sql = sql & "           PR.Salary_account,"
sql = sql & "           0 AS TypeSalary,"
    
    
sql = sql & "           FromDate = GETDATE(),"
sql = sql & "           ToDate = GETDATE(),"
sql = sql & "           p.NoofDays AS NoDay"
sql = sql & "       FROM dbo.TblChangedComponentRegisterDetails p"
sql = sql & "       INNER JOIN TblChangedComponentRegister PP"
sql = sql & "           ON PP.ChangedComponentid = P.ChangedComponentid"
sql = sql & "       LEFT JOIN dbo.projects pr"
sql = sql & "           ON p.projectid = pr.id"
sql = sql & "       LEFT JOIN dbo.TblEmployee"
sql = sql & "           ON p.Emp_id = TblEmployee.Emp_ID"

sql = sql & "           LEFT OUTER JOIN dbo.mofrad"
sql = sql & "           ON PP.ComponentID= dbo.mofrad.ID"

sql = sql & "           Where"
sql = sql & "            ISNULL(p.projectid, 0) <> 0"

sql = sql & "       AND ((PP.Actualmonth <=@SelectedMonth"
sql = sql & "       AND (PP.Actualmonth >= @SelectedMonth))"
sql = sql & "       OR (PP.Actualmonth < @SelectedMonth"
sql = sql & "       AND PP.Actualmonth = @SelectedMonth))"

sql = sql & "  AND (p.Emp_id in( " & empDes & ")) AND "
sql = sql & "       (PP.Actualyear =  " & val(CboYear.text) & "   )"


rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then 'xx
rs2.MoveFirst
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 2
.rows = .rows + rs2.RecordCount
Dim mEmpRow As Long
For i = .FixedRows To .rows - 2
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Salary_account")) = IIf(IsNull(rs2("Salary_account").value), "", rs2("Salary_account").value)
.TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)

mEmpRow = Grid.FindRow(val(rs2("EmpID").value & ""), Grid.FixedRows, Grid.ColIndex("Emp_ID"), False, True)

'.TextMatrix(i, .ColIndex("Emp_ID"))
.TextMatrix(i, .ColIndex("EmpID")) = val(IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value))
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("ProjID")) = IIf(IsNull(rs2("ProjID").value), 0, rs2("ProjID").value)
.TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(rs2("MofrdID").value), 0, rs2("MofrdID").value)
'.TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(rs2("Valuee").value), 0, rs2("Valuee").value)


.TextMatrix(i, .ColIndex("Valuee")) = val(rs2!valuee & "") * (val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays"))) - val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("AbcentDay")))) / val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")))
.TextMatrix(i, .ColIndex("Valuee")) = val(rs2!valuee & "") * (val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")))) / val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")))
'(val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays")) - val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("AbcentDay"))
'/ val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays"))

'x = val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("CountDays"))) - val(Grid.TextMatrix(mEmpRow, Grid.ColIndex("AbcentDay")))

'grid.TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(.TextMatrix(i, .ColIndex("AbcentDay"))) - val(val(.TextMatrix(i, .ColIndex("vacDay"))))
.TextMatrix(i, .ColIndex("Total")) = val(rs2("NoDay") & "") * val(.TextMatrix(i, .ColIndex("Valuee"))) 'IIf(IsNull(rs2("Total").value), 0, rs2("Total").value)
.TextMatrix(i, .ColIndex("NoDay")) = IIf(IsNull(rs2("NoDay").value), 0, rs2("NoDay").value)
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
Sub RetriveEmpMove()
Dim rs2 As ADODB.Recordset
Dim i As Integer

Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     dbo.TblMoveEmp1.ID, dbo.TblMoveEmp1.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
sql = sql & "                       dbo.EmpSalaryComponent.AccountCode, dbo.mofrdat.mofrad_type, dbo.mofrad.name, dbo.mofrad.nameE, dbo.EmpSalaryComponent.[Value],"
sql = sql & "                       dbo.TblMoveEmp1.moveDate, dbo.TblMoveEmp1.ToDate, dbo.TblMoveEmp1.FroBranchID, TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee,"
sql = sql & "                       dbo.TblMoveEmp1.ToBranchID, TblBranchesData_1.branch_name AS Tobranch_name, TblBranchesData_1.branch_namee AS Tobranch_nameE,"
sql = sql & "                       MONTH(dbo.TblMoveEmp1.moveDate) AS monthname, YEAR(dbo.TblMoveEmp1.moveDate) AS YearID"
sql = sql & "  FROM         dbo.TblEmployee RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblMoveEmp1 ON TblBranchesData_1.branch_id = dbo.TblMoveEmp1.ToBranchID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData TblBranchesData_2 ON dbo.TblMoveEmp1.FroBranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
sql = sql & "                       dbo.mofrdat LEFT OUTER JOIN"
sql = sql & "                       dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.id RIGHT OUTER JOIN"
sql = sql & "                       dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode ON"
sql = sql & "                       dbo.TblMoveEmp1.EmpID = dbo.EmpSalaryComponent.emp_ID ON dbo.TblEmployee.Emp_ID = dbo.TblMoveEmp1.EmpID"
sql = sql & "   Where (year(dbo.TblMoveEmp1.moveDate) = " & val(CboYear.ListIndex + 2006) & ") And (MONTH(dbo.TblMoveEmp1.moveDate) = " & val(CmbMonth.ListIndex + 1) & ")"
sql = sql & " order by dbo.TblMoveEmp1.EmpID , dbo.TblMoveEmp1.ToDate"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With VSFlexGrid2
.Clear flexClearScrollable, flexClearEverything
.rows = 2
.rows = .rows + rs2.RecordCount
For i = .FixedRows To .rows - 2
.TextMatrix(i, .ColIndex("Ser")) = i

.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs2("FroBranchID").value), 0, rs2("FroBranchID").value)
.TextMatrix(i, .ColIndex("ToBranchID")) = IIf(IsNull(rs2("ToBranchID").value), 0, rs2("ToBranchID").value)
.TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value)
.TextMatrix(i, .ColIndex("From")) = IIf(IsNull(rs2("moveDate").value), Date, rs2("moveDate").value)
.TextMatrix(i, .ColIndex("Valuee")) = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)

.TextMatrix(i, .ColIndex("To")) = IIf(IsNull(rs2("ToDate").value), Date, rs2("ToDate").value)
.TextMatrix(i, .ColIndex("NoDay")) = DateDiff("d", .TextMatrix(i, .ColIndex("From")), .TextMatrix(i, .ColIndex("To")))
.TextMatrix(i, .ColIndex("Total")) = Round((val(.TextMatrix(i, .ColIndex("Valuee"))) / 30) * val(.TextMatrix(i, .ColIndex("NoDay"))), SystemOptions.EmpSalaryDigts)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
.TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
.TextMatrix(i, .ColIndex("Tobranch_name")) = IIf(IsNull(rs2("Tobranch_name").value), "", rs2("Tobranch_name").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
Else
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
.TextMatrix(i, .ColIndex("Tobranch_name")) = IIf(IsNull(rs2("Tobranch_nameE").value), "", rs2("Tobranch_nameE").value)
.TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs2("nameE").value), "", rs2("nameE").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub

Private Sub CboPayMentType_Change()

    If Me.CboPayMentType.ListIndex = 0 Then
        DCAccount.Visible = False
        Me.DcboBox.Visible = True
        Me.DcboBankName.Visible = False

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(8).Caption = " «·Œ“Ì‰…"
        Else
            lbl(8).Caption = " Box"
        End If

        Me.TxtChequeNumber.Enabled = False
      
        Me.DtpChequeDueDate.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 2 Or Me.CboPayMentType.ListIndex = 3 Then
        Me.DcboBox.Visible = False
        DCAccount.Visible = False
        Me.DcboBankName.Visible = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(8).Caption = " «·»‰ﬂ"
        Else
            lbl(8).Caption = "Bank"
        End If
    
    ElseIf Me.CboPayMentType.ListIndex = 4 Then
        Me.DcboBox.Visible = False
        Me.DcboBankName.Visible = False
        DCAccount.Visible = True
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(8).Caption = " «·Õ”«»"
        Else
            lbl(8).Caption = "Account"
        End If
    
    Else
        Me.DcboBankName.Visible = False
        Me.DcboBox.Visible = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If
FillGridWithData2
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Code")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Code")) = True

    End If

End Sub

Private Sub Check10_Click()

    If Check10.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("SalesCom")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("SalesCom")) = True

    End If

End Sub

Private Sub Check11_Click()

    If Check11.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("total1")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("total1")) = True

    End If

End Sub

Private Sub Check12_Click()

    If Check12.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = True

    End If

End Sub

Private Sub Check13_Click()

    If Check13.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("TotalDiscount")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("TotalDiscount")) = True

    End If

End Sub

Private Sub Check14_Click()

    If Check14.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("total2")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("total2")) = True

    End If

End Sub

Private Sub Check15_Click()

    If Check15.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("EmpTotalNet")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("EmpTotalNet")) = True

    End If

End Sub

Private Sub Check16_Click()

    If Check16.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("sgn")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("sgn")) = True

    End If

End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.GRID1
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.GRID1

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If

    Me.lbl(14).Caption = Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
 End Sub
Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = True

    End If

End Sub

Private Sub Check3_Click()

    If Check3.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary")) = True

    End If

End Sub

Private Sub check4_Click()

    If Check4.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_sakn")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_sakn")) = True

    End If

End Sub

Private Sub check5_Click()

    If Check5.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_bus")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_bus")) = True

    End If

End Sub

Private Sub check6_Click()

    If Check6.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_food")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_food")) = True

    End If

End Sub

Private Sub check7_Click()

    If Check7.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_others")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_others")) = True

    End If

End Sub

Private Sub Check8_Click()

    If Check8.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("OverTimePrice")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("OverTimePrice")) = True

    End If

End Sub

Private Sub Check9_Click()

    If Check9.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Mokafea")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Mokafea")) = True

    End If

End Sub

Private Sub CheckBox1_Click()
If Me.CheckBox1.value = vbChecked Then
Grid2.ColHidden(Grid2.ColIndex("Emp_Name")) = True
Else
Grid2.ColHidden(Grid2.ColIndex("Emp_Name")) = False
End If
End Sub

Private Sub CmbMonth_Click()
'CmbMonth.Enabled = False
'firstrun = True
    If firstrun = True Then
     
'     If getTitlesName = True Then
   
'   End If
   
        Exit Sub
    End If

  '  CmdOk_Click
  '  CmbMonth.Enabled = True
    'FillGridWithData
End Sub

Private Sub CmbMonth_GotFocus()
    firstrun = False
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
  
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
  
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
    Exit Function
 
End Function

Function ShowGl()
Dim rs As New ADODB.Recordset
Dim My_SQL As String

My_SQL = "  SELECT     dbo.Notes.NoteSerial, dbo.Notes.NoteDate, dbo.Notes.branch_no, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name"
My_SQL = My_SQL + "   FROM         dbo.Notes LEFT OUTER JOIN"
My_SQL = My_SQL + "                        dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
'My_SQL = My_SQL + "   WHERE     (dbo.Notes.salary = '201511') AND (dbo.Notes.NoteType = 555)"
My_SQL = My_SQL + "   WHERE     (dbo.Notes.salary  = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')     AND (dbo.Notes.NoteType = 555) "

        My_SQL = My_SQL + " order by   ( dbo.Notes.NoteSerial ) "
        '  My_SQL = My_SQL + " order by   LPAD(Emp_code,6,'0') ASC"
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.GridGE
            .rows = 2
            .Clear flexClearScrollable



Dim i As Integer

            If rs.RecordCount > 0 Then
                .rows = rs.RecordCount + 1
                rs.MoveFirst

                For i = 1 To .rows - 1
        
                    .TextMatrix(i, .ColIndex("Ser")) = i
            
                    '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
                       "", Rs.Fields("ID").value)
            
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                    .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
            
                    .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs.Fields("NoteDate").value), "", rs.Fields("NoteDate").value)
                    
                  Next i
                  
           End If
           
           End With
            
           
End Function
Private Sub CmdOk_Click()
     On Error Resume Next
     'DcBranch1
            
            
  If SystemOptions.EmployeeSalaryBYBranch = True And dcBranch1.BoundText = "" Then

MsgBox "·«»œ „‰ «Œ Ì«— ›—⁄ ", vbCritical
Exit Sub
End If

'firstrun = False
     If getTitlesName = True Then
   
   End If
    
    
    If firstrun = True Then
 
        Exit Sub
    End If

    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text

    DTP_Date.value = MonthLastDay(CDate(str))
    DTPicker1 = MonthLastDay(CDate(str))
    TxtNoteSerial.text = ""
    Set cProgress = New ClsProgress
    cProgress.ProgressCaption = "xxxxxxx"
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    DoEvents
     FillGridWithData2
    FillGridWithData
    CalculateNets
    Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("EmpTotalNet")) = 0
    Dim SngTotal As Double
              SngTotal = Grid.Aggregate(flexSTSum, Grid.FixedRows, Grid.ColIndex("EmpTotalNet"), Grid.rows - 1, Grid.ColIndex("EmpTotalNet"))
            Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("EmpTotalNet")) = SngTotal
    RetriveProjectSalar
    RetriveEmpMove
    Me.TxtNoteSerial.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 66)
    Me.TxtNoteSerial2.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 555)

    DoEvents
    cProgress.StopProgess
    cProgress.FinishProgress
   
   Dim mRemainDay As Double
   
    Set cProgress = Nothing
    Dim i As Integer
        With Grid
For i = 1 To 40

                If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                  .ColHidden(.ColIndex("Comp" & i)) = True
                End If


                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               
               If val((.TextMatrix(.rows - 1, .ColIndex("VacDay")))) = 0 Then
                  .ColHidden(.ColIndex("VacDay")) = False
                Else
                    .ColHidden(.ColIndex("VacDay")) = False
                
                End If
                
                
              
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalVacValue")))) = 0 Then
                  .ColHidden(.ColIndex("TotalVacValue")) = False
                Else
                    .ColHidden(.ColIndex("TotalVacValue")) = False
                
                End If
                
                   
                
                
                If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation2")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation2")) = True
                End If
                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation3")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation3")) = True
                End If

                                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation4")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation4")) = True
                End If

                
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                End If
                
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If






'
Next i
End With


        With GRID1
For i = 1 To 40

                If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                  .ColHidden(.ColIndex("Comp" & i)) = True
                End If
                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If
Next i
End With

End Sub

Function create_report_data()
    On Error Resume Next
    Dim StrSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim ColumnName As String
    
    
 
'StrSQL = "Delete   emp_salary where m_year='" & CboYear.Text & "' and m_month='" & CmbMonth.Text & "'" ' '& " and Branchid=" & Current_branch

If val(dcBranch1.BoundText) = 0 Then
StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
   Else
    StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and Branchid=" & Current_branch
   End If
   
   
    Cn.Execute StrSQL, , adExecuteNoRecords
 
 
   ' StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'" & " and Branchid=" & Current_branch
   ' Cn.Execute StrSQL, , adExecuteNoRecords

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "emp_salary", Cn, adOpenStatic, adLockOptimistic, adCmdTable


    With Grid

        For i = .FixedRows To .rows - 2
   
            rs.AddNew
       
            rs("BranchId").value = .TextMatrix(i, .ColIndex("BranchId"))

            rs("Emp_ID").value = .TextMatrix(i, .ColIndex("Emp_ID"))
            rs("Emp_Code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
            rs("cost_center_id").value = .TextMatrix(i, .ColIndex("cost_center_id"))
            rs("CountDays").value = val(.TextMatrix(i, .ColIndex("CountDays")))
            rs("AbcentDay").value = val(.TextMatrix(i, .ColIndex("AbcentDay")))
            rs("RemainDay").value = val(.TextMatrix(i, .ColIndex("RemainDay")))
            rs("RemainDay").value = val(.TextMatrix(i, .ColIndex("RemainDay")))
            rs("ToalInsurance").value = val(.TextMatrix(i, .ColIndex("ToalInsurance")))

            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True Then
                    rs(ColumnName).value = val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If
    
            Next j

             rs("Emp_Salary").value = GetSalaryEmployee(val(.TextMatrix(i, .ColIndex("Emp_ID"))))
            ' rs("Emp_Salary_sakn").value = .TextMatrix(i, .ColIndex("Emp_Salary_sakn"))
            ' rs("Emp_Salary_bus").value = .TextMatrix(i, .ColIndex("Emp_Salary_bus"))
            ' rs("Emp_Salary_food").value = .TextMatrix(i, .ColIndex("Emp_Salary_food"))
            ' rs("Emp_Salary_mob").value = .TextMatrix(i, .ColIndex("Emp_Salary_mob"))
            ' rs("Emp_Salary_mang").value = .TextMatrix(i, .ColIndex("Emp_Salary_mang"))
            ' rs("Emp_Salary_others").value = .TextMatrix(i, .ColIndex("Emp_Salary_others"))
            ' rs("OverTimePrice").value = .TextMatrix(i, .ColIndex("OverTimePrice"))
            rs("Mokafea").value = .TextMatrix(i, .ColIndex("Mokafea"))
            rs("TotalAdvance").value = val(.TextMatrix(i, .ColIndex("TotalAdvance")))
            rs("EmpBalance").value = val(.TextMatrix(i, .ColIndex("EmpBalance")))
            rs("EmpReaminB").value = val(.TextMatrix(i, .ColIndex("EmpReaminB")))
            rs("TotalDiscount").value = .TextMatrix(i, .ColIndex("TotalDiscount"))
            rs("SalesCom").value = .TextMatrix(i, .ColIndex("SalesCom"))
            rs("total1").value = val(.TextMatrix(i, .ColIndex("total1")))
            rs("total2").value = val(.TextMatrix(i, .ColIndex("total2")))
            rs("EmpTotalNet").value = val(.TextMatrix(i, .ColIndex("EmpTotalNet")))
            rs("m_year").value = CboYear.text
            rs("m_month").value = CmbMonth.text
            rs("DepartmentID").value = .TextMatrix(i, .ColIndex("dep"))
            rs("project_id").value = val(.TextMatrix(i, .ColIndex("project")))
            rs("sgn").value = CboYear.text & CmbMonth.ListIndex + 1
            rs("RecordDate").value = DTP_Date.value
            rs("LocationID").value = val(.TextMatrix(i, .ColIndex("LocationID")))
            'khaeddone
            rs("VoCation2").value = val(.TextMatrix(i, .ColIndex("VoCation2")))
            rs("VoCation4").value = val(.TextMatrix(i, .ColIndex("VoCation4")))
            rs("VoCation3").value = val(.TextMatrix(i, .ColIndex("VoCation3")))
            
            rs("VoCation3").value = val(.TextMatrix(i, .ColIndex("VoCation3")))
            rs("VacDay").value = val(.TextMatrix(i, .ColIndex("VacDay")))
          '  rs("AbcentDay").value = .TextMatrix(i, .ColIndex("AbcentDay"))
             ',,
            rs("OverTime").value = val(.TextMatrix(i, .ColIndex("OverTime")))
            rs("WorkHours").value = val(.TextMatrix(i, .ColIndex("WorkHours")))
            rs("VoCation").value = val(.TextMatrix(i, .ColIndex("VoCation")))
            rs("TotalVacValue").value = val(.TextMatrix(i, .ColIndex("TotalVacValue")))
            
            
            rs("VoCation3").value = val(.TextMatrix(i, .ColIndex("VoCation3")))
            
            rs.update
   
        Next i

    End With

End Function

Private Sub CmdPrint_Click()
    
    
    
    
    On Error Resume Next
    Dim i As Integer
 



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

Private Sub cmdReCreateJL_Click()
   On Error Resume Next
     'DcBranch1
            
            
  If SystemOptions.EmployeeSalaryBYBranch = True And dcBranch1.BoundText = "" Then

MsgBox "·«»œ „‰ «Œ Ì«— ›—⁄ ", vbCritical
Exit Sub
End If

'firstrun = False
     If getTitlesName = True Then
   
   End If
    
    
    If firstrun = True Then
 
        Exit Sub
    End If

    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text

    DTP_Date.value = MonthLastDay(CDate(str))
    DTPicker1 = MonthLastDay(CDate(str))
    TxtNoteSerial.text = ""
    Set cProgress = New ClsProgress
    cProgress.ProgressCaption = "xxxxxxx"
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    DoEvents
     FillGridWithData2 True
    FillGridWithData True
    CalculateNets
    Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("EmpTotalNet")) = 0
    Dim SngTotal As Double
              SngTotal = Grid.Aggregate(flexSTSum, Grid.FixedRows, Grid.ColIndex("EmpTotalNet"), Grid.rows - 1, Grid.ColIndex("EmpTotalNet"))
            Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("EmpTotalNet")) = SngTotal
    RetriveProjectSalar
    RetriveEmpMove
    Me.TxtNoteSerial.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 66)
    Me.TxtNoteSerial2.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 555)

    DoEvents
    cProgress.StopProgess
    cProgress.FinishProgress
   
   Dim mRemainDay As Double
   
    Set cProgress = Nothing
    Dim i As Integer
        With Grid
For i = 1 To 40

                If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                  .ColHidden(.ColIndex("Comp" & i)) = True
                End If


                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               
               If val((.TextMatrix(.rows - 1, .ColIndex("VacDay")))) = 0 Then
                  .ColHidden(.ColIndex("VacDay")) = False
                Else
                    .ColHidden(.ColIndex("VacDay")) = False
                
                End If
                
                
              
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalVacValue")))) = 0 Then
                  .ColHidden(.ColIndex("TotalVacValue")) = False
                Else
                    .ColHidden(.ColIndex("TotalVacValue")) = False
                
                End If
                
                   
                
                
                If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation2")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation2")) = True
                End If
                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation3")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation3")) = True
                End If

                                          If val((.TextMatrix(.rows - 1, .ColIndex("VoCation4")))) = 0 Then
                  .ColHidden(.ColIndex("VoCation4")) = True
                End If

                
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                End If
                
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If






'
Next i
End With


        With GRID1
For i = 1 To 40

                If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                  .ColHidden(.ColIndex("Comp" & i)) = True
                End If
                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If
Next i
End With
Dim ss As String
 ss = " Delete notes   where salary=" & CboYear.text & CmbMonth.ListIndex + 1
 Cn.Execute ss
ALLButton2_Click

End Sub

Private Sub CmdShow_Click()
ShowGl
End Sub

Private Sub Combo1_Click()

    If Combo1.ListIndex > -1 Then
                         If Combo1.ListIndex = 0 Then
                             ISButton2_Click
                         ElseIf Combo1.ListIndex = 1 Then
                             ISButton2_Click
                         ElseIf Combo1.ListIndex = 2 Then
                             ISButton4_Click
                         ElseIf Combo1.ListIndex = 3 Then
                             ISButton4_Click
                         ElseIf Combo1.ListIndex = 4 Then
                             ISButton6_Click
                         
                        ElseIf Combo1.ListIndex = 5 Then
                             ShowReports (5)
   ElseIf Combo1.ListIndex = 6 Then
                             ShowReports (6)
   ElseIf Combo1.ListIndex = 7 Then
                             ShowReports (7)
   ElseIf Combo1.ListIndex = 8 Then
                             ShowReports (8)
   ElseIf Combo1.ListIndex = 9 Then
                             PrinNomothig 9
   ElseIf Combo1.ListIndex = 10 Then
                             PrinNomothig 10
   ElseIf Combo1.ListIndex = 11 Then
                             PrinNomothig 11
   ElseIf Combo1.ListIndex = 12 Then
                             PrinNomothig 12
   ElseIf Combo1.ListIndex = 13 Then
                             PrinNomothig 13
   ElseIf Combo1.ListIndex = 14 Then
                             PrinNomothig 14
   ElseIf Combo1.ListIndex = 15 Then
                             PrinNomothig 15
   ElseIf Combo1.ListIndex = 16 Then
                             PrinNomothig 16
   ElseIf Combo1.ListIndex = 17 Then
                             PrinNomothig 17
   ElseIf Combo1.ListIndex = 18 Then
                             PrinNomothig 18

                        End If
    End If

End Sub
Sub PrinNomothig(Optional Indx As Integer)

    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

'My_SQL = " SELECT     dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(dbo.TblEmployee.Emp_ID)) AS PrePaidvalue, dbo.TblEmpJobsTypes.JobTypeName AS HJobTypeName, "
'My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee AS HJobTypeNameE, dbo.TblBranchesData.branch_name AS Hbranch_name,"
'My_SQL = My_SQL + "                      dbo.TblBranchesData.branch_namee AS Hbranch_nameH, dbo.TblEmployee.GroupID, dbo.TblEmployee.BankCode, dbo.TblEmployee.BankCard,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.NumEkama, dbo.TblEmployee.ContractID, dbo.TblEmployee.BranchId AS HBranchId, dbo.TblEmployee.Fullcode,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1 AS HEmp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.projects.Project_name, dbo.projects.Project_nameE,"
'My_SQL = My_SQL + "                      dbo.projects.Fullcode AS ProjectFullcode, dbo.TblEmployee.NationalityE, dbo.TblEmployee.BignDateWork, dbo.Contract.Contract_date, dbo.Contract.DateH,"
'My_SQL = My_SQL + "                      dbo.Contract.Contract_Enddate, dbo.Contract.DateH1, dbo.emp_salary.*, dbo.TblEmployee.EmpNotes, dbo.TblEmployee.kafeladd, dbo.TblEmployee.DOB,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.SpecificationID, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmpSpecifications.SpecificationNameE, dbo.TblEmployee.PayType,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.SalaryCode, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.NationlID,"
'My_SQL = My_SQL + "                      dbo.Nationality.name AS Nationalname, dbo.Nationality.namee AS NationalnameE, dbo.TblEmployee.DeptID2,"
'My_SQL = My_SQL + "                      dbo.TblEmpDepartmentsDet.Name AS DepartmentName2, dbo.TblEmpDepartmentsDet.NameE AS DepartmentName2E, dbo.TblEmployee.NoAdded,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.MachinCode, dbo.TblEmployee.TypeEmp, dbo.TblEmployee.ResourceBox, dbo.TblEmployee.HowIqamaEndH, dbo.TblEmployee.HowIqamaStH,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.SafEBox, dbo.TblEmployee.BankIBan, dbo.TblEmployee.BanckName, dbo.TblEmployee.BankIAddress, dbo.TblEmployee.MaritalStatus,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.SalaryInstrunse, dbo.TblEmployee.Sex, dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.DriverLicense,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.lastHolidaydateH,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.VisaNo, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.IssueDateH,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.DOBH, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoket,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.NumPoket, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLinc,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.NumLicn, dbo.TblEmployee.placeEkama, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.InsuranceState,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.pasplace, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.KafelName,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdodno, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.NumPasp,"
'My_SQL = My_SQL + "                      dbo.TblEmployee.DateExpoekamaH , dbo.TblEmployee.KafelID, dbo.TblEmployee.DateEndekamaH, dbo.TblEmployee.placeWORK, dbo.TblEmployee.dean"
'My_SQL = My_SQL + " FROM         dbo.TblEmpSpecifications RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.Nationality RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.TblEmpDepartmentsDet RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.TblEmployee ON dbo.TblEmpDepartmentsDet.ID = dbo.TblEmployee.DeptID2 ON dbo.Nationality.id = dbo.TblEmployee.NationlID ON"
'My_SQL = My_SQL + "                      dbo.TblEmpSpecifications.SpecificationID = dbo.TblEmployee.SpecificationID LEFT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.projects RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.Contract RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.TblEmpDepartments RIGHT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.emp_salary ON dbo.TblEmpDepartments.DeparmentID = dbo.emp_salary.DepartmentID ON dbo.Contract.Emp_id = dbo.emp_salary.emp_id ON"
'My_SQL = My_SQL + "                      dbo.projects.id = dbo.emp_salary.project_id ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
'My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"

    My_SQL = " SELECT VoCation2,VoCation3,VoCation4, dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(TblEmployee.Emp_ID)) AS PrePaidvalue, TblEmpJobsTypes.JobTypeName AS HJobTypeName, TblEmpJobsTypes.JobTypeNamee AS HJobTypeNameE,"
    My_SQL = My_SQL + " TblBranchesData.branch_name AS Hbranch_name, TblBranchesData.branch_namee AS Hbranch_nameH, TblEmployee.GroupID, TblEmployee.BankCode, TblEmployee.BankCard, TblEmployee.NumEkama,"
    My_SQL = My_SQL + " TblEmployee.ContractID, TblEmployee.BranchId AS HBranchId, TblEmployee.Fullcode, TblEmployee.Emp_Name AS HEmp_Name, TblEmployee.Emp_Name1 AS HEmp_Name1, TblEmployee.Emp_Name2,"
    My_SQL = My_SQL + " TblEmployee.Emp_Name3, TblEmployee.Emp_Name4, TblEmployee.Nationality, TblEmployee.Emp_Mail, TblEmployee.Emp_Phone, TblEmployee.Emp_mobile, TblEmployee.Emp_Remark, TblEmployee.Emp_Comm,"
    My_SQL = My_SQL + " TblEmployee.EmpProfitCom, TblEmployee.Emp_Namee, TblEmployee.Emp_Namee1, TblEmployee.Emp_Namee2, TblEmployee.Emp_Namee3, TblEmployee.Emp_Namee4, projects.Project_name, projects.Project_nameE,"
    My_SQL = My_SQL + " projects.Fullcode AS ProjectFullcode, TblEmployee.NationalityE, TblEmployee.BignDateWork, Contract.Contract_date, Contract.DateH, Contract.Contract_Enddate, Contract.DateH1, emp_salary.id, emp_salary.emp_id,"
    My_SQL = My_SQL + " emp_salary.Emp_Code, emp_salary.Emp_Name, emp_salary.Emp_Salary, emp_salary.Emp_Salary_sakn, emp_salary.Emp_Salary_bus, emp_salary.Emp_Salary_food, emp_salary.Emp_Salary_mob,"
    My_SQL = My_SQL + " emp_salary.Emp_Salary_mang, emp_salary.Emp_Salary_others, emp_salary.OverTimePrice, emp_salary.Mokafea, emp_salary.SalesCom, emp_salary.total1, emp_salary.TotalAdvance, emp_salary.TotalDiscount,"
    My_SQL = My_SQL + " emp_salary.total2, emp_salary.EmpTotalNet, emp_salary.sgn, emp_salary.m_year, emp_salary.m_month, emp_salary.payed, emp_salary.DepartmentID, emp_salary.project_id, emp_salary.cost_center_id,"
    My_SQL = My_SQL + " emp_salary.BranchId, emp_salary.Comp1, emp_salary.Comp2, emp_salary.Comp3, emp_salary.Comp4, emp_salary.Comp5, emp_salary.Comp6, emp_salary.Comp7, emp_salary.Comp8, emp_salary.Comp9,"
    My_SQL = My_SQL + " emp_salary.Comp10, emp_salary.Comp11, emp_salary.Comp12, emp_salary.Comp13, emp_salary.Comp14, emp_salary.Comp15, emp_salary.Comp16, emp_salary.Comp17, emp_salary.Comp18, emp_salary.Comp19,"
    My_SQL = My_SQL + " emp_salary.Comp20, emp_salary.Comp21, emp_salary.Comp22, emp_salary.Comp23, emp_salary.Comp24, emp_salary.Comp25, emp_salary.Comp26, emp_salary.Comp27, emp_salary.Comp28, emp_salary.Comp29,"
    My_SQL = My_SQL + " emp_salary.Comp30, emp_salary.Comp31, emp_salary.Comp32, emp_salary.Comp33, emp_salary.Comp34, emp_salary.Comp35, emp_salary.Comp36, emp_salary.Comp37, emp_salary.Comp38, emp_salary.Comp39,"
    My_SQL = My_SQL + " emp_salary.Comp40, emp_salary.CountDays, emp_salary.LockedInterval, emp_salary.RecordDate, emp_salary.ToalInsurance, emp_salary.AbcentDay, emp_salary.RemainDay, emp_salary.VocEntitID, emp_salary.LocationID,"
    My_SQL = My_SQL + " emp_salary.EmpBalance, emp_salary.EmpReaminB, TblEmployee.EmpNotes, TblEmployee.kafeladd, TblEmployee.DOB, TblEmployee.SpecificationID, TblEmpSpecifications.SpecificationName,"
    My_SQL = My_SQL + " TblEmpSpecifications.SpecificationNameE, TblEmployee.PayType, TblEmployee.SalaryCode, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee, TblEmployee.NationlID,"
    My_SQL = My_SQL + " Nationality.name AS Nationalname, Nationality.namee AS NationalnameE, TblEmployee.DeptID2, TblEmpDepartmentsDet.Name AS DepartmentName2, TblEmpDepartmentsDet.NameE AS DepartmentName2E,"
    My_SQL = My_SQL + " TblEmployee.NoAdded, TblEmployee.MachinCode, TblEmployee.TypeEmp, TblEmployee.ResourceBox, TblEmployee.HowIqamaEndH, TblEmployee.HowIqamaStH, TblEmployee.SafEBox, TblEmployee.BankIBan,"
    My_SQL = My_SQL + " TblEmployee.BanckName, TblEmployee.BankIAddress, TblEmployee.MaritalStatus, TblEmployee.SalaryInstrunse, TblEmployee.Sex, TblEmployee.InstanceDateH, TblEmployee.InstanceDateM, TblEmployee.DriverLicense,"
    My_SQL = My_SQL + " TblEmployee.DriverLicenseend, TblEmployee.DriverLicenseStartdH, TblEmployee.DriverLicenseendH, TblEmployee.lastHolidaydateH, TblEmployee.lastHolidaydate, TblEmployee.VisaNo, TblEmployee.LastDateH,"
    My_SQL = My_SQL + " TblEmployee.LastDate, TblEmployee.IssueDateH, TblEmployee.DOBH, TblEmployee.InsuranceNO, TblEmployee.dateendpoketh, TblEmployee.Dateexppoketh, TblEmployee.dateendpoket, TblEmployee.Dateexppoket,"
    My_SQL = My_SQL + " TblEmployee.NumPoket, TblEmployee.DateEndLincH, TblEmployee.DateExpLincH, TblEmployee.DateEndLinc, TblEmployee.DateExpLinc, TblEmployee.NumLicn, TblEmployee.placeEkama, TblEmployee.InsuranceValue,"
    My_SQL = My_SQL + " TblEmployee.InsuranceState, TblEmployee.pasplace, TblEmployee.DateExpoekama, TblEmployee.DateEndekama, TblEmployee.hdomnfaz, TblEmployee.KafelName, TblEmployee.hdoddate, TblEmployee.hdodno,"
    My_SQL = My_SQL + " TblEmployee.DateExpPasp, TblEmployee.DateEndPasp, TblEmployee.NumPasp, TblEmployee.DateExpoekamaH, TblEmployee.KafelID, TblEmployee.DateEndekamah, TblEmployee.placeWORK, TblEmployee.dean,"
    My_SQL = My_SQL + " EmpGroupDep.GroupName , EmpGroupDep.GroupNameE"
    My_SQL = My_SQL + " FROM projects RIGHT OUTER JOIN"
    My_SQL = My_SQL + " Contract RIGHT OUTER JOIN"
    My_SQL = My_SQL + " TblEmpDepartments RIGHT OUTER JOIN"
    My_SQL = My_SQL + " EmpGroupDep RIGHT OUTER JOIN"
    My_SQL = My_SQL + " emp_salary ON EmpGroupDep.GroupID = emp_salary.LocationID ON TblEmpDepartments.DeparmentID = emp_salary.DepartmentID ON Contract.Emp_id = emp_salary.emp_id ON"
    My_SQL = My_SQL + " projects.id = emp_salary.project_id LEFT OUTER JOIN"
    My_SQL = My_SQL + " TblEmpSpecifications RIGHT OUTER JOIN"
    My_SQL = My_SQL + " Nationality RIGHT OUTER JOIN"
    My_SQL = My_SQL + " TblEmpDepartmentsDet RIGHT OUTER JOIN"
    My_SQL = My_SQL + " TblEmployee ON TblEmpDepartmentsDet.ID = TblEmployee.DeptID2 ON Nationality.id = TblEmployee.NationlID ON TblEmpSpecifications.SpecificationID = TblEmployee.SpecificationID LEFT OUTER JOIN"
    My_SQL = My_SQL + " TblEmpJobsTypes ON TblEmployee.JobTypeID = TblEmpJobsTypes.JobTypeID ON emp_salary.emp_id = TblEmployee.Emp_ID LEFT OUTER JOIN"
    My_SQL = My_SQL + " TblBranchesData ON emp_salary.BranchId = TblBranchesData.branch_id"
    My_SQL = My_SQL + " WHERE  (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  "


    If Dcdep.BoundText <> "" And Dcdep.text <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and emp_salary.DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DcEmp.BoundText <> "" And Me.DcEmp.text <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_salary.emp_id=" & val(Me.DcEmp.BoundText)
    End If
    If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
          My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex) & " "
    End If
      If DcbDepartment2.text <> "" And val(DcbDepartment2.BoundText) <> 0 Then
          My_SQL = My_SQL + " and dbo.TblEmployee.DeptID2=" & val(DcbDepartment2.BoundText) & " "
    End If
    
    If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "
    End If
   If Me.DcbTeam.BoundText <> "" And Me.DcbTeam.text <> "" Then
        My_SQL = My_SQL + "  and TblEmployee.SpecificationID=" & val(Me.DcbTeam.BoundText)
    End If
    
       If Me.dcempcontract.BoundText <> "" And Me.dcempcontract.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.ContractID=" & val(Me.dcempcontract.BoundText)
    End If
 
        If Me.dcBranch1.BoundText <> "" And Me.dcBranch1.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.BranchId=" & val(Me.dcBranch1.BoundText)
    End If
 
 
        If Me.DCGroupID.BoundText <> "" And Me.DCGroupID.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.GroupID=" & val(Me.DCGroupID.BoundText)
    End If
  
  My_SQL = My_SQL + "  order by TblEmployee.Fullcode"
 '
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim StrFileName As String
    Select Case Indx
    Case 9
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT101.rpt"
    Case 10
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT102.rpt"
    Case 11
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT103.rpt"
    Case 12
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT104.rpt"
    Case 13
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT105.rpt"
    Case 14
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT107.rpt"
    Case 15
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT108.rpt"
    Case 16
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT109.rpt"
    Case 17
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT1010.rpt"
    Case 18
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT1011.rpt"
    
    End Select
        If Dir(StrFileName) = "" Then
     MsgBox "«·„·› €Ì— „ÊÃÊœ ", vbCritical
     Exit Sub
    End If
    
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Dim str As String
    
    If dcBranch1.text <> "" Then
    str = "«·›—⁄ : " & dcBranch1.text & CHR(13)
    End If
     
        If DCGroupID.text <> "" Then
    str = str & CHR(13) & "«·„Êﬁ⁄ : " & DCGroupID.text & CHR(13)
    End If
      
        If dcproject.text <> "" Then
    str = str & CHR(13) & "«·„‘—Ê⁄ : " & dcproject.text & CHR(13)
    End If
            
           If Dcdep.text <> "" Then
    str = str & CHR(13) & "«·ﬁ”„ : " & Dcdep.text & CHR(13)
    End If
      
     
           If dcempcontract.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «· ⁄«ﬁœ : " & dcempcontract.text & CHR(13)
    End If
           
           If DcEmp.text <> "" Then
    str = str & CHR(13) & "«·„ÊŸ› : " & DcEmp.text & CHR(13)
    End If
     If cboPayType.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «·”œ«œ : " & cboPayType.text & CHR(13)
    End If
     If DcbHemiaSalary.text <> "" Then
    str = str & CHR(13) & "ﬂÊœ Õ„«Ì… «·«ÃÊ— : " & DcbHemiaSalary.text & CHR(13)
    End If
           
    xReport.ParameterFields(6).AddCurrentValue str
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.dcproject.BoundText <> "" Then
            '   xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & dcproject.text
            Else
            '   xReport.ParameterFields(48).AddCurrentValue "  " & dcproject.text
            End If
       xReport.ParameterFields(48).AddCurrentValue "  " '& dcproject.text
       
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = StrFileName
    FrmReport.CRViewer.viewReport
       FrmReport.TXTSTRSQL = My_SQL
       CreateLogo xReport, val(dcBranch1.BoundText)
 
      
    ' FrmReport.Show
  
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
     Sendkeys "{RIGHT}"
End Sub

Function ShowReports(indexs As Integer)
Dim FileName As String

Select Case indexs
Case 5
    FileName = App.path & "\reports\emp\REPORT10project.rpt"
Case 6
    
    If SystemOptions.UserInterface = EnglishInterface Then
        FileName = App.path & "\reports\emp\REPORT10empE.rpt"
    Else
        FileName = App.path & "\reports\emp\REPORT10emp.rpt"
    End If
Case 7
   ' FileName = App.path & "\reports\emp\REPORT106.rpt"
FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT106.rpt"
Case 8
   ' FileName = App.path & "\reports\emp\REPORT106.rpt"
FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10Team.rpt"

End Select


 

    'FillGridWithData

    'DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report
   ' My_SQL = " SELECT     *"
   ' My_SQL = My_SQL & " FROM         dbo.emp_salary INNER JOIN"
   ' My_SQL = My_SQL & "  dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
   ' My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"


My_SQL = " SELECT     dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblBranchesData.*, dbo.emp_salary.*, dbo.projects.Fullcode, "
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1 AS HEmp_Name1, dbo.TblEmployee.Emp_Name2 AS HEmp_Name2,"
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Name3 AS HEmp_Name3, dbo.TblEmployee.Emp_Name4 AS HEmp_Name4, dbo.TblEmployee.Nationality AS HNationality,"
My_SQL = My_SQL & "                      dbo.TblEmployee.Fullcode AS HFullcode, dbo.TblEmployee.Emp_Namee4 AS HEmp_NameE4, dbo.TblEmployee.Emp_Namee3 AS HEmp_NameE3,"
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Namee2 AS HEmp_NameE2, dbo.TblEmployee.Emp_Namee1 AS HEmp_NameE1, dbo.TblEmployee.Emp_Namee AS HEmp_NameE,"
My_SQL = My_SQL & "                      dbo.TblEmployee.NumEkama AS HNumEkama, dbo.TblEmployee.BankCode AS HBankCode, dbo.TblEmployee.BankCard AS HBankCard,"
My_SQL = My_SQL & "                      dbo.TblEmployee.NationalityE AS HNationalityE, dbo.TblEmployee.kafeladd, dbo.TblEmployee.EmpNotes, dbo.TblEmpSpecifications.SpecificationName,"
My_SQL = My_SQL & "                      dbo.TblEmpSpecifications.SpecificationNameE, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
My_SQL = My_SQL & "                      dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.PayType, dbo.TblEmployee.SalaryCode, dbo.TblEmployee.DeptID2,"
My_SQL = My_SQL & "                      dbo.TblEmpDepartmentsDet.Name AS DepartmentName2, dbo.TblEmpDepartmentsDet.NameE AS DepartmentName2E, dbo.TblEmployee.NationlID,"
My_SQL = My_SQL & "                      dbo.Nationality.name AS Nationalname, dbo.Nationality.namee AS NationalnameE, dbo.TblEmployee.dean, dbo.TblEmployee.DateEndekamah,"
My_SQL = My_SQL & "                      dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.DateExpPasp,"
My_SQL = My_SQL & "                      dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate, dbo.TblEmployee.KafelName, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.Emp_Mail,"
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace,"
My_SQL = My_SQL & "                      dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc,"
My_SQL = My_SQL & "                      dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket,"
My_SQL = My_SQL & "                      dbo.TblEmployee.dateendpoket, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.DOB, dbo.TblEmployee.kafeltel, dbo.TblEmployee.Dateexppoketh,"
My_SQL = My_SQL & "                      dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.DOBH, dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate,"
My_SQL = My_SQL & "                      dbo.TblEmployee.LastDateH, dbo.TblEmployee.VisaNo, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.DriverLicense,"
My_SQL = My_SQL & "                      dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.InstanceDateM,"
My_SQL = My_SQL & "                      dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.Sex, dbo.TblEmployee.MaritalStatus, dbo.TblEmployee.BankIAddress, dbo.TblEmployee.BanckName,"
My_SQL = My_SQL & "                      dbo.TblEmployee.BankIBan, dbo.TblEmployee.SafEBox, dbo.TblEmployee.HowIqamaStH, dbo.TblEmployee.HowIqamaEndH, dbo.TblEmployee.ResourceBox,"
My_SQL = My_SQL & "                      dbo.TblEmployee.TypeEmp , dbo.TblEmployee.MachinCode, dbo.TblEmployee.NoAdded"
My_SQL = My_SQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.Nationality.id = dbo.TblEmployee.NationlID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmpSpecifications ON dbo.TblEmployee.SpecificationID = dbo.TblEmpSpecifications.SpecificationID RIGHT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.emp_salary INNER JOIN"
My_SQL = My_SQL & "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.projects ON dbo.emp_salary.project_id = dbo.projects.id"
My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"

    If Dcdep.BoundText <> "" Then
        'My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and dbo.TblEmployee.DepartmentID=" & val(Dcdep.BoundText)
        'Else
        'My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        'My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
        My_SQL = My_SQL + " and PayType=" & val(cboPayType.ListIndex)
    End If
 
    If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "
    End If
    
    If Me.DcEmp.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and TblEmployee.emp_id=" & val(Me.DcEmp.BoundText)
    End If
    
    If Me.DcbTeam.BoundText <> "" Then
        My_SQL = My_SQL + "  and dbo.TblEmployee.SpecificationID=" & val(Me.DcbTeam.BoundText)
    End If

    If Me.dcproject.BoundText <> "" Then
        'My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and project_id=" & val(Me.dcproject.BoundText)
    End If
    
    If val(dcBranch1.BoundText) <> 0 Then
        My_SQL = My_SQL + "and dbo.TblEmployee.BranchId =" & val(Me.dcBranch1.BoundText)
    End If
    
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
        If Dir(FileName) = "" Then
     MsgBox "«·„·› €Ì— „ÊÃÊœ ", vbCritical
     Exit Function
    End If
    
    Set xReport = xApp.OpenReport(FileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
    
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.dcproject.BoundText <> "" Then
               xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & dcproject.text
            Else
               xReport.ParameterFields(48).AddCurrentValue "  " & dcproject.text
            End If
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
   'Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
Select Case indexs
Case 5
    FrmReport.txtPath = FileName
Case 6
    FrmReport.txtPath = FileName
Case 7
    FrmReport.txtPath = FileName
Case 8
    FrmReport.txtPath = FileName

End Select
    FrmReport.CRViewer.viewReport
    FrmReport.txtPath = My_SQL
    ' FrmReport.Show
  
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
  Sendkeys "{RIGHT}"

End Function
Function CheckClosed(Optional sgn1 As String) As Boolean
Dim Rs11 As ADODB.Recordset
Set Rs11 = New ADODB.Recordset
Dim sql As String
CheckClosed = False
sql = "SELECT     LockedInterval, sgn"
sql = sql & " From dbo.emp_salary"
sql = sql & " WHERE     (LockedInterval = 1) AND (sgn = '" & sgn1 & "')"
Rs11.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs11.RecordCount > 0 Then
CheckClosed = True
Else
CheckClosed = False
End If
End Function
Private Sub Command1_Click()
          If SystemOptions.EmployeeSalaryBYBranch = True Then
            Me.dcBranch1.BoundText = Current_branch
            Me.dcBranch1.Enabled = False
            Else
            Me.dcBranch1.Enabled = True
            End If
            
    Dim X As Integer
    Dim Msg As String
    Dim StrSQL  As String
If CheckClosed(CboYear.text & CmbMonth.ListIndex + 1) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·› —… „ﬁ›·…"
Else
MsgBox "Can Not Delete This Period Closed"
End If
Exit Sub
End If
    If SystemOptions.LockSalary = True Then
    If check_Lock_Salary(CboYear.text, CmbMonth.ListIndex + 1) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·„”Ì— „ﬁ›· ·«Ì„ﬂ‰  ⁄œÌ·Â «Ê Õ–›Â", vbCritical
        Else
            MsgBox "JV Alraedy Locked", vbCritical
        End If

        Exit Sub
   End If
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ Ê«·”œ«œ ·Â–« «·‘Â— "
    Else
        Msg = "Confirm Delete This Month Vouchers "
    End If

    Msg = Msg + CboYear.text & " /" & CmbMonth.ListIndex + 1
    X = MsgBox(Msg, vbCritical + vbYesNo)

    If X = vbYes Then

        StrSQL = "Delete  marakes_taklefa_temp  where kedno=" & get_notes_foxy_no(CboYear.text & CmbMonth.ListIndex + 1, "foxy_no")
        Cn.Execute StrSQL, , adExecuteNoRecords

       ' StrSQL = "Delete  Notes  where salary=" & CboYear.text & CmbMonth.ListIndex + 1 & " and Branch_no=" & Current_branch
       ' Cn.Execute StrSQL, , adExecuteNoRecords
       '
        
       ' StrSQL = "Delete   emp_salary where SGN='" & CboYear.text & CmbMonth.ListIndex + 1 & "'" & " and BranchId=" & Current_branch
       ' Cn.Execute StrSQL, , adExecuteNoRecords


UnUpdatePrePaid
UnUpdateAdvanc
If val(dcBranch1.BoundText) = 0 Then
   StrSQL = "Delete  Notes  where salary=" & CboYear.text & CmbMonth.ListIndex + 1 '& " and Branch_no=" & Current_branch
   Else
   StrSQL = "Delete  Notes  where salary=" & CboYear.text & CmbMonth.ListIndex + 1 & " And branch_no = " & val(dcBranch1.BoundText)
   End If
   
        Cn.Execute StrSQL, , adExecuteNoRecords
       
        ' StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
     If val(dcBranch1.BoundText) = 0 Then
        StrSQL = "Delete   emp_salary where SGN='" & CboYear.text & CmbMonth.ListIndex + 1 & "'" '& " and BranchId=" & Current_branch
    Else
    StrSQL = "Delete   emp_salary where SGN='" & CboYear.text & CmbMonth.ListIndex + 1 & "'  and BranchId=" & val(dcBranch1.BoundText)
    End If
    
        Cn.Execute StrSQL, , adExecuteNoRecords
        UnUpdateVoCationPayed
        UpdateInsurancePayed
        
        With Me.GRID1
            .rows = 2
            .Clear flexClearScrollable
        End With

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ Ê«·”œ«œ ·Â–« «·‘Â— "
        Else
            Msg = " this voucher deleted for "
        End If

        Msg = Msg + CboYear.text & " /" & CmbMonth.ListIndex + 1
        X = MsgBox(Msg, vbCritical)

    End If
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 2
End With
With VSFlexGrid2
.Clear flexClearScrollable, flexClearEverything
.rows = 2
End With
    LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & CHR(13) & " «·‘Â—     " & CmbMonth.text & CHR(13) & "  «·”‰…   " & CboYear.text & CHR(13) & " «· «—ÌŒ " & DTP_Date.value
                     
    LogTexte = ""
       AddToLogFile CInt(user_id), 66, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtNoteSerial), ""
       
 
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\Payrolll.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
    'Me.Grid.saveGrid StrFileName, flexFileExcel, True
  '  Me.Grid.saveGrid StrFileName, flexFileCustomText, True
    
  '  OpenFile StrFileName
  
      On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Payroll"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.Grid.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
    
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\PayrollBankl.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
    Dim i As Integer
 '            With GRID2
'
'        For i = .FixedRows To .Rows - 2
'
'
'                If (GRID2.TextMatrix(i, GRID2.ColIndex("Emp_Namee"))) = "" Then
'                GRID2.RemoveItem i
'
'                GRID2.Refresh
'            DoEvents
'                End If
'        Next i
'
'    End With
    
'Grid.RightToLeft = True
  '  Me.GRID2.saveGrid StrFileName, flexFileCustomText, True
  '  OpenFile StrFileName
    
    
    
        On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Bank"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.Grid2.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
    
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = "C:\" & "\PayrollPayments.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
'    Me.Grid1.saveGrid StrFileName, flexFileCustomText, True
'    OpenFile StrFileName

        On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "ParrollPay"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.GRID1.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
End Sub

Private Sub Command5_Click()
print_report
End Sub



Private Sub Command7_Click()
If TxtNoteSerial.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «‰‘«¡ «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Create JV of this Month", vbCritical
End If
Exit Sub
End If
 
If SystemOptions.LockSalary = True Then
UpdateLockSalary CboYear.text, CmbMonth.ListIndex + 1
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «ﬁ›«· «·„”Ì—"
Else
MsgBox "Salary Has Been Locked"
End If

End If
End Sub

Private Sub DCAccount_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 192
            
    End If

End Sub

Private Sub Dcedara_Click(Area As Integer)
    'CmdOk_Click
End Sub

 

Private Sub DcBranch1_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcBranch1.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
End Sub

Private Sub DcbTeam_KeyUp(KeyCode As Integer, Shift As Integer)
    If DcbTeam.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
End Sub

Private Sub Dcdep_Change()
LoadDept
End Sub

Private Sub Dcdep_Click(Area As Integer)
Dcdep_Change
End Sub

Private Sub dcdep_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If Dcdep.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If

End Sub

Private Sub dcEmp_Change()
       If val(DcEmp.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcEmp.BoundText, EmpCode
    TxtSearchCode.text = EmpCode


End Sub

Private Sub Dcemp_Click(Area As Integer)
dcEmp_Change
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyReturn Then
        CmdOk_Click
    End If
    

    If KeyCode = vbKeyF3 Then
      
        FrmEmployeeSearch.lbltype = 16118
      
        'Set FrmEmployeeSearch.RetrunFrm = Me

      FrmEmployeeSearch.show

    End If
End Sub

Private Sub dcempcontract_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcempcontract.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
End Sub

Private Sub DCGroupID_KeyUp(KeyCode As Integer, Shift As Integer)
    If DCGroupID.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
    
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

    With GRID1
     
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

Private Sub dcproject_Click(Area As Integer)
' CmdOk_Click
    
    
End Sub

Function CheckAccounts() As Boolean
CheckAccounts = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
            showMofradAll(i) = IIf(IsNull(rs("showMofradAll").value), False, rs("showMofradAll").value)
            culc30orRminder(i) = IIf(IsNull(rs("culc30orRminder").value), 0, rs("culc30orRminder").value)
            showinMosirVac(i) = IIf(IsNull(rs("showinMosirVac").value), False, rs("showinMosirVac").value)
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            
    
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
              
            If ViewComp(i) = True And Account_code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
            MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
            CheckAccounts = False
          
           ' Unload Me
              Exit Function
            End If
          
             
              
         If SystemOptions.ProjectEmployeeGV = True And SystemOptions.ProjectDiscountPolicy = 1 Then 'xxx
                  If ViewComp(i) = True And AddOrDiscount(i) = -1 And Account_code1(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
                MsgBox " ·„ Ì „ —»ÿ Õ”«» «·«Ì—«œ«  «· Ì  ⁄·Ì «·Œ’„ «·Œ«’ » " & componentname(i), vbCritical
        '        CheckAccounts = False
                
                '  Unload Me
                    Exit Function
                  End If
              
             End If
             
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function
Sub FillTable()
Dim i As Integer
Dim j As Integer
Dim ColumnName As String
Dim Yar As Double
Dim Manth As Double
Dim StrSQL As String
Yar = val(Me.CboYear.text)
Manth = (Me.CmbMonth.ListIndex + 1)

                 StrSQL = "Delete From TblEmpMofrd Where Yar=" & Yar & " and Manth =" & Manth & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        With Me.Grid
          For i = .FixedRows To .rows - 1
          If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 Then
          For j = 1 To 40
           ColumnName = "Comp" & j
           
     '      Debug.Print ColumnName & " ----  " & val(.TextMatrix(i, .ColIndex(ColumnName))) & " --- " & val(AddOrDiscount(j))
          If val(.TextMatrix(i, .ColIndex(ColumnName))) <> 0 Then
          If val(AddOrDiscount(j)) = -1 Then
          SaveTable .TextMatrix(0, .ColIndex(ColumnName)), val(.TextMatrix(i, .ColIndex("Emp_ID"))), val(.TextMatrix(i, .ColIndex("BranchId"))), val(.TextMatrix(i, .ColIndex(ColumnName))) * -1
          Else
          SaveTable .TextMatrix(0, .ColIndex(ColumnName)), val(.TextMatrix(i, .ColIndex("Emp_ID"))), val(.TextMatrix(i, .ColIndex("BranchId"))), val(.TextMatrix(i, .ColIndex(ColumnName)))
          End If
          End If
          Next j
          If val(.TextMatrix(i, .ColIndex("PrePaidvalue"))) <> 0 Then
          SaveTable .TextMatrix(0, .ColIndex("PrePaidvalue")), val(.TextMatrix(i, .ColIndex("Emp_ID"))), val(.TextMatrix(i, .ColIndex("BranchId"))), val(.TextMatrix(i, .ColIndex("PrePaidvalue"))) * -1
          End If
            If val(.TextMatrix(i, .ColIndex("ToalInsurance"))) <> 0 Then
          SaveTable .TextMatrix(0, .ColIndex("ToalInsurance")), val(.TextMatrix(i, .ColIndex("Emp_ID"))), val(.TextMatrix(i, .ColIndex("BranchId"))), val(.TextMatrix(i, .ColIndex("ToalInsurance"))) * -1
          End If
             If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) <> 0 Then
          SaveTable .TextMatrix(0, .ColIndex("TotalAdvance")), val(.TextMatrix(i, .ColIndex("Emp_ID"))), val(.TextMatrix(i, .ColIndex("BranchId"))), val(.TextMatrix(i, .ColIndex("TotalAdvance"))) * -1
          End If

          End If
          Next i
         End With
End Sub
Sub SaveTable(Optional MofrdName As String = "", Optional EmpID As Double = 0, Optional BranchID As Double = 0, Optional Valu As Double = 0)
Dim StrSQL As String
Dim RsDetails As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
 StrSQL = "SELECT  * from TblEmpMofrd Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   RsDetails.AddNew
   RsDetails("MofrdName").value = MofrdName
   RsDetails("EmpID").value = EmpID
   RsDetails("BrnchID").value = BranchID
   RsDetails("Valu").value = Valu
   RsDetails("RecorDate").value = DTP_Date.value
   RsDetails("Yar").value = val(Me.CboYear.text)
   RsDetails("Manth").value = val((Me.CmbMonth.ListIndex + 1))
   RsDetails.update
End Sub

Function getTitlesName() As Boolean
Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
getTitlesName = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
             Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
             MofrdAbcen(i) = IIf(IsNull(rs("MofrdAbcen").value), False, rs("MofrdAbcen").value)
            
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            

            
            
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
             
         '   If ViewComp(i) = True And Account_Code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
         '   MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
         '   getTitlesName = False
          
           ' Unload Me
         '     Exit Function
         '   End If
              
              
            With Me.Grid
             
                ColumnName = "Comp" & i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                     
                If ViewComp(i) = True Then
                    .ColHidden(.ColIndex(ColumnName)) = False
                Else
                    .ColHidden(.ColIndex(ColumnName)) = True
                End If
                     
            End With
             
            With Me.GRID1
                ColumnName = "Comp" & i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                     
                If ViewComp(i) = True Then
                    .ColHidden(.ColIndex(ColumnName)) = False
                Else
                    .ColHidden(.ColIndex(ColumnName)) = True
                End If

            End With
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
   If dcproject.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
    
    
End Sub

Private Sub DTP_Date_Change()
    TxtNoteSerial.text = ""
End Sub

  Sub LoadDept()
     Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcbDepartment2
     If val(Dcdep.BoundText) <> 0 Then
     Dcombos.GetNewDwpartMent DcbDepartment2, True, val(Dcdep.BoundText)
     Else
     Dcombos.GetNewDwpartMent DcbDepartment2
    End If
  End Sub
Private Sub Form_Load()
            If SystemOptions.EmployeeSalaryBYBranch = True Then
            Me.dcBranch1.BoundText = Current_branch
            Me.dcBranch1.Enabled = False
            Else
            Me.dcBranch1.Enabled = True
            End If
            
            
                  If SystemOptions.SpecialVersion = True Then
     Frame1.Visible = False
End If

With Me.Grid

 If SystemOptions.ShowBalanceOfEmpInSalary = True Then
 .ColHidden(.ColIndex("EmpBalance")) = False
 .ColHidden(.ColIndex("EmpReaminB")) = False
 Else
 .ColHidden(.ColIndex("EmpBalance")) = True
 .ColHidden(.ColIndex("EmpReaminB")) = True
 End If

End With


    Dim My_SQL As String
 C1Tab1.CurrTab = 0
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Command7.Visible = False
    firstrun = True
RenameCompoPrint
    DTP_Date.value = Date
  If SystemOptions.LockSalary = True Then
  Command7.Visible = True
  End If
    My_SQL = "select Emp_id,Emp_Name From TblEmployee  order by  Emp_Name"
    fill_combo DcEmp, My_SQL

    My_SQL = "select DeparmentID,DepartmentName From TblEmpDepartments  order by DepartmentName "
    fill_combo Dcdep, My_SQL

    My_SQL = " select id,Project_name from projects order by Project_name"
    fill_combo dcproject, My_SQL

    My_SQL = "SELECT  (branch_id) From TblBranchesData"
   
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If
    My_SQL = "SELECT  (DeparmentID),DepartmentName ,DepartmentNamee From TblEmpDepartments where DeparmentID in(select DepartmentID  from TblEmployee)"
    RsDepartment.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDepartment.RecordCount > 0 Then
        RsDepartment.MoveFirst
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    With cboPayType
.Clear
.AddItem "‰ﬁœ«"
.AddItem "‘Ìﬂ"
.AddItem "’—«›"
.AddItem " ÕÊÌ· »‰ﬂÌ"
.AddItem "«Œ—Ì"
End With
Else
With cboPayType
.Clear
.AddItem "Cash"
.AddItem "Cheque"
.AddItem "ATM"
.AddItem "Transfer"
.AddItem "Others"

End With
End If
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCmboEmp, True
    Set cSearchDCombo = New clsDCboSearch
    Set cSearchDCombo.Client = DCmboEmp
    Dcombos.GetEmpSpecifications Me.DcbTeam
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.dcBranch1
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetAccountingCodes Me.DCAccount
    Dcombos.GetEmpLocations Me.DCGroupID
    Dcombos.GetEmpSalaryCode Me.DcbHemiaSalary
    Dcombos.GetNewDwpartMent DcbDepartment2
     Dcombos.Getemp_Contract_type Me.dcempcontract
      If SystemOptions.Allowpayroll = True Then
   ALLButton2.Enabled = True
  Command1.Enabled = True

       End If


    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        '    .Rows = 1
        '    .ExplorerBar = flexExSortShowAndMove
        '    .RowHeightMin = 300
        '    .ExtendLastCol = True
        '    .WallPaper = BKGrndPic.Picture
        '  .AutoSize 0, .Cols - 1, False
    End With

    With Me.Grid
        .rows = 1
        .Clear flexClearScrollable
    End With

    With Me.GRID1
        .rows = 1
        .Clear flexClearScrollable
    End With

    Me.C1Tab1.TabVisible(1) = False
    'SetDtpickerDate Me.DtpFrom
    'SetDtpickerDate Me.DtpTO

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'SHow_grig_col

    ' GetMySetting
     
   If getTitlesName = True Then
   
   End If
   

     If CheckAccounts = False And detect_employee_work_type = 1 Then
    ALLButton2.Enabled = False
'    Exit Sub
    
    End If

    YearMonth

    If detect_employee_work_type = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            ALLButton2.Caption = "«‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
            ALLButton2.Caption = "  Create JE Voucher  "
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            ALLButton2.Caption = "«‰‘«¡   ”‰œ «·—« »"
        Else
            ALLButton2.Caption = "  Create  Salary Doc  "
        End If
    End If
 
    'Resize_Form Me, True
End Sub

Function ChangeLang()


 TranslateForm Me, True
Command7.Caption = "Lock Salary"
CheckBox1.RightToLeft = False
CheckBox1.Caption = "Hide Name Arabic"
    lbl(11).Caption = "Date"
    lbl(12).Caption = "Date"
 lbl(16).Caption = "Branch"
  lbl(15).Caption = "Location"
   lbl(17).Caption = "Con. Type"
   cmdOK.Caption = "View"
   lbl(13).Caption = "Total"
   Command3.Caption = "Exprt"
   Command2.Caption = "Export"
   Command4.Caption = "Export"
   CMDShow.Caption = "Show GE"
   lbl(20).Caption = "Wag.Pro.Code"
   lbl(19).Caption = "Payment Type"
   lbl(18).Caption = "Team"
  Label3.Caption = "Select Criteria and press Enter"
  
    Me.Caption = "Monthly Payroll"
    ALLButton2.Caption = "Salary Allocation JV"
    ALLButton3.Caption = "Salary Payment JV"
    Me.C1Tab1.TabCaption(0) = "Salary Allocation "
    Me.C1Tab1.TabCaption(2) = "Salary Payment"
    Me.C1Tab1.TabCaption(3) = "Pay GL"
    Me.C1Tab1.TabCaption(4) = "Project Details"
    Me.C1Tab1.TabCaption(5) = "Transportation details"

    Ele(3).Caption = "Select Date"
    lbl(0).Caption = "Month"
    lbl(2).Caption = "Year"
    Fra.Caption = "Work Hours"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Emp Name"
    lbl(4).Caption = "Management"
    lbl(21).Caption = "Departement"
    lbl(5).Caption = "Project"
    lbl(6).Caption = "Select Report"
    lbl(7).Caption = "Payment Type"
    lbl(8).Caption = "Box"
    lbl(9).Caption = "Cheque No."
    lbl(10).Caption = "Due Date"

    ALLButton1.Caption = "Change Screen"
    CmdPrint.Caption = "Print"
    CmdExit.Caption = "Exit"
    Command1.Caption = "Delete JL"

    Check17.Caption = "Select All"
Label4.Caption = "Press Enter on any row to show Employee File"
Label5.Caption = "Press Enter on any row to show Employee File"
With VSFlexGrid2
            .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
            .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
            .TextMatrix(0, .ColIndex("Ser")) = "Serial"
            .TextMatrix(0, .ColIndex("branch_name")) = "From Branch"
            .TextMatrix(0, .ColIndex("mofrad_name")) = "Component"
            .TextMatrix(0, .ColIndex("Valuee")) = "Value"
            .TextMatrix(0, .ColIndex("From")) = "From"
            .TextMatrix(0, .ColIndex("To")) = "To"
            .TextMatrix(0, .ColIndex("NoDay")) = "No.Days"
            .TextMatrix(0, .ColIndex("Total")) = "Total"
            .TextMatrix(0, .ColIndex("Tobranch_name")) = "To Barnch"
            
End With
With VSFlexGrid1
            .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
            .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
            .TextMatrix(0, .ColIndex("Ser")) = "Serial"
            .TextMatrix(0, .ColIndex("Project_name")) = "Project"
            .TextMatrix(0, .ColIndex("mofrad_name")) = "Component"
            .TextMatrix(0, .ColIndex("Valuee")) = "Value"
            .TextMatrix(0, .ColIndex("NoDay")) = "No.Days"
            .TextMatrix(0, .ColIndex("Total")) = "Total"
            .TextMatrix(0, .ColIndex("TypeSalary")) = "Type Salary"
            
End With

With GridGE
            .TextMatrix(0, .ColIndex("branch_name")) = "Branch_name"
            .TextMatrix(0, .ColIndex("NoteDate")) = "NoteDate"
            .TextMatrix(0, .ColIndex("NoteSerial")) = "NoteSerial"
            .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
            
End With
    With Me.Grid
            .TextMatrix(0, .ColIndex("Nationality")) = "Nationality"
                    .TextMatrix(0, .ColIndex("JobTypeName")) = "Job Name"
                            .TextMatrix(0, .ColIndex("BignDateWork")) = "BignDateWork"
                                    .TextMatrix(0, .ColIndex("lastHolidaydate")) = "Last Holiday date"
                                            .TextMatrix(0, .ColIndex("CountDays")) = "Work Days"
                                                    .TextMatrix(0, .ColIndex("ToalInsurance")) = "ToalInsurance"
                                                    .TextMatrix(0, .ColIndex("VoCation")) = "Holiday"
                                                            .TextMatrix(0, .ColIndex("PrePaidvalue")) = "PrePaid Value"
                                                            
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("AbcentDay")) = "Days Absence"
        .TextMatrix(0, .ColIndex("RemainDay")) = "Remaining Days"
        
        .TextMatrix(0, .ColIndex("Emp_id")) = "Emp.ID"
        .TextMatrix(0, .ColIndex("emp_code")) = "Emp.Code"
        .TextMatrix(0, .ColIndex("emp_name")) = "Emp.Name"
        .TextMatrix(0, .ColIndex("Mokafea")) = "Additional"
        .TextMatrix(0, .ColIndex("TotalAdvance")) = "Advances"
        .TextMatrix(0, .ColIndex("TotalDiscount")) = "Discounts"
        .TextMatrix(0, .ColIndex("SalesCom")) = "Sales Com."
        .TextMatrix(0, .ColIndex("EmpTotalNet")) = "Net "
        .TextMatrix(0, .ColIndex("sgn")) = "sgn"
        .TextMatrix(0, .ColIndex("total1")) = "Total Add. "
        .TextMatrix(0, .ColIndex("total2")) = "Total Discount. "
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
        .TextMatrix(0, .ColIndex("Location")) = "Location"
        .ColHidden(.ColIndex("dep")) = True
        .ColHidden(.ColIndex("Branchid")) = True
        .ColHidden(.ColIndex("branchname")) = True
        .ColHidden(.ColIndex("project")) = True
        .ColHidden(.ColIndex("Emp_id")) = True
        .ColHidden(.ColIndex("WorkHours")) = True
        .ColHidden(.ColIndex("OverTime")) = True
        .ColHidden(.ColIndex("SalesCom")) = True
        .ColHidden(.ColIndex("cost_center_id")) = True
        .ColHidden(.ColIndex("CorrectEmpTotalNet")) = True
        .ColHidden(.ColIndex("DefWorkHours")) = True
        .ColHidden(.ColIndex("LocationID")) = True

    End With
 
    Frame1.Caption = "JV Data"
    Label1.Caption = "JV NO."
    ALLButton4.Caption = "Print JV"

    With Me.GRID1
  
  .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("payed")) = "Select"
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("Emp_id")) = "Emp.ID"
        .TextMatrix(0, .ColIndex("emp_code")) = "Emp.Code"
        .TextMatrix(0, .ColIndex("emp_name")) = "Emp.Name"
        .TextMatrix(0, .ColIndex("Mokafea")) = "Additional"
        .TextMatrix(0, .ColIndex("TotalAdvance")) = "Advances"
        .TextMatrix(0, .ColIndex("TotalDiscount")) = "Discounts"
        .TextMatrix(0, .ColIndex("SalesCom")) = "Sales Com."
        .TextMatrix(0, .ColIndex("EmpTotalNet")) = "Net "
        .TextMatrix(0, .ColIndex("sgn")) = "sgn"
        .ColHidden(.ColIndex("dep")) = True
        .ColHidden(.ColIndex("Branchid")) = True
        .ColHidden(.ColIndex("branchname")) = True
        .ColHidden(.ColIndex("project")) = True
        .ColHidden(.ColIndex("Emp_id")) = True
        .ColHidden(.ColIndex("WorkHours")) = True
        .ColHidden(.ColIndex("OverTime")) = True
        .ColHidden(.ColIndex("SalesCom")) = True
        .ColHidden(.ColIndex("cost_center_id")) = True
        .ColHidden(.ColIndex("DefWorkHours")) = True
        .TextMatrix(0, .ColIndex("total1")) = "Total Add. "
        .TextMatrix(0, .ColIndex("total2")) = "Total Discount. "

    End With

    ALLButton2.Caption = "Create Jv"
 
    With CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
         .AddItem "ATM"
          .AddItem "Transfer"
          
        .AddItem "Account"
    End With

End Function
 Public Sub FillGridWithData(Optional ByVal IsRecreate As Boolean = False)
    Dim i As Integer
    Dim j As Integer
Dim AllwIntro As Double
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
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
empDes = ""
    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    If val(Me.TxtMonthHours.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈œŒ«· ⁄œœ ”«⁄«  «·⁄„· ·Â–« «·‘Â—"
        Else
            Msg = "Enter Work Hours to this Month"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String
    '    My_SQL = " Select  lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id"
    '    My_SQL = My_SQL + ",IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  "
    '    My_SQL = My_SQL + "IsNUll( TotalDiscount,0)as TotalDiscount,"
    '    My_SQL = My_SQL + "IsNUll(TotalMokafea, 0) As TotalMokafea"
    '    My_SQL = My_SQL + ""
    '    My_SQL = My_SQL + ",(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-"
    '    My_SQL = My_SQL + "(IsNUll(TotalDiscount,0)) as EmpTotalNet "
    '
    '    My_SQL = My_SQL + " From "
    '    My_SQL = My_SQL + "("
    '    My_SQL = My_SQL + "SELECT TOP 100 PERCENT lastHolidaydate,BignDateWork,  fullcode,groupid, BranchId,dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID , dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
    '    My_SQL = My_SQL + "dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary,dbo.TblEmployee.cost_center_id ,"
    '    My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount,"
    '    My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea"
    '    My_SQL = My_SQL + ""
    '
    '    My_SQL = My_SQL + " From dbo.QryAllDiscountWithMkafea(" & IntMonth & "," & IntYear & ")"
    '    My_SQL = My_SQL + " QryAllDiscountWithMkafea RIGHT OUTER JOIN"
    '    My_SQL = My_SQL + "  dbo.TblEmployee ON QryAllDiscountWithMkafea.Emp_ID = dbo.TblEmployee.Emp_ID"
    '
'IntMonth IntYear
'    My_SQL = " Select Nationality,lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE ,dbo.EmpInsurances(" & IntMonth - 1 & "," & IntYear & ",Emp_ID) AS ToalInsurance ,EmpPrePaymentID(Emp_ID)as PrePaidID ,EmpPrePaymentValue(EmpPrePaymentID(Emp_ID))as PrePaidvalue" & Chr(13)
'  My_SQL = My_SQL + "  From (" & Chr(13)
'
'  My_SQL = My_SQL + "  SELECT    TOP 100 PERCENT  Nationality, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & Chr(13)
'  My_SQL = My_SQL + "                     dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & Chr(13)
'  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE  " & Chr(13)
'  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & Chr(13)
'  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & Chr(13)
'  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & Chr(13)
'  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & Chr(13)
  My_SQL = " SELECT   SalaryType, Nationality, lastHolidaydate, BignDateWork, Fullcode, GroupID, BranchId, Emp_ID, Emp_Code, Emp_Name,Emp_Namee ,DepartmentID,DeptID2, project_id, cost_center_id,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary, 0) AS Emp_Salary, ISNULL(Emp_Salary_sakn, 0) AS Emp_Salary_sakn, ISNULL(XTable.Emp_Salary_bus, 0) AS Emp_Salary_bus,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary_food, 0) AS Emp_Salary_food, ISNULL(Emp_Salary_others, 0) AS Emp_Salary_others, ISNULL(Emp_Salary_mob, 0) AS Emp_Salary_mob,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary_mang, 0) AS Emp_Salary_mang, ISNULL(TotalDiscount, 0) AS TotalDiscount, ISNULL(TotalMokafea, 0) AS TotalMokafea, ISNULL(Emp_Salary, 0)"
  My_SQL = My_SQL + "                     + ISNULL(TotalMokafea, 0) - ISNULL(XTable.TotalDiscount, 0) AS EmpTotalNet,JobTypeName,JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,"
 My_SQL = My_SQL + "                     dbo.EmpVoCation2(" & CmbMonth.ListIndex + 1 & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation2,"
 My_SQL = My_SQL + "                     dbo.EmpVoCation4(" & CmbMonth.ListIndex + 1 & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation4,"
  My_SQL = My_SQL + "                     dbo.EmpVoCation3(" & CmbMonth.ListIndex + 1 & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation3,"
 
  
  My_SQL = My_SQL + "                     Project_nameE,dbo.EmpVoCation(" & CmbMonth.ListIndex & ", " & val(CboYear.text) & ",Emp_ID) AS VoCation, dbo.EmpInsurances(" & val(CmbMonth.ListIndex) & ", " & val(CboYear.text) & ",Emp_ID) AS ToalInsurance, dbo.EmpPrePaymentID(XTable.Emp_ID) AS PrePaidID, dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(XTable.Emp_ID))"
  
  My_SQL = My_SQL + "                     AS PrePaidvalue  , dbo.GetAbcentDay(XTable.Emp_ID," & val(CboYear.text) & ", " & val(CmbMonth.ListIndex) + 1 & ") AS AbcentDay ,"
  My_SQL = My_SQL + "                     dbo.GetAbcentDay2(XTable.Emp_ID," & val(CboYear.text) & ", " & val(CmbMonth.ListIndex) + 1 & ") AS VacDay"
  My_SQL = My_SQL + "                     ,XTable.SalaryType ,XTable.SalaryCode ,name,XTable.namee , DepartmentName, DepartmentNamee,GroupName,GroupNameE"
  
  
  My_SQL = My_SQL + ""
  My_SQL = My_SQL + "    FROM         (SELECT      dbo.TblEmployee.SalaryType, dbo.TblEmployee.Nationality, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, "
  My_SQL = My_SQL + "                    dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID,"
  My_SQL = My_SQL + "                     dbo.TblEmployee.DeptID2, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
  My_SQL = My_SQL + "                     dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
  My_SQL = My_SQL + "                    dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount,"
  My_SQL = My_SQL + "                    SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
  My_SQL = My_SQL + "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name,"
  My_SQL = My_SQL + "                    dbo.projects.Project_nameE, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.SalaryCode, dbo.TblEmployee.NationlID, dbo.Nationality.name, dbo.Nationality.namee,"
  My_SQL = My_SQL + "                    dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee ,dbo.EmpGroupDep.GroupNameE , dbo.EmpGroupDep.GroupName"
  My_SQL = My_SQL + "    FROM            dbo.TblEmpJobsTypes INNER JOIN"
  My_SQL = My_SQL + "                    dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                    dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID"

      
                        
        If DcEmp.text <> "" Then
            My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.Emp_id=" & val(DcEmp.BoundText) ' & "'"
        Else
  
            If Dcdep.text <> "" Then
    
                If dcproject.BoundText = "" Then
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & val(Dcdep.BoundText) & "'"
                Else
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & Dcdep.BoundText & "' and dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If dcproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                    Else
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                    End If
    
                Else
    
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                End If
            End If
        End If

        My_SQL = My_SQL + " and NOT( dbo.TblEmployee.BignDateWork   IS NULL )"
       ' My_SQL = My_SQL + " and dbo.TblEmployee.lastHolidaydate<" & SQLDate(DTP_Date.value, True)
        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork  <" & SQLDate(DTP_Date.value, True)
        
        
        
       If val(DcbDepartment2.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DeptID2=" & val(DcbDepartment2.BoundText)
   End If
   
   If val(DCGroupID.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.workstate=1 and dbo.TblEmployee.GroupID=" & val(DCGroupID.BoundText)
   End If
        If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
        My_SQL = My_SQL + " and PayType=" & val(cboPayType.ListIndex) & " and  dbo.TblEmployee.workstate=1"
    End If
     If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and  dbo.TblEmployee.workstate=1  and SalaryCode=N'" & Trim(DcbHemiaSalary.text) & "' "

    End If
    
   If val(DcbTeam.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and  dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.SpecificationID=" & val(DcbTeam.BoundText)
   End If
   
   
    If val(dcBranch1.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and  dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.BranchId=" & val(dcBranch1.BoundText)
   End If
   
        If val(dcempcontract.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.workstate=1 and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
   End If
   
  
 '       My_SQL = My_SQL + " GROUP BY  lastHolidaydate,BignDateWork,  fullcode,groupid,BranchId, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code,dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others,dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,dbo.TblEmployee.cost_center_id ,"
 '       My_SQL = My_SQL + " dbo.TblEmployee.Emp_Salary,dbo.TblEmployee.DepartmentID ,dbo.TblEmployee.project_id"
 '
 '       My_SQL = My_SQL + " ORDER BY (  dbo.TblEmployee.fullcode)"
  
 '       My_SQL = My_SQL +  "  )XTable"
 
  If IsRecreate Then
        Dim WS As String
        WS = "(dbo.TblEmployee.workstate=1 OR " & _
             VoucherExistsClause(IntYear, IntMonth, "dbo.TblEmployee", "Account_code1") & ")"

        ' «” »œ· ﬂ· «·„Ê«÷⁄ «··Ì ﬂ‰  Õ«ÿÿ ›ÌÂ« workstate=1
        My_SQL = Replace(My_SQL, "dbo.TblEmployee.workstate=1", WS)
    End If
 
My_SQL = My_SQL + "   GROUP BY dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, "
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name,"
My_SQL = My_SQL + "                                              dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblEmployee.Nationality,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.SalaryType, dbo.TblEmployee.SalaryCode, dbo.TblEmployee.NationlID, dbo.Nationality.name,"
My_SQL = My_SQL + "                                              dbo.Nationality.NameE,dbo.TblEmployee.DeptID2 ,dbo.TblEmpDepartments.DepartmentName , dbo.TblEmpDepartments.DepartmentNamee ,dbo.EmpGroupDep.GroupName , dbo.EmpGroupDep.GroupNameE"
My_SQL = My_SQL + "                        ) XTable  ORDER BY Fullcode"




 Else
        FrstDay = "1-" & CmbMonth.ListIndex + 1 & "-" & year(Date)
        LstDay = DateAdd("d", -1, "1-" & CmbMonth.ListIndex + 2 & "-" & year(Date))

        My_SQL = "select Emp_ID,Emp_Name,Emp_Salary ,sum(TotalDiscount) as TotalDiscount," & "sum(Mokafea) as Mokafea  From QryEmpAllValues where TransDate >=#" & Format(FrstDay, "mm/dd/yyyy") & "# and TransDate<=#" & Format(LstDay, "mm/dd/yyyy") & "# " & StrWhere & " GROUP BY Emp_ID, Emp_Name, " & "Emp_Salary  "
    End If





    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'·Õ›Ÿ «·ﬁÌ„ ﬁ»· «· √ÀÌ— »«·Œ’Ê„«  «Ê Õ”«» «Ï „” ﬁÿ⁄«  „‰ «·«Ã«“« 
Dim mValueWithOutDisc As Double

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
        Debug.Print rs.RecordCount
            .rows = rs.RecordCount + 1
            rs.MoveFirst
            Dim DaysInMonth22 As Double
            Dim CountDays22 As Double
Dim CountDays As Double
Dim countFlag As Double
Dim MonthDayNo  As Double
                    If SystemOptions.MonthIs30days = True Then
MonthDayNo = 30
Else
MonthDayNo = daysInMonth(DTP_Date.value)

End If
CountDays = MonthDayNo
Dim mRemainDay  As Double
Dim mRemainDay2 As Double
            For i = 1 To .rows - 1
            
            mRemainDay2 = 0
            mRemainDay = 0
            
                If SystemOptions.MonthIs30days = True Then
                    MonthDayNo = 30
                Else
                    MonthDayNo = daysInMonth(DTP_Date.value)
    
                End If
                CountDays = MonthDayNo
            
         countFlag = 0
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("Location")) = IIf(IsNull(rs.Fields("GroupName").value), "", rs.Fields("GroupName").value)
               .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
            Else
               .TextMatrix(i, .ColIndex("Location")) = IIf(IsNull(rs.Fields("GroupNameE").value), "", rs.Fields("GroupNameE").value)
                .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
          End If
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(rs.Fields("BignDateWork").value), "", rs.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(rs.Fields("lastHolidaydate").value), "", rs.Fields("lastHolidaydate").value)
            .TextMatrix(i, .ColIndex("SalaryType")) = IIf(IsNull(rs.Fields("SalaryType").value), 0, rs.Fields("SalaryType").value)
            .TextMatrix(i, .ColIndex("LocationID")) = IIf(IsNull(rs.Fields("GroupID").value), 0, rs.Fields("GroupID").value)
      
           If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("BignDateWork"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("BignDateWork"))) Then
           'CountDays
               countFlag = 1
               CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("BignDateWork")), DTP_Date.value)
               
            
              
              
              CountDays = CountDays + 1
               If CountDays = daysInMonth(DTP_Date.value) Then
                  ' CountDays = 30
              End If
                   .TextMatrix(i, .ColIndex("CountDays")) = CountDays
              Else
                   countFlag = 0
                   .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
              End If
           
           
           
           If IsDate(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           
                      If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("lastHolidaydate"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           'CountDays
                countFlag = 1
                CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("lastHolidaydate")), DTP_Date.value)
                CountDays22 = DateDiff("D", .TextMatrix(i, .ColIndex("lastHolidaydate")), DTP_Date.value)
                CountDays22 = CountDays22 + 1
                DaysInMonth22 = daysInMonth(DTP_Date.value)
                CountDays = CountDays + 1
                 If CountDays = daysInMonth(DTP_Date.value) Then
                     If SystemOptions.MonthIs30days = True Then
                         CountDays = 30
                     End If
                End If
                    .TextMatrix(i, .ColIndex("CountDays")) = CountDays
         Else
                countFlag = 0
                 .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
           End If
           
           
           End If
           
            .TextMatrix(i, .ColIndex("PrePaidvalue")) = IIf(IsNull(rs.Fields("PrePaidvalue").value), 0, rs.Fields("PrePaidvalue").value)
            .TextMatrix(i, .ColIndex("PrePaidID")) = IIf(IsNull(rs.Fields("PrePaidID").value), 0, rs.Fields("PrePaidID").value)
            
            .TextMatrix(i, .ColIndex("AbcentDay")) = IIf(IsNull(rs.Fields("AbcentDay").value), 0, rs.Fields("AbcentDay").value)
            .TextMatrix(i, .ColIndex("vacDay")) = IIf(IsNull(rs.Fields("vacDay").value), 0, rs.Fields("vacDay").value)
                
                
             If CountDays = 12 Then
                CountDays = CountDays
             End If
                
                
                
                .TextMatrix(i, .ColIndex("VoCation")) = IIf(IsNull(rs.Fields("VoCation").value), 0, Round(rs.Fields("VoCation").value, SystemOptions.EmpSalaryDigts))
                
.TextMatrix(i, .ColIndex("VoCation2")) = IIf(IsNull(rs.Fields("VoCation2").value), 0, Round(rs.Fields("VoCation2").value, SystemOptions.EmpSalaryDigts))
.TextMatrix(i, .ColIndex("VoCation3")) = IIf(IsNull(rs.Fields("VoCation3").value), 0, Round(rs.Fields("VoCation3").value, SystemOptions.EmpSalaryDigts))
.TextMatrix(i, .ColIndex("VoCation4")) = IIf(IsNull(rs.Fields("VoCation4").value), 0, Round(rs.Fields("VoCation4").value, SystemOptions.EmpSalaryDigts))
                ' .TextMatrix(i, .ColIndex("ToalInsurance")) = 1
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
                empDes = (.TextMatrix(i, .ColIndex("Emp_ID"))) + "," + empDes
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                
                     If .TextMatrix(i, .ColIndex("Emp_Code")) = "2127S1171" Then
                
            CountDays = CountDays
           End If
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
                '                .TextMatrix(i, .ColIndex("Comp1")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", Round(rs.Fields("Emp_Salary").value, Decimal_Places))
            
                .TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(.TextMatrix(i, .ColIndex("AbcentDay"))) - val(val(.TextMatrix(i, .ColIndex("vacDay"))))
                mRemainDay = val(.TextMatrix(i, .ColIndex("RemainDay")))
                
                          If SystemOptions.UserInterface = ArabicInterface Then
                  
                  Else
                  
                  End If
                  
                  If CountDays = 0 Then
                     If CountDays = daysInMonth(DTP_Date.value) Then
                        CountDays = 30
                    Else
                       CountDays = MonthDayNo
                        .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo

                    End If

               End If
                .TextMatrix(i, .ColIndex("CountDays")) = CountDays
                
                   
                
                   
                   '.TextMatrix(i, .ColIndex("ToalInsurance")) = ((IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))) / MonthDayNo) * val(.TextMatrix(i, .ColIndex("RemainDay")))
'                   .TextMatrix(i, .ColIndex("ToalInsurance")) = val(((IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))) / MonthDayNo) * val(.TextMatrix(i, .ColIndex("MonthDayNo"))))
                   
                   .TextMatrix(i, .ColIndex("ToalInsurance")) = ((IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))))
          If SystemOptions.UserInterface = ArabicInterface Then
          .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs.Fields("DepartmentName").value), "", rs.Fields("DepartmentName").value)
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeName").value), "", rs.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs.Fields("DepartmentNamee").value), "", rs.Fields("DepartmentNamee").value)
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0
If val(.TextMatrix(i, .ColIndex("SalaryType"))) <> 4 Then

                For j = 1 To 40
                    ColumnName = "Comp" & j
                    If j = 11 Then
                        j = j
                    End If
                    If ViewComp(j) = True Then
                    AllwIntro = Round(GetValueAllwIntro(CmbMonth.ListIndex + 1, val(CboYear.text), val(.TextMatrix(i, .ColIndex("Emp_ID"))), j), SystemOptions.EmpSalaryDigts)
                    If AllwIntro > 0 Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                                           If countFlag = 1 Then
                                           If showMofradAll(j) = False Then
                                           If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, SystemOptions.EmpSalaryDigts)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, SystemOptions.EmpSalaryDigts)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), SystemOptions.EmpSalaryDigts)
                                          End If
                                           End If
                                           .TextMatrix(i, .ColIndex("TotalVacValue")) = ((IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))) / MonthDayNo) * val(.TextMatrix(i, .ColIndex("RemainDay")))
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                                                     
                        End If
                    Else
                                            If FixedOrChanged(j) = 0 Then
                                            
                                                     ' «·ﬂÊœ «·ﬁœÌ„ ﬂ«‰ »Ì÷—» ›Ï ⁄œœ «·«Ì«„ «·„ „À·… ›Ï MonthDayNo
                                                    '.TextMatrix(i, .ColIndex("TotalVacValue")) = Round(val(.TextMatrix(i, .ColIndex("TotalVacValue"))) + (val(.TextMatrix(i, .ColIndex(ColumnName)))) / MonthDayNo * val(.TextMatrix(i, .ColIndex("vacDay"))), SystemOptions.EmpSalaryDigts)
                                                    '  „ «· ⁄œÌ· » «—ÌŒ 14-06-2023 »Ê«”ÿ… Ê«∆· »‰«¡« ⁄·Ï ÿ·» „ ⁄„«œ Ê⁄„«— «·Ï «·«Ì«„ «·„ »ﬁÌ…
                                            
                                            .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DTP_Date.value, val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)
                                            mValueWithOutDisc = val(.TextMatrix(i, .ColIndex(ColumnName)))
                                            If mRemainDay <> 0 And mRemainDay <> MonthDayNo Then
                                                If val(val(.TextMatrix(i, .ColIndex("vacDay")))) <> 0 Then
                                                
                                                    
                                                    mRemainDay2 = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(val(.TextMatrix(i, .ColIndex("vacDay"))))
                                                    
                                                    .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * mRemainDay2, SystemOptions.EmpSalaryDigts)
                                                 '   .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * mRemainDay, SystemOptions.EmpSalaryDigts)
                                                Else
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) <> 0 Then
                                                      '  .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * mRemainDay, SystemOptions.EmpSalaryDigts)
                                                    End If
                                                End If
                                            ElseIf MonthDayNo = val(val(.TextMatrix(i, .ColIndex("vacDay")))) + val(.TextMatrix(i, .ColIndex("AbcentDay"))) Then
                                                '.TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * mRemainDay2, SystemOptions.EmpSalaryDigts)
                                                mRemainDay2 = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(val(.TextMatrix(i, .ColIndex("vacDay"))))
                                                
                                                If MonthDayNo = val(val(.TextMatrix(i, .ColIndex("vacDay")))) Then
                                                    mRemainDay2 = MonthDayNo
                                                End If
                                                .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * mRemainDay2, SystemOptions.EmpSalaryDigts)
                                               ' .TextMatrix(i, .ColIndex("TotalVacValue")) = Round(val(.TextMatrix(i, .ColIndex("TotalVacValue"))) + (val(mValueWithOutDisc)) / MonthDayNo * val(.TextMatrix(i, .ColIndex("vacDay"))), SystemOptions.EmpSalaryDigts)
                                            End If
                                            If val(val(.TextMatrix(i, .ColIndex("vacDay")))) <> 0 And val(.TextMatrix(i, .ColIndex(ColumnName))) <> 0 Then
                                                If culc30orRminder(j) = 0 Then
                                                    
                                                    .TextMatrix(i, .ColIndex("TotalVacValue")) = Round(val(.TextMatrix(i, .ColIndex("TotalVacValue"))) + (val(mValueWithOutDisc)) / MonthDayNo * val(.TextMatrix(i, .ColIndex("vacDay"))), SystemOptions.EmpSalaryDigts)
                                                Else
                                                    .TextMatrix(i, .ColIndex("TotalVacValue")) = Round(val(.TextMatrix(i, .ColIndex("TotalVacValue"))) + (val(mValueWithOutDisc)) / DaysInMonth22 * val(.TextMatrix(i, .ColIndex("vacDay"))), SystemOptions.EmpSalaryDigts)
                                                End If
                                          
                                                
                                                
                                            End If
                                           If countFlag = 1 Then
                                                If showMofradAll(j) = False Then
                                                   If culc30orRminder(j) = 0 Then
                                                          '.TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, SystemOptions.EmpSalaryDigts)
                                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round((Round(GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DTP_Date.value, val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)) / MonthDayNo * CountDays, SystemOptions.EmpSalaryDigts)
                                                    Else
                                                          '.TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, SystemOptions.EmpSalaryDigts)
                                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round((Round(GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DTP_Date.value, val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)) / DaysInMonth22 * CountDays22, SystemOptions.EmpSalaryDigts)
                                                    End If
                                                 Else
                                                    .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), SystemOptions.EmpSalaryDigts)
                                                 End If
                                           End If
                            '.TextMatrix(i, .ColIndex("TotalVacValue")) = ((IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))) / MonthDayNo) * val(.TextMatrix(i, .ColIndex("RemainDay")))
                        Else
                           .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1), SystemOptions.EmpSalaryDigts)
                            
                                                     
                        
                        End If
                    End If
                        
                    End If
                If val(.TextMatrix(i, .ColIndex(ColumnName))) <> 0 Then
                     i = i
                End If
                Next j
      Else
                      For j = 1 To 40
                    ColumnName = "Comp" & j
                    
                    If ViewComp(j) = True Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeSalaryProject(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)
                    End If
    
                Next j
      End If
    
                '   .TextMatrix(I, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), _
                    "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                '      .TextMatrix(i, .ColIndex("total1"))
    '«Ìﬁ«› «” œ⁄«¡ «·„ﬂ«ﬁ√  Ê «·Œ’Ê„«  30 09 2018
    '************************************************************
   '             .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, Decimal_Places))
   '
   '             .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Round(rs.Fields("TotalMokafea").value, Decimal_Places))
              '************************************************************
              
                rs.MoveNext
            
            Next

            rs.Close
        End If
If CheckVacation() = False Then
  If Len(empDes) > 0 Then
    empDes = mId(empDes, 1, Len(empDes) - 1)
    End If
        GetAdvanceValues IntMonth, IntYear
        GetEmpBalance
        ' GetWorkHours
        CalculateNets
      .rows = .rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
        End If

        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        '    SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal

        For j = 1 To 40
            ColumnName = "Comp" & j
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex(ColumnName), .rows - 1, .ColIndex(ColumnName))
            .TextMatrix(.rows - 1, .ColIndex(ColumnName)) = SngTotal

        Next j

        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
'        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
        
          SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpBalance"), .rows - 1, .ColIndex("EmpBalance"))
        .TextMatrix(.rows - 1, .ColIndex("EmpBalance")) = SngTotal
        
        
          SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpReaminB"), .rows - 1, .ColIndex("EmpReaminB"))
         .TextMatrix(.rows - 1, .ColIndex("EmpReaminB")) = SngTotal
        

        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal

        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal

        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
        
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VoCation2"), .rows, .ColIndex("VoCation2"))
        .TextMatrix(.rows - 1, .ColIndex("VoCation2")) = SngTotal
'
'            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
'            .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
            .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
'       SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VoCation3"), .rows - 1, .ColIndex("VoCation3"))
'            .TextMatrix(.rows - 1, .ColIndex("VoCation3")) = SngTotal
    
    
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ToalInsurance"), .rows - 1, .ColIndex("ToalInsurance"))
            .TextMatrix(.rows - 1, .ColIndex("ToalInsurance")) = SngTotal
            
   
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalVacValue"), .rows - 1, .ColIndex("TotalVacValue"))
        .TextMatrix(.rows - 1, .ColIndex("TotalVacValue")) = SngTotal
        
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VoCation3"), .rows, .ColIndex("VoCation3"))
        .TextMatrix(.rows - 1, .ColIndex("VoCation3")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("VoCation4"), .rows, .ColIndex("VoCation4"))
        .TextMatrix(.rows - 1, .ColIndex("VoCation4")) = SngTotal

        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
        
         .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = 0
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
            

        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
     .AutoSize 0, .Cols - 1, False
       Else
 FillGridWithDataAbscens
End If
  End With

'rs.Close
Set rs = Nothing

    Coloring
    If SystemOptions.UserInterface = EnglishInterface Then
        ChangeLang
    End If
ErrTrap:
End Sub
Function CheckEmpAlredyExists(Optional EmpID As Double) As Boolean
Dim i As Integer
With Me.Grid
CheckEmpAlredyExists = False
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("Emp_ID"))) = EmpID Then
CheckEmpAlredyExists = True
Exit Function
End If
Next i
End With
End Function
Public Sub FillGridWithDataAbscens()
    Dim i As Integer
    Dim j As Integer
Dim AllwIntro As Double
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
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim m As Integer
empDes = ""
    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    If val(Me.TxtMonthHours.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈œŒ«· ⁄œœ ”«⁄«  «·⁄„· ·Â–« «·‘Â—"
        Else
            Msg = "Enter Work Hours to this Month"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String

  My_SQL = " SELECT   SalaryType,  Nationality, lastHolidaydate, BignDateWork, Fullcode, GroupID, BranchId, Emp_ID, Emp_Code, Emp_Name,Emp_Namee ,DepartmentID,DeptID2, project_id, cost_center_id,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary, 0) AS Emp_Salary, ISNULL(Emp_Salary_sakn, 0) AS Emp_Salary_sakn, ISNULL(Emp_Salary_bus, 0) AS Emp_Salary_bus,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary_food, 0) AS Emp_Salary_food, ISNULL(Emp_Salary_others, 0) AS Emp_Salary_others, ISNULL(Emp_Salary_mob, 0) AS Emp_Salary_mob,"
  My_SQL = My_SQL + "                     ISNULL(Emp_Salary_mang, 0) AS Emp_Salary_mang, ISNULL(TotalDiscount, 0) AS TotalDiscount, ISNULL(TotalMokafea, 0) AS TotalMokafea, ISNULL(Emp_Salary, 0)"
  My_SQL = My_SQL + "                     + ISNULL(TotalMokafea, 0) - ISNULL(TotalDiscount, 0) AS EmpTotalNet, JobTypeName, JobTypeNamee, branch_name, branch_namee, projectFullcode, Project_name,"
  My_SQL = My_SQL + "                     Project_nameE,dbo.EmpVoCation(" & CmbMonth.ListIndex & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation, dbo.EmpInsurances(" & CmbMonth.ListIndex & ", " & val(CboYear.text) & ", Emp_ID) AS ToalInsurance, dbo.EmpPrePaymentID(Emp_ID) AS PrePaidID, dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(Emp_ID))"
  My_SQL = My_SQL + "                     ,dbo.EmpVoCation2(" & CmbMonth.ListIndex + 1 & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation2,"
  My_SQL = My_SQL + "                     dbo.EmpVoCation3(" & CmbMonth.ListIndex + 1 & ", " & val(CboYear.text) & ", Emp_ID) AS VoCation3,"
  My_SQL = My_SQL + "                    0 AS PrePaidvalue  , dbo.GetAbcentDay(Emp_ID," & val(CboYear.text) & ", " & val(CmbMonth.ListIndex) + 1 & ") AS AbcentDay "
    My_SQL = My_SQL + "                        , dbo.GetAbcentDay2(Emp_ID," & val(CboYear.text) & ", " & val(CmbMonth.ListIndex) + 1 & ") AS VacDay"
    My_SQL = My_SQL + "                      ,SalaryType ,SalaryCode , name, namee "
  My_SQL = My_SQL + ""
  My_SQL = My_SQL + "    FROM         (SELECT      dbo.TblEmployee.SalaryType, dbo.TblEmployee.Nationality, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork,"
  My_SQL = My_SQL + "                                            dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID,dbo.TblEmployee.DeptID2 ,"
  My_SQL = My_SQL + "                                            dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
  My_SQL = My_SQL + "                                            dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
  My_SQL = My_SQL + "                                            dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount)"
  My_SQL = My_SQL + "                                            AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea, dbo.TblEmpJobsTypes.JobTypeName,"
  My_SQL = My_SQL + "                                            dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  My_SQL = My_SQL + "                                            dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblEmployee.Emp_Namee,"
  My_SQL = My_SQL + "                                            dbo.TblEmployee.SalaryCode , dbo.TblEmployee.NationlID, dbo.Nationality.Name, dbo.Nationality.NameE"
  My_SQL = My_SQL + "                      FROM         dbo.TblEmpJobsTypes INNER JOIN"
  My_SQL = My_SQL + "                                            dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN"
  My_SQL = My_SQL + "                                            dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                                            dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                                            dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  My_SQL = My_SQL + "                                             dbo.QryAllDiscountWithMkafea (" & IntMonth & ", " & IntYear & ")  QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID"
      Dim JobStus As Integer
           JobStus = GetJobStatus
           If JobStus = 0 Then
           JobStus = -1
           End If
        If DcEmp.text <> "" Then
            My_SQL = My_SQL + " Where   dbo.TblEmployee.Emp_id=" & val(DcEmp.BoundText) ' & "'"
        Else
  
            If Dcdep.text <> "" Then
    
                If dcproject.BoundText = "" Then
                    My_SQL = My_SQL + " Where dbo.TblEmployee.jopstatusid=" & JobStus & " and dbo.TblEmployee.DepartmentID='" & val(Dcdep.BoundText) & "'"
                Else
                    My_SQL = My_SQL + " Where dbo.TblEmployee.jopstatusid=" & JobStus & " and dbo.TblEmployee.DepartmentID='" & Dcdep.BoundText & "' and dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If dcproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + " Where dbo.TblEmployee.jopstatusid=" & JobStus & " and  dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                    Else
                        My_SQL = My_SQL + " Where dbo.TblEmployee.jopstatusid=" & JobStus & ""
                    End If
    
                Else
    
                    My_SQL = My_SQL + " Where dbo.TblEmployee.jopstatusid=" & JobStus & ""
                End If
            End If
        End If
    My_SQL = My_SQL + " and  DATEDIFF(m, Dbo.GetVacationDate(dbo.TblEmployee.Emp_id), " & SQLDate(DTP_Date.value, True) & ")  >=1 "
      My_SQL = My_SQL + " and NOT( dbo.TblEmployee.BignDateWork   IS NULL )"
       ' My_SQL = My_SQL + " and dbo.TblEmployee.lastHolidaydate<" & SQLDate(DTP_Date.value, True)
      '  My_SQL = My_SQL + " and isnull( dbo.TblEmployee.lastHolidaydate,dbo.TblEmployee.BignDateWork ) <" & SQLDate(DTP_Date.value, True)
        
        '
        
       If val(DcbDepartment2.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.jopstatusid=" & JobStus & " and dbo.TblEmployee.DeptID2=" & val(DcbDepartment2.BoundText)
   End If
   
   If val(DCGroupID.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.jopstatusid=" & JobStus & " and dbo.TblEmployee.GroupID=" & val(DCGroupID.BoundText)
   End If
        If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
        My_SQL = My_SQL + " and PayType=" & val(cboPayType.ListIndex) & " and  dbo.TblEmployee.workstate=1"
    End If
     If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and  dbo.TblEmployee.jopstatusid=" & JobStus & "  and SalaryCode=N'" & Trim(DcbHemiaSalary.text) & "' "

    End If
    
   If val(DcbTeam.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and  dbo.TblEmployee.jopstatusid=" & JobStus & " and  dbo.TblEmployee.SpecificationID=" & val(DcbTeam.BoundText)
   End If
   
   
    If val(dcBranch1.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and  dbo.TblEmployee.jopstatusid=" & JobStus & " and  dbo.TblEmployee.BranchId=" & val(dcBranch1.BoundText)
   End If
   
        If val(dcempcontract.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.jopstatusid=" & JobStus & " and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
   End If
My_SQL = My_SQL + "   GROUP BY dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, "
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name,"
My_SQL = My_SQL + "                                              dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblEmployee.Nationality,"
My_SQL = My_SQL + "                                              dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.SalaryType, dbo.TblEmployee.SalaryCode, dbo.TblEmployee.NationlID, dbo.Nationality.name,"
My_SQL = My_SQL + "                                              dbo.Nationality.NameE,dbo.TblEmployee.DeptID2 "
' DATEDIFF(m, lastHolidaydate, " & SQLDate(DTP_Date.value, True) & ") AS Diff"
'My_SQL = My_SQL + "                        ORDER BY dbo.TblEmployee.Fullcode"
My_SQL = My_SQL + ") XTable"
 Else
        FrstDay = "1-" & CmbMonth.ListIndex + 1 & "-" & year(Date)
        LstDay = DateAdd("d", -1, "1-" & CmbMonth.ListIndex + 2 & "-" & year(Date))

        My_SQL = "select Emp_ID,Emp_Name,Emp_Salary ,sum(TotalDiscount) as TotalDiscount," & "sum(Mokafea) as Mokafea  From QryEmpAllValues where TransDate >=#" & Format(FrstDay, "mm/dd/yyyy") & "# and TransDate<=#" & Format(LstDay, "mm/dd/yyyy") & "# " & StrWhere & " GROUP BY Emp_ID, Emp_Name, " & "Emp_Salary  "
    End If
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    With Me.Grid
    Dim k As Integer
    k = .rows
        '.Rows = 2
       ' .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
        Debug.Print rs.RecordCount
            .rows = .rows + rs.RecordCount
            rs.MoveFirst
            Dim DaysInMonth22 As Double
            Dim CountDays22 As Double
Dim CountDays As Double
Dim countFlag As Double
Dim MonthDayNo  As Double
                    If SystemOptions.MonthIs30days = True Then
MonthDayNo = 30
Else
MonthDayNo = daysInMonth(DTP_Date.value)
End If

i = k - 1
            For m = k To .rows - 1
          If CheckEmpAlredyExists(IIf(IsNull(rs.Fields("Emp_ID").value), 0, rs.Fields("Emp_ID").value)) = False Then
          i = i + 1
          
         countFlag = 0
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
            Else
            .TextMatrix(i, .ColIndex("Nationality")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
            End If
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(rs.Fields("BignDateWork").value), "", rs.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(rs.Fields("lastHolidaydate").value), "", rs.Fields("lastHolidaydate").value)
            .TextMatrix(i, .ColIndex("SalaryType")) = IIf(IsNull(rs.Fields("SalaryType").value), 0, rs.Fields("SalaryType").value)
           
           
           If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("BignDateWork"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("BignDateWork"))) Then
           'CountDays
           countFlag = 1
           CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("BignDateWork")), DTP_Date.value)
           CountDays = CountDays + 1
                     If CountDays = daysInMonth(DTP_Date.value) Then
           CountDays = 30
           End If
           .TextMatrix(i, .ColIndex("CountDays")) = CountDays
           Else
           countFlag = 0
            .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
           End If
           
           If IsDate(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           
                      If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("lastHolidaydate"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           'CountDays
           countFlag = 1
           CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("lastHolidaydate")), DTP_Date.value)
           CountDays22 = DateDiff("D", .TextMatrix(i, .ColIndex("lastHolidaydate")), DTP_Date.value)
           CountDays22 = CountDays22 + 1
           DaysInMonth22 = daysInMonth(DTP_Date.value)
           CountDays = CountDays + 1
                     If CountDays = daysInMonth(DTP_Date.value) Then
           CountDays = 30
           End If
           .TextMatrix(i, .ColIndex("CountDays")) = CountDays
           Else
           countFlag = 0
            .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
           End If
           
           
           End If
           
            .TextMatrix(i, .ColIndex("PrePaidvalue")) = IIf(IsNull(rs.Fields("PrePaidvalue").value), 0, rs.Fields("PrePaidvalue").value)
            .TextMatrix(i, .ColIndex("PrePaidID")) = IIf(IsNull(rs.Fields("PrePaidID").value), 0, rs.Fields("PrePaidID").value)
            
            .TextMatrix(i, .ColIndex("AbcentDay")) = IIf(IsNull(rs.Fields("AbcentDay").value), 0, rs.Fields("AbcentDay").value)
            .TextMatrix(i, .ColIndex("vacDay")) = IIf(IsNull(rs.Fields("vacDay").value), 0, rs.Fields("vacDay").value)
            
                .TextMatrix(i, .ColIndex("ToalInsurance")) = IIf(IsNull(rs.Fields("ToalInsurance").value), 0, Round(rs.Fields("ToalInsurance").value, SystemOptions.EmpSalaryDigts))
                .TextMatrix(i, .ColIndex("VoCation")) = IIf(IsNull(rs.Fields("VoCation").value), 0, Round(rs.Fields("VoCation").value, SystemOptions.EmpSalaryDigts))
                .TextMatrix(i, .ColIndex("VoCation2")) = IIf(IsNull(rs.Fields("VoCation2").value), 0, Round(rs.Fields("VoCation2").value, SystemOptions.EmpSalaryDigts))
                .TextMatrix(i, .ColIndex("VoCation3")) = IIf(IsNull(rs.Fields("VoCation3").value), 0, Round(rs.Fields("VoCation3").value, SystemOptions.EmpSalaryDigts))
                ' .TextMatrix(i, .ColIndex("ToalInsurance")) = 1
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
                empDes = (.TextMatrix(i, .ColIndex("Emp_ID"))) + "," + empDes
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
                '                .TextMatrix(i, .ColIndex("Comp1")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", Round(rs.Fields("Emp_Salary").value, Decimal_Places))
            '.TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - .TextMatrix(i, .ColIndex("AbcentDay"))
            .TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(.TextMatrix(i, .ColIndex("AbcentDay"))) - val(.TextMatrix(i, .ColIndex("vacDay")))
            
                .TextMatrix(i, .ColIndex("RemainDay")) = val(.TextMatrix(i, .ColIndex("CountDays"))) - val(.TextMatrix(i, .ColIndex("AbcentDay"))) - val(.TextMatrix(i, .ColIndex("vacDay")))
                      If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeName").value), "", rs.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0
If val(.TextMatrix(i, .ColIndex("SalaryType"))) <> 4 Then

                For j = 1 To 40
                If showinMosirVac(j) = True Then
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                    AllwIntro = Round(GetValueAllwIntro(CmbMonth.ListIndex + 1, val(CboYear.text), val(.TextMatrix(i, .ColIndex("Emp_ID"))), j), SystemOptions.EmpSalaryDigts)
                    If AllwIntro > 0 Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                                           If countFlag = 1 Then
                                           If showMofradAll(i) = False Then
                                           If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, SystemOptions.EmpSalaryDigts)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, SystemOptions.EmpSalaryDigts)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), SystemOptions.EmpSalaryDigts)
                                          End If
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                                                     
                        End If
                    Else
                                            If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DTP_Date.value, val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)
                                           If countFlag = 1 Then
                                         If showMofradAll(j) = False Then
                                         If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, SystemOptions.EmpSalaryDigts)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, SystemOptions.EmpSalaryDigts)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), SystemOptions.EmpSalaryDigts)
                                          End If
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1), SystemOptions.EmpSalaryDigts)
                                                     
                        End If
                    End If
                        
                    End If
    End If
                Next j
      Else
                      For j = 1 To 40
                      If showinMosirVac(j) = True Then
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = Round(GetEmployeeSalaryProject(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), val(CmbMonth.ListIndex), val(CboYear.ListIndex)), SystemOptions.EmpSalaryDigts)
                    End If
    End If
                Next j
      End If
    
        
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, SystemOptions.EmpSalaryDigts)) ''Decimal_Places))
             
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Round(rs.Fields("TotalMokafea").value, SystemOptions.EmpSalaryDigts)) 'Decimal_Places))
              
                rs.MoveNext
            End If
            Next m
.Row = .rows - (m - i)
            rs.Close
        End If
  If Len(empDes) > 0 Then
    empDes = mId(empDes, 1, Len(empDes) - 1)
    End If
        GetAdvanceValues IntMonth, IntYear
        GetEmpBalance
        ' GetWorkHours
        CalculateNets
        .rows = .rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
        End If

        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        '    SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
   
        For j = 1 To 40
            ColumnName = "Comp" & j
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex(ColumnName), .rows - 1, .ColIndex(ColumnName))
            .TextMatrix(.rows - 1, .ColIndex(ColumnName)) = SngTotal
     
        Next j
      
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalVacValue"), .rows - 1, .ColIndex("TotalVacValue"))
        .TextMatrix(.rows - 1, .ColIndex("TotalVacValue")) = SngTotal
    
    
     
    
         SngTotal = val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpBalance"), .rows - 1, .ColIndex("TotaEmpBalancelAdvance")))
        .TextMatrix(.rows - 1, .ColIndex("EmpBalance")) = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpReaminB"), .rows - 1, .ColIndex("EmpReaminB"))
        .TextMatrix(.rows - 1, .ColIndex("EmpReaminB")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
    
        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With
 

'rs.Close
Set rs = Nothing

    Coloring
ErrTrap:
End Sub
Function CheckVacation() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     showinMosirVac"
sql = sql & " From dbo.MOFRAD"
sql = sql & " Where (showinMosirVac = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckVacation = True
Else
CheckVacation = False
End If
End Function
Function GetJobStatus() As Integer
Dim sql As String
GetJobStatus = 0
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     id, Vacation"
sql = sql & " From dbo.jopstatus"
sql = sql & " Where (Vacation = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetJobStatus = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
GetJobStatus = 0
End If
End Function
Function GetValueAllwIntro(Optional MothID As Integer, Optional YerID As Integer, Optional EmpID As Double, Optional MofrdID As Integer) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     MordValue / ISNULL(TypeMofrd, 1) AS Valu"
sql = sql & " From dbo.TblComponentYearDet"
sql = sql & " WHERE       (EmpID = " & EmpID & ") AND (MofrdID = " & MofrdID & ") and "
sql = sql & "               ((month(RecDate1) =" & MothID & " and Year(RecDate1) =" & YerID & ") or    ((month(RecDate2) =" & MothID & " and Year(RecDate2) =" & YerID & ")))"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetValueAllwIntro = IIf(IsNull(Rs3("Valu").value), 0, Rs3("Valu").value)
Else
GetValueAllwIntro = 0
End If
End Function
Public Sub FillGridWithData3()
'Exit Sub
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
    Dim j As Integer
    Dim ColumnName As String
    Dim SumAdVance As Double
    SumAdVance = 0
    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String
    
        ' My_SQL = "SELECT    BranchId,id,project_id, DepartmentID,id,Emp_id, Emp_Code, Emp_Name, Emp_Salary, Emp_Salary_sakn, Emp_Salary_bus, Emp_Salary_food, Emp_Salary_mob, Emp_Salary_mang, Emp_Salary_others,cost_center_id,"
        '  My_SQL = My_SQL + "OverTimePrice, Mokafea, SalesCom, total1, TotalAdvance, TotalDiscount, total2, EmpTotalNet, sgn, m_year, m_month, payed"
    '    My_SQL = "SELECT   * "
 
    '    My_SQL = My_SQL + "  from dbo.emp_salary WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

'My_SQL = "SELECT     *"
'My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID"

My_SQL = "SELECT     *"
My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
                      
'My_SQL = My_SQL + "   WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "
My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  AND (payed =0) "

        If DcEmp.text <> "" Then
            My_SQL = My_SQL + "  and  emp_salary.emp_code='" & DcEmp.BoundText & "'"
        Else

            If Dcdep.text <> "" Then
    
                If dcproject.BoundText = "" Then
                    My_SQL = My_SQL + "  and  emp_salary.DepartmentID='" & Dcdep.BoundText & "'"
                Else
                    My_SQL = My_SQL + "   and  emp_salary.DepartmentID='" & Dcdep.BoundText & "' and  project_id='" & Me.dcproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If dcproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + "  and  emp_salary.project_id='" & Me.dcproject.BoundText & "'"
                    End If
    
                End If
            End If
        End If

  '      If SystemOptions.usertype <> UserAdminAll Then
  '          My_SQL = My_SQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
            
  '      End If
        
        If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
          My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex) & " "
        End If
       If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "

       End If
          If Me.DcbDepartment2.text <> "" And val(DcbDepartment2.BoundText) <> 0 Then
          My_SQL = My_SQL + " and dbo.TblEmployee.DeptID2=" & val(DcbDepartment2.BoundText) & " "

       End If
      
    If val(dcBranch1.BoundText) <> 0 Then
   
     My_SQL = My_SQL + " and dbo.emp_salary.BranchId=" & val(dcBranch1.BoundText)
   End If
   
 
   
        If val(CboPayMentType.ListIndex) <> -1 Then
   
      My_SQL = My_SQL + "  and ( dbo.TblEmployee.PayType is null or  dbo.TblEmployee.PayType=" & val(CboPayMentType.ListIndex) & ")"
      End If
   
   
        If val(dcempcontract.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
   End If
   
   
   
        My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
        '  My_SQL = My_SQL + " order by   LPAD(Emp_code,6,'0') ASC"
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.Grid2
            .rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
                .rows = rs.RecordCount + 1
                rs.MoveFirst

                For i = 1 To .rows - 1

                    .TextMatrix(i, .ColIndex("Ser")) = i
       If GRID1.cell(flexcpChecked, i, GRID1.ColIndex("payed")) = flexUnchecked Then
       GoTo ll
       End If
                    '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
                       "", Rs.Fields("ID").value)
            
            
                   ' .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
            
            '        .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
                   ' .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
                     .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("NumEkama").value), "", rs.Fields("NumEkama").value)
                    
                             If Trim(.TextMatrix(i, .ColIndex("Emp_Code"))) = "" Then
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("NumPoket").value), "", rs.Fields("NumPoket").value)
                    End If
                    
     '             .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)

            '        .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
            '        .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            '
            '        .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                  .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    .TextMatrix(i, .ColIndex("Emp_Namee")) = IIf(IsNull(rs.Fields("Emp_Namee").value), IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value), rs.Fields("Emp_Namee").value)
                    If Trim(.TextMatrix(i, .ColIndex("Emp_Namee"))) = "" Then
                    .TextMatrix(i, .ColIndex("Emp_Namee")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    End If
                    .TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
                    .TextMatrix(i, .ColIndex("BankCard")) = IIf(IsNull(rs.Fields("BankCard").value), "", rs.Fields("BankCard").value)
                    .TextMatrix(i, .ColIndex("BanckCode")) = IIf(IsNull(rs.Fields("BankCode").value), "", rs.Fields("BankCode").value)
                    
               
                    '            .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", rs.Fields("Emp_Salary").value)
            
            '        .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
            '        .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("Mokafea").value), "", Round(rs.Fields("Mokafea").value, SystemOptions.SysDefCurrencyForamt))
            
            '        .TextMatrix(i, .ColIndex("TotalAdvance")) = IIf(IsNull(rs.Fields("TotalAdvance").value), "", Round(rs.Fields("TotalAdvance").value))
            '
            '        .TextMatrix(i, .ColIndex("SalesCom")) = IIf(IsNull(rs.Fields("SalesCom").value), "", Round(rs.Fields("SalesCom").value))
            '
            '        .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(rs.Fields("total1").value), "", Round(rs.Fields("total1").value, 2))
            '
            '        .TextMatrix(i, .ColIndex("total2")) = IIf(IsNull(rs.Fields("total2").value), "", Round(rs.Fields("total2").value, 2))
            
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, SystemOptions.EmpSalaryDigts))

                    For j = 1 To 40
            '            ColumnName = "Comp" & J
            '            .TextMatrix(i, .ColIndex(ColumnName)) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Format(rs.Fields(ColumnName).value))
                    Next j
    
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), _
                                          "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), _
                                          "", Format(rs.Fields("Emp_Salary_bus").value))
            
                    '
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), _
                                          "", Format(rs.Fields("Emp_Salary_food").value))
                               
                    '                            .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), _
                                                 "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                    '                                 .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), _
                                                      "", Format(rs.Fields("Emp_Salary_mang").value))
            
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), _
                                          "", Format(rs.Fields("Emp_Salary_others").value))
            
                    '                           .TextMatrix(i, .ColIndex("OverTimePrice")) = IIf(IsNull(rs.Fields("OverTimePrice").value), _
                                                "", Format(rs.Fields("OverTimePrice").value))
ll:
                    rs.MoveNext

                Next

                rs.Close
            End If
    
            GetAdvanceValues IntMonth, IntYear
            GetEmpBalance
            GetWorkHours
            CalculateNets
            .rows = .rows + 1

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = ""
            Else
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = ""
            End If

            .IsSubtotal(.rows - 1) = True
            Dim SngTotal As Single
            '    SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
            net_value1 = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    

    
            '        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
            '        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
            .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
            .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
            .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
    
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalVacValue"), .rows - 1, .ColIndex("TotalVacValue"))
            .TextMatrix(.rows - 1, .ColIndex("TotalVacValue")) = SngTotal
    
    
   
           ' SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpBalance"), .Rows - 1, .ColIndex("EmpBalance"))
           ' .TextMatrix(.Rows - 1, .ColIndex("EmpBalance")) = SngTotal
           '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpReaminB"), .Rows - 1, .ColIndex("EmpReaminB"))
           ' .TextMatrix(.Rows - 1, .ColIndex("EmpReaminB")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
            .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
            .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
            .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
            .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
            .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
            .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
            .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
            '  .AutoSize 0, .Cols - 1, False
        End With

    End If
'rs.Close
Set rs = Nothing
    Coloring
ErrTrap:
End Sub
Public Sub FillGridWithData2(Optional ByVal IsRecreate As Boolean = False)
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
    Dim j As Integer
    Dim ColumnName As String
    Dim SumAvanced As Double
    SumAvanced = 0
    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim ID As String
    
        ' My_SQL = "SELECT    BranchId,id,project_id, DepartmentID,id,Emp_id, Emp_Code, Emp_Name, Emp_Salary, Emp_Salary_sakn, Emp_Salary_bus, Emp_Salary_food, Emp_Salary_mob, Emp_Salary_mang, Emp_Salary_others,cost_center_id,"
        '  My_SQL = My_SQL + "OverTimePrice, Mokafea, SalesCom, total1, TotalAdvance, TotalDiscount, total2, EmpTotalNet, sgn, m_year, m_month, payed"
    '    My_SQL = "SELECT   * "
 
    '    My_SQL = My_SQL + "  from dbo.emp_salary WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

'My_SQL = "SELECT     *"
'My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID"
'My_SQL = My_SQL + "   WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

''My_SQL = "SELECT     *"
''My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
''My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
''My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'My_SQL = My_SQL + "   WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

My_SQL = " SELECT      *, dbo.emp_salary.TotalAdvance AS Expr1"
My_SQL = My_SQL + " FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'My_SQL = " SELECT     *, dbo.emp_salary.TotalAdvance AS Expr1, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee"
'My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
Dim sgn2 As String
sgn2 = Me.CboYear.text & CmbMonth.ListIndex + 1

If IsRecreate Then

    'My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  AND (payed =0) "
    
    My_SQL = My_SQL & " where 1 = 1 " & VoucherFilterBySgnMonthExact(sgn2, "TblEmployee", "Account_code1")
    My_SQL = My_SQL + "   and     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  AND (payed =0) "
Else
        My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  AND (payed =0) "
        
        
        
                If DcEmp.text <> "" Then
                    My_SQL = My_SQL + "  and  emp_salary.emp_code='" & DcEmp.BoundText & "'"
                Else
        
                    If Dcdep.text <> "" Then
            
                        If dcproject.BoundText = "" Then
                            My_SQL = My_SQL + "  and  emp_salary.DepartmentID='" & Dcdep.BoundText & "'"
                        Else
                            My_SQL = My_SQL + "   and  emp_salary.DepartmentID='" & Dcdep.BoundText & "' and  project_id='" & Me.dcproject.BoundText & "'"
                        End If
        
                    Else
        
                        If Dcdep.text = "" Then
            
                            If dcproject.BoundText <> "" Then
                
                                My_SQL = My_SQL + "  and  emp_salary.project_id='" & Me.dcproject.BoundText & "'"
                            End If
            
                        End If
                    End If
                End If
        
          '      If SystemOptions.usertype <> UserAdminAll Then
          '          My_SQL = My_SQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
                    
          '      End If
                
                
            If val(dcBranch1.BoundText) <> 0 Then
           
           My_SQL = My_SQL + " and dbo.emp_salary.BranchId=" & val(dcBranch1.BoundText)
           End If
            If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
                  My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex) & " "
            End If
          If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
                  My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "
        
            End If
              If Me.DcbDepartment2.text <> "" And val(DcbDepartment2.BoundText) <> 0 Then
                  My_SQL = My_SQL + " and dbo.TblEmployee.DeptID2=" & val(DcbDepartment2.BoundText) & " "
        
            End If
        
           
                If val(CboPayMentType.ListIndex) <> -1 Then
           
           My_SQL = My_SQL + "  and ( dbo.TblEmployee.PayType is null or  dbo.TblEmployee.PayType=" & val(CboPayMentType.ListIndex) & ")"
           End If
           
           
                If val(dcempcontract.BoundText) <> 0 Then
           
           My_SQL = My_SQL + " and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
           End If
           
 End If
   
        My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
        '  My_SQL = My_SQL + " order by   LPAD(Emp_code,6,'0') ASC"
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.GRID1
            .rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
                .rows = rs.RecordCount + 1
                rs.MoveFirst

                For i = 1 To .rows - 1
        
                    .TextMatrix(i, .ColIndex("Ser")) = i
            
                    '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
                       "", Rs.Fields("ID").value)
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
            
                    .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
                  '  .TextMatrix(i, .ColIndex("LocationID")) = IIf(IsNull(rs.Fields("LocationID").value), "", rs.Fields("LocationID").value)
                    .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
        
                '  .TextMatrix(i, .ColIndex("ToalInsurance")) = IIf(IsNull(rs.Fields("ToalInsurance").value), 0, rs.Fields("ToalInsurance").value)

               
                    '            .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", rs.Fields("Emp_Salary").value)
            
                    .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, SystemOptions.EmpSalaryDigts))
                
                    .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("Mokafea").value), "", Round(rs.Fields("Mokafea").value, SystemOptions.EmpSalaryDigts))
            
                    .TextMatrix(i, .ColIndex("TotalAdvance")) = IIf(IsNull(rs.Fields("TotalAdvance").value), 0, Round(rs.Fields("TotalAdvance").value, SystemOptions.EmpSalaryDigts))
            
                    .TextMatrix(i, .ColIndex("SalesCom")) = IIf(IsNull(rs.Fields("SalesCom").value), "", Round(rs.Fields("SalesCom").value, SystemOptions.EmpSalaryDigts))
                    
                    .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(rs.Fields("total1").value), "", Round(rs.Fields("total1").value, SystemOptions.EmpSalaryDigts))
            
                    .TextMatrix(i, .ColIndex("total2")) = IIf(IsNull(rs.Fields("total2").value), "", Round(rs.Fields("total2").value, SystemOptions.EmpSalaryDigts))
            
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, SystemOptions.EmpSalaryDigts))

                    For j = 1 To 40
                        ColumnName = "Comp" & j
                        .TextMatrix(i, .ColIndex(ColumnName)) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Format(rs.Fields(ColumnName).value))
                    Next j
    
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), _
                                          "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), _
                                          "", Format(rs.Fields("Emp_Salary_bus").value))
            
                    '
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), _
                                          "", Format(rs.Fields("Emp_Salary_food").value))
                               
                    '                            .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), _
                                                 "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                    '                                 .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), _
                                                      "", Format(rs.Fields("Emp_Salary_mang").value))
            
                    '                     .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), _
                                          "", Format(rs.Fields("Emp_Salary_others").value))
            
                    '                           .TextMatrix(i, .ColIndex("OverTimePrice")) = IIf(IsNull(rs.Fields("OverTimePrice").value), _
                                                "", Format(rs.Fields("OverTimePrice").value))
            
                    rs.MoveNext
            
                Next

                rs.Close
            End If
    
           GetAdvanceValues IntMonth, IntYear
           GetEmpBalance
          GetWorkHours
       CalculateNets
            .rows = .rows + 1
Exit Sub
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
            Else
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
            End If

            .IsSubtotal(.rows - 1) = True
            Dim SngTotal As Single
            '    SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
            net_value1 = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
            '        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
            '        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
            '    .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
            .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
            .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
             SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
            .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
           '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpBalance"), .Rows - 1, .ColIndex("EmpBalance"))
           ' .TextMatrix(.Rows - 1, .ColIndex("EmpBalance")) = SngTotal
           '   SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpReaminB"), .Rows - 1, .ColIndex("EmpReaminB"))
           ' .TextMatrix(.Rows - 1, .ColIndex("EmpReaminB")) = SngTotal
            
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
            .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
            .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
            .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
            .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
            .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
            .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
            .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
            '  .AutoSize 0, .Cols - 1, False
        End With

    End If
'rs.Close
Set rs = Nothing
    Coloring
 FillGridWithData3
    
ErrTrap:
End Sub


 
Private Sub GetWorkHours()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngFindRow As Long
    Dim i As Integer
    Dim X As Long
    Dim Y  As Long
    Dim Z As Long
    Dim IntYear As Integer, IntMonth As Integer
    Dim IntDefWorkHours As Integer

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    StrSQL = "SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(sum(dbo.tblPresentTime.WorkHoursCount)) AS WorkHours,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(SUM( dbo.tblPresentTime.WorkHoursCount - dbo.tblPresentTime.CurrentWorkMints))as OverTime"
    StrSQL = StrSQL + " FROM  dbo.TblEmployee LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.tblPresentTime ON dbo.TblEmployee.Emp_ID = dbo.tblPresentTime.Emp_ID"
    'CONVERT (nvarchar(50),GenPresentTime ,111)
    'StrSQL = StrSQL + " Where CONVERT (nvarchar(50),GenPresentTime ,101) >=" & SQLDate(Me.DtpFrom.Value, True) & " AND " & _
     " CONVERT (nvarchar(50),GenPresentTime ,101) <=" & SQLDate(Me.DtpTO.Value, True)
    StrSQL = StrSQL + " Where Month(GenPresentTime)=" & IntMonth & " AND Year(GenPresentTime)=" & IntYear & ""
    StrSQL = StrSQL + " Group By dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    IntDefWorkHours = val(Me.TxtMonthHours.text)

    If IntDefWorkHours = 0 Then Exit Sub

    Y = ConvertHoursToMints(IntDefWorkHours & ":00")

    With Me.Grid
        .cell(flexcpText, .FixedRows, .ColIndex("DefWorkHours"), .rows - 1, .ColIndex("DefWorkHours")) = IntDefWorkHours

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("WorkHours").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = rs("WorkHours").value
                    Z = ConvertHoursToMints(rs("WorkHours").value)
                    X = Z - Y

                    If X < 0 Then
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "-" & ConvertMintsToHours(Abs(X))
                    Else
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = ConvertMintsToHours(Abs(X))
                    End If
                
                    If InStr(1, .TextMatrix(LngFindRow, .ColIndex("OverTime")), "-", vbTextCompare) <> 0 Then
                        .cell(flexcpForeColor, LngFindRow, .ColIndex("OverTime")) = vbRed
                    End If

                Else
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = "00:00"
                    .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "00:00"
                End If
            End If

            rs.MoveNext
        Next i

    End With

End Sub

Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim ColumnName As String
    Dim SngTotal As Double
    Dim j As Integer
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        If .FixedRows = .rows Then Exit Sub

        For i = .FixedRows To .rows - 1
            '     SngHourPrice = Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) / Val(.TextMatrix(i, .ColIndex("DefWorkHours")))
            '     If .TextMatrix(i, .ColIndex("OverTime")) <> "" Then
            '         SngTemp = ConvertHoursToMints(.TextMatrix(i, .ColIndex("OverTime")))
            '         SngTemp = SngTemp * (1 / 60)
            '         SngOverTimePrice = SngTemp * SngHourPrice
            '         .TextMatrix(i, .ColIndex("OverTimePrice")) = SngOverTimePrice
            '         If SngOverTimePrice < 0 Then
            '             .Cell(flexcpForeColor, i, .ColIndex("OverTimePrice")) = vbRed
            '         End If
            '     End If

            TotalAddtion = 0
            TotalDiscount = 0

            For j = 1 To 40
                ColumnName = "Comp" & j

                If AddOrDiscount(j) = 0 Then
                    TotalAddtion = TotalAddtion + val(.TextMatrix(i, .ColIndex(ColumnName)))
                Else
                 '   If Not MofrdAbcen(j) Then
                        TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                 '   End If
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion + val(.TextMatrix(i, .ColIndex("TotalVacValue")))
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("PrePaidvalue"))) + val(.TextMatrix(i, .ColIndex("ToalInsurance"))) + val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + val(.TextMatrix(i, .ColIndex("VoCation3"))) + TotalDiscount
            
           ' If i <> .rows - 1 Then
                If val(.TextMatrix(i, .ColIndex("RemainDay"))) <> 0 Then
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2"))) ' + val(.TextMatrix(i, .ColIndex("TotalVacValue")))
                Else
                    
                    If val(.TextMatrix(i, .ColIndex("VacDay"))) + val(.TextMatrix(i, .ColIndex("AbcentDay"))) = daysInMonth(DTP_Date.value) Then
                        If val(.TextMatrix(i, .ColIndex("VacDay"))) = daysInMonth(DTP_Date.value) Then
                            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("TotalVacValue")))
                            '.TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2"))) '- val(.TextMatrix(i, .ColIndex("TotalVacValue")))
                        End If
                        
                            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2"))) '- val(.TextMatrix(i, .ColIndex("TotalVacValue")))
                        
                    Else
                        .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))
                    End If
                    
                End If
                If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) = 0 Then
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2"))) + val(.TextMatrix(i, .ColIndex("TotalVacValue")))
                End If
           ' End If

            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
  
            
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMySetting
    rsBranch.Close
    RsDepartment.Close
End Sub

Private Sub Grid_Click()
 
     Static lNoteRow&, lNoteCol&, r&, c&

    With Me.Grid
 
        r = .Row
        c = .Col

        If .ColKey(c) = "Emp_Name" And .TextMatrix(r, .ColIndex("Emp_ID")) <> "" Then
            FrmEmployee.show
            FrmEmployee.Retrive val(.TextMatrix(r, .ColIndex("Emp_ID")))
        End If
    
    End With
    
End Sub

Private Sub Grid_StartPage(ByVal hDC As Long, _
                           ByVal Page As Long, _
                           Cancel As Boolean)
    Dim s As String

    s = "„— »«  «·„ÊŸ›Ì‰ - Page " & Page & " - " & Now
    TextOut hDC, 100, 100, s, Len(s)
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    Me.lbl(14).Caption = Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Function Calculate_TotalSelected() As Long
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
    If GRID1.rows = 1 Then Exit Function
    Calculate_TotalSelected = 0

    For i = 1 To GRID1.rows - 1
        
        If GRID1.cell(flexcpChecked, i, GRID1.ColIndex("payed")) = flexChecked Then
            
            Calculate_TotalSelected = Calculate_TotalSelected + val(GRID1.TextMatrix(i, GRID1.ColIndex("EmpTotalNet")))
'            branchs_nos = val(Grid1.TextMatrix(i, Grid1.ColIndex("EmpTotalNet"))) + "," + branchs_nos


        End If

    Next i
   FillGridWithData3
End Function

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With GRID1
  
        If .ColKey(Col) = "payed" Then
            Cancel = False
        Else
            Cancel = True
        End If
         
    End With

End Sub

Private Sub Grid1_DblClick()
     Static lNoteRow&, lNoteCol&, r&, c&

    With Me.GRID1
 
        r = .Row
        c = .Col

        If .TextMatrix(r, .ColIndex("Emp_ID")) <> "" Then
            FrmEmployee.show
            FrmEmployee.Retrive val(.TextMatrix(r, .ColIndex("Emp_ID")))
        End If
    
    End With
End Sub

Private Sub Grid2_DblClick()
     Static lNoteRow&, lNoteCol&, r&, c&

    With Me.Grid2
 
        r = .Row
        c = .Col

        If .TextMatrix(r, .ColIndex("Emp_ID")) <> "" Then
            FrmEmployee.show
            FrmEmployee.Retrive val(.TextMatrix(r, .ColIndex("Emp_ID")))
        End If
    
    End With
End Sub

Private Sub GridGE_Click()
    With GridGE

        Select Case .Col

 

            Case 8
                ShowGL_cc val(.TextMatrix(.Row, .ColIndex("NoteSerial"))), , 200

        End Select

    End With
    
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub ISButton2_Click()
    'FillGridWithData

    'DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

My_SQL = " SELECT      dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(dbo.TblEmployee.Emp_ID)) AS PrePaidvalue,"

My_SQL = My_SQL + "                          HourRate =("
My_SQL = My_SQL + "                          SELECT top 1"
    
My_SQL = My_SQL + "                      ISNULL(TblChangedComponentRegisterDetails.HourRate, 0) HourRate"
   
My_SQL = My_SQL + "                      From TblChangedComponentRegisterDetails"
My_SQL = My_SQL + "                      INNER JOIN TblChangedComponentRegister"
My_SQL = My_SQL + "                          ON TblChangedComponentRegisterDetails.ChangedComponentid = TblChangedComponentRegister.ChangedComponentid"
My_SQL = My_SQL + "                      Where IsNull(TblChangedComponentRegisterDetails.HourRate, 0) <> 0"
My_SQL = My_SQL + "                      and TblChangedComponentRegister.Actualyear = " & val(Me.CboYear.text) & "  And TblChangedComponentRegister.Actualmonth = " & val(CmbMonth.ListIndex + 1)
My_SQL = My_SQL + "                      and TblChangedComponentRegisterDetails.Emp_id = TblEmployee.Emp_ID"
My_SQL = My_SQL + "                      )"

My_SQL = My_SQL + "                          ,NoOfHour =("
My_SQL = My_SQL + "                          SELECT top 1"
    
My_SQL = My_SQL + "                      ISNULL(TblChangedComponentRegisterDetails.HourRate, 0) NoOfHour"
   
My_SQL = My_SQL + "                      From TblChangedComponentRegisterDetails"
My_SQL = My_SQL + "                      INNER JOIN TblChangedComponentRegister"
My_SQL = My_SQL + "                          ON TblChangedComponentRegisterDetails.ChangedComponentid = TblChangedComponentRegister.ChangedComponentid"
My_SQL = My_SQL + "                      Where IsNull(TblChangedComponentRegisterDetails.NoOfHour, 0) <> 0"
My_SQL = My_SQL + "                      and TblChangedComponentRegister.Actualyear = " & val(Me.CboYear.text) & "  And TblChangedComponentRegister.Actualmonth = " & val(CmbMonth.ListIndex + 1)
My_SQL = My_SQL + "                      and TblChangedComponentRegisterDetails.Emp_id = TblEmployee.Emp_ID"
My_SQL = My_SQL + "                      ),"


My_SQL = My_SQL + "                          valueHour =("
My_SQL = My_SQL + "                          SELECT top 1"
    
My_SQL = My_SQL + "                      ISNULL(TblChangedComponentRegisterDetails.value, 0) NoOfHour"
   
My_SQL = My_SQL + "                      From TblChangedComponentRegisterDetails"
My_SQL = My_SQL + "                      INNER JOIN TblChangedComponentRegister"
My_SQL = My_SQL + "                          ON TblChangedComponentRegisterDetails.ChangedComponentid = TblChangedComponentRegister.ChangedComponentid"
My_SQL = My_SQL + "                      Where IsNull(TblChangedComponentRegisterDetails.value, 0) <> 0"
My_SQL = My_SQL + "                      and TblChangedComponentRegister.Actualyear = " & val(Me.CboYear.text) & "  And TblChangedComponentRegister.Actualmonth = " & val(CmbMonth.ListIndex + 1)
My_SQL = My_SQL + "                      and TblChangedComponentRegisterDetails.Emp_id = TblEmployee.Emp_ID"
My_SQL = My_SQL + "                      ),"

My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeName AS HJobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee AS HJobTypeNameE,"
My_SQL = My_SQL + "                      dbo.TblBranchesData.branch_name AS Hbranch_name, dbo.TblBranchesData.branch_namee AS Hbranch_nameH, dbo.TblEmployee.GroupID,"
My_SQL = My_SQL + "                      dbo.TblEmployee.BankCode, dbo.TblEmployee.BankCard, dbo.TblEmployee.NumEkama, dbo.TblEmployee.ContractID, dbo.TblEmployee.BranchId AS HBranchId,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1 AS HEmp_Name1, dbo.TblEmployee.Emp_Name2,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Namee,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.projects.Project_name,"
My_SQL = My_SQL + "                      dbo.projects.Project_nameE, dbo.projects.Fullcode AS ProjectFullcode, dbo.TblEmployee.NationalityE, dbo.TblEmployee.BignDateWork, dbo.Contract.Contract_date,"
My_SQL = My_SQL + "                      dbo.Contract.DateH, dbo.Contract.Contract_Enddate, dbo.Contract.DateH1, dbo.emp_salary.*, dbo.TblEmployee.EmpNotes, dbo.TblEmployee.kafeladd,"
My_SQL = My_SQL + "                      dbo.TblEmployee.DOB, dbo.TblEmployee.SpecificationID, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmpSpecifications.SpecificationNameE,"
My_SQL = My_SQL + "                      dbo.TblEmployee.PayType, dbo.TblEmployee.SalaryCode, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
My_SQL = My_SQL + "                      dbo.TblEmployee.NationlID, dbo.Nationality.name AS Nationalname, dbo.Nationality.namee AS NationalnameE, dbo.TblEmployee.DeptID2,"
My_SQL = My_SQL + "                      dbo.TblEmpDepartmentsDet.Name AS DepartmentName2, dbo.TblEmpDepartmentsDet.NameE AS DepartmentName2E, dbo.TblEmployee.dean,"
My_SQL = My_SQL + "                      dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.KafelID, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
My_SQL = My_SQL + "                      dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.KafelName, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.pasplace,"
My_SQL = My_SQL + "                      dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket,"
My_SQL = My_SQL + "                      dbo.TblEmployee.dateendpoket, dbo.TblEmployee.kafeltel, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.InsuranceNO,"
My_SQL = My_SQL + "                      dbo.TblEmployee.DOBH, dbo.TblEmployee.IssueDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.LastDateH, dbo.TblEmployee.VisaNo,"
My_SQL = My_SQL + "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.Sex,"
My_SQL = My_SQL + "                      dbo.TblEmployee.MaritalStatus, dbo.TblEmployee.BankIAddress, dbo.TblEmployee.BanckName, dbo.TblEmployee.BankIBan, dbo.TblEmployee.SafEBox,"
My_SQL = My_SQL + "                      dbo.TblEmployee.HowIqamaStH , dbo.TblEmployee.HowIqamaEndH, dbo.TblEmployee.TypeEmp, dbo.TblEmployee.MachinCode"
My_SQL = My_SQL + " FROM         dbo.TblEmpSpecifications RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.Nationality RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmpDepartmentsDet RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmployee ON dbo.TblEmpDepartmentsDet.ID = dbo.TblEmployee.DeptID2 ON dbo.Nationality.id = dbo.TblEmployee.NationlID ON"
My_SQL = My_SQL + "                      dbo.TblEmpSpecifications.SpecificationID = dbo.TblEmployee.SpecificationID LEFT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.projects RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.Contract RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmpDepartments RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.emp_salary ON dbo.TblEmpDepartments.DeparmentID = dbo.emp_salary.DepartmentID ON dbo.Contract.Emp_id = dbo.emp_salary.emp_id ON"
My_SQL = My_SQL + "                      dbo.projects.id = dbo.emp_salary.project_id ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  "



    If Dcdep.BoundText <> "" And Dcdep.text <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and emp_salary.DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DcEmp.BoundText <> "" And Me.DcEmp.text <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_salary.emp_id=" & val(Me.DcEmp.BoundText)
    End If
        If cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
          My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex) & " "
    End If
     If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "

    End If
    If val(Me.DcbDepartment2.BoundText) <> 0 And Me.DcbDepartment2.text <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and dbo.TblEmployee.DeptID2=" & val(Me.DcbDepartment2.BoundText)
    End If
    
   If Me.DcbTeam.BoundText <> "" And Me.DcbTeam.text <> "" Then
        My_SQL = My_SQL + "  and TblEmployee.SpecificationID=" & val(Me.DcbTeam.BoundText)
    End If
    
       If Me.dcempcontract.BoundText <> "" And Me.dcempcontract.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.ContractID=" & val(Me.dcempcontract.BoundText)
    End If
 
        If Me.dcBranch1.BoundText <> "" And Me.dcBranch1.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.BranchId=" & val(Me.dcBranch1.BoundText)
    End If
 
 
        If Me.DCGroupID.BoundText <> "" And Me.DCGroupID.text <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.GroupID=" & val(Me.DCGroupID.BoundText)
    End If
  
  My_SQL = My_SQL + "  order by TblEmployee.Fullcode"
 '
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim StrFileName As String
StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"
  If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Dim str As String
    
    If dcBranch1.text <> "" Then
    str = "«·›—⁄ : " & dcBranch1.text & CHR(13)
    End If
     
        If DCGroupID.text <> "" Then
    str = str & CHR(13) & "«·„Êﬁ⁄ : " & DCGroupID.text & CHR(13)
    End If
      
        If dcproject.text <> "" Then
    str = str & CHR(13) & "«·„‘—Ê⁄ : " & dcproject.text & CHR(13)
    End If
            
           If Dcdep.text <> "" Then
    str = str & CHR(13) & "«·«œ«—… : " & Dcdep.text & CHR(13)
    End If
      
     
           If dcempcontract.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «· ⁄«ﬁœ : " & dcempcontract.text & CHR(13)
    End If
           
           If DcEmp.text <> "" Then
    str = str & CHR(13) & "«·„ÊŸ› : " & DcEmp.text & CHR(13)
    End If
     If cboPayType.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «·”œ«œ : " & cboPayType.text & CHR(13)
    End If
   If DcbHemiaSalary.text <> "" Then
    str = str & CHR(13) & "ﬂÊœ Õ„«Ì… «·«ÃÊ— : " & DcbHemiaSalary.text & CHR(13)
    End If
           
    xReport.ParameterFields(6).AddCurrentValue str
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.dcproject.BoundText <> "" Then
            '   xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & dcproject.text
            Else
            '   xReport.ParameterFields(48).AddCurrentValue "  " & dcproject.text
            End If
       xReport.ParameterFields(48).AddCurrentValue "  " '& dcproject.text
       
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = StrFileName
    FrmReport.CRViewer.viewReport
    FrmReport.TXTSTRSQL = My_SQL
    
    ' FrmReport.Show
  'xxxx
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Sub

Private Sub ISButton3_Click()
    'Form3.Show
    'Form3.case_id = 11
    FillGridWithData

    DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

    If Dcdep.BoundText <> "" Then
        My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
    Else
        My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    End If
    
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT11.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.show
    FrmReport.txtPath = App.path & "\reports\emp\REPORT11.rpt"
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"
End Sub

Private Sub ISButton4_Click()
      Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

'    My_SQL = " SELECT     *"
'    My_SQL = My_SQL & " FROM         dbo.emp_salary INNER JOIN"
'    My_SQL = My_SQL & "  dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'    My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"


'My_SQL = "SELECT     *"
'My_SQL = My_SQL & "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID"
 
'1 My_SQL = "SELECT     *"
'1 My_SQL = My_SQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.emp_salary ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
 My_SQL = " SELECT     TOP 100 PERCENT dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(dbo.TblEmployee.Emp_ID)) AS PrePaidvalue,"
 My_SQL = My_SQL & "                      dbo.TblEmpJobsTypes.JobTypeName AS HJobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee AS HJobTypeNameE,"
 My_SQL = My_SQL & "                     dbo.TblBranchesData.branch_name AS Hbranch_name, dbo.TblBranchesData.branch_namee AS Hbranch_nameH, dbo.TblEmployee.GroupID,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.BankCode, dbo.TblEmployee.BankCard, dbo.TblEmployee.NumEkama, dbo.TblEmployee.ContractID, dbo.TblEmployee.BranchId AS HBranchId,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1 AS HEmp_Name1, dbo.TblEmployee.Emp_Name2,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Namee,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.projects.Project_name,"
 My_SQL = My_SQL & "                     dbo.projects.Project_nameE, dbo.projects.Fullcode AS ProjectFullcode, dbo.TblEmployee.NationalityE, dbo.TblEmployee.BignDateWork, dbo.Contract.Contract_date,"
 My_SQL = My_SQL & "                     dbo.Contract.DateH, dbo.Contract.Contract_Enddate, dbo.Contract.DateH1, dbo.emp_salary.*, dbo.TblEmployee.EmpNotes, dbo.TblEmployee.kafeladd,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.DOB, dbo.TblEmployee.SpecificationID, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmpSpecifications.SpecificationNameE,"
 My_SQL = My_SQL & "                     dbo.TblEmployee.PayType,dbo.TblEmployee.SalaryCode"
 My_SQL = My_SQL & " FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.TblEmpSpecifications RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.TblEmployee ON dbo.TblEmpSpecifications.SpecificationID = dbo.TblEmployee.SpecificationID ON"
 My_SQL = My_SQL & "                     dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.projects RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.Contract RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.emp_salary ON dbo.Contract.Emp_id = dbo.emp_salary.emp_id ON dbo.projects.id = dbo.emp_salary.project_id ON"
 My_SQL = My_SQL & "                     dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
 My_SQL = My_SQL & "  WHERE     (dbo.emp_salary.sgn ='" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "') AND dbo.emp_salary.payed=1"
'My_SQL = My_SQL & " ORDER BY dbo.TblEmployee.Fullcode"
'My_SQL = My_SQL & " where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
'My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  "
'1 My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.Text & CmbMonth.ListIndex + 1 & "')  AND (payed =1) "
      '   If cboPayType.Text <> "" And val(cboPayType.ListIndex) <> -1 Then
      '    My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex) & " "
    'End If

    If Dcdep.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and emp_salary.DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DcEmp.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_salary.emp_id=" & val(Me.DcEmp.BoundText)
    End If

       If Me.dcempcontract.BoundText <> "" Then
        My_SQL = My_SQL + "  and TblEmployee.ContractID=" & val(Me.dcempcontract.BoundText)
    End If
       If Me.cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
         
        My_SQL = My_SQL + "  and TblEmployee.PayType=" & val(Me.cboPayType.ListIndex)
    End If
  If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "
    End If
        If Me.dcBranch1.BoundText <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.BranchId=" & val(Me.dcBranch1.BoundText)
    End If
 
 
        If Me.DCGroupID.BoundText <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.GroupID=" & val(Me.DCGroupID.BoundText)
    End If
  
  My_SQL = My_SQL + "  order by TblEmployee.Fullcode"
 '
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim StrFileName As String
StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"
  If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Dim str As String
    
    If dcBranch1.text <> "" Then
    str = "«·›—⁄ : " & dcBranch1.text & CHR(13)
    End If
     
        If DCGroupID.text <> "" Then
    str = str & CHR(13) & "«·„Êﬁ⁄ : " & DCGroupID.text & CHR(13)
    End If
      
        If dcproject.text <> "" Then
    str = str & CHR(13) & "«·„‘—Ê⁄ : " & dcproject.text & CHR(13)
    End If
            
           If Dcdep.text <> "" Then
    str = str & CHR(13) & "«·ﬁ”„ : " & Dcdep.text & CHR(13)
    End If
      
     
           If dcempcontract.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «· ⁄«ﬁœ : " & dcempcontract.text & CHR(13)
    End If
           
     If DcEmp.text <> "" Then
    str = str & CHR(13) & "«·„ÊŸ› : " & DcEmp.text & CHR(13)
    End If
     If cboPayType.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «·”œ«œ : " & cboPayType.text & CHR(13)
    End If
    If DcbHemiaSalary.text <> "" Then
    str = str & CHR(13) & "ﬂÊœ Õ„«Ì… «·«ÃÊ— : " & DcbHemiaSalary.text & CHR(13)
    End If

    xReport.ParameterFields(6).AddCurrentValue str
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.dcproject.BoundText <> "" Then
            '   xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & dcproject.text
            Else
            '   xReport.ParameterFields(48).AddCurrentValue "  " & dcproject.text
            End If
       xReport.ParameterFields(48).AddCurrentValue "  " '& dcproject.text
       
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = StrFileName
    FrmReport.CRViewer.viewReport
    ' FrmReport.Show
  
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
     Sendkeys "{RIGHT}"

End Sub

Private Sub ISButton5_Click()
    FillGridWithData

    DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

    If Dcdep.BoundText <> "" Then
        My_SQL = "SELECT * from emp_salary where payed=1 and m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
    Else
        My_SQL = "SELECT * from emp_salary where payed=1 and m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    End If

    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT11.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.show
    FrmReport.txtPath = App.path & "\reports\emp\REPORT11.rpt"
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Sub

Private Sub ISButton6_Click()
      Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

'    My_SQL = " SELECT     *"
'    My_SQL = My_SQL & " FROM         dbo.emp_salary INNER JOIN"
'    My_SQL = My_SQL & "  dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'    My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"


'My_SQL = "SELECT     *"
'My_SQL = My_SQL & "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID"
 
'1 My_SQL = "SELECT     *"
'1 My_SQL = My_SQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.emp_salary ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
'1 My_SQL = My_SQL & "                        dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
                      
'My_SQL = My_SQL & " where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
'My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  "
'1 My_SQL = My_SQL + "   WHERE     (sgn = '" & Me.CboYear.Text & CmbMonth.ListIndex + 1 & "')  AND (payed =0) "

My_SQL = " SELECT      dbo.EmpPrePaymentValue(dbo.EmpPrePaymentID(dbo.TblEmployee.Emp_ID)) AS PrePaidvalue,"
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeName AS HJobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee AS HJobTypeNameE,"
My_SQL = My_SQL + "                      dbo.TblBranchesData.branch_name AS Hbranch_name, dbo.TblBranchesData.branch_namee AS Hbranch_nameH, dbo.TblEmployee.GroupID,"
My_SQL = My_SQL + "                      dbo.TblEmployee.BankCode, dbo.TblEmployee.BankCard, dbo.TblEmployee.NumEkama, dbo.TblEmployee.ContractID, dbo.TblEmployee.BranchId AS HBranchId,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1 AS HEmp_Name1, dbo.TblEmployee.Emp_Name2,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Namee,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.projects.Project_name,"
My_SQL = My_SQL + "                      dbo.projects.Project_nameE, dbo.projects.Fullcode AS ProjectFullcode, dbo.TblEmployee.NationalityE, dbo.TblEmployee.BignDateWork, dbo.Contract.Contract_date,"
My_SQL = My_SQL + "                      dbo.Contract.DateH, dbo.Contract.Contract_Enddate, dbo.Contract.DateH1, dbo.emp_salary.*, dbo.TblEmployee.EmpNotes, dbo.TblEmployee.kafeladd,"
My_SQL = My_SQL + "                      dbo.TblEmployee.DOB, dbo.TblEmployee.SpecificationID, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmpSpecifications.SpecificationNameE,"
My_SQL = My_SQL + "                      dbo.TblEmployee.PayType,dbo.TblEmployee.SalaryCode"
My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmpSpecifications RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmployee ON dbo.TblEmpSpecifications.SpecificationID = dbo.TblEmployee.SpecificationID ON"
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.projects RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.Contract RIGHT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.emp_salary ON dbo.Contract.Emp_id = dbo.emp_salary.emp_id ON dbo.projects.id = dbo.emp_salary.project_id ON"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL + " WHERE     (dbo.emp_salary.sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "') AND (dbo.emp_salary.payed = 0)"

    If Dcdep.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and emp_salary.DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DcEmp.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_salary.emp_id=" & val(Me.DcEmp.BoundText)
    End If

    
       If Me.dcempcontract.BoundText <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.ContractID=" & val(Me.dcempcontract.BoundText)
    End If
 
        If Me.dcBranch1.BoundText <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.BranchId=" & val(Me.dcBranch1.BoundText)
    End If
 
 
        If Me.DCGroupID.BoundText <> "" Then
         
        My_SQL = My_SQL + "  and TblEmployee.GroupID=" & val(Me.DCGroupID.BoundText)
    End If
           If Me.cboPayType.text <> "" And val(cboPayType.ListIndex) <> -1 Then
        My_SQL = My_SQL + "  and TblEmployee.PayType=" & val(Me.cboPayType.ListIndex)
    End If
    If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode=N'" & DcbHemiaSalary.text & "' "
    End If
    
  My_SQL = My_SQL + "  order by TblEmployee.Fullcode"
 '
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim StrFileName As String
StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"
  If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Dim str As String
    
    If dcBranch1.text <> "" Then
    str = "«·›—⁄ : " & dcBranch1.text & CHR(13)
    End If
     
        If DCGroupID.text <> "" Then
    str = str & CHR(13) & "«·„Êﬁ⁄ : " & DCGroupID.text & CHR(13)
    End If
      
        If dcproject.text <> "" Then
    str = str & CHR(13) & "«·„‘—Ê⁄ : " & dcproject.text & CHR(13)
    End If
            
           If Dcdep.text <> "" Then
    str = str & CHR(13) & "«·ﬁ”„ : " & Dcdep.text & CHR(13)
    End If
      
     
           If dcempcontract.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «· ⁄«ﬁœ : " & dcempcontract.text & CHR(13)
    End If
           
           If DcEmp.text <> "" Then
    str = str & CHR(13) & "«·„ÊŸ› : " & DcEmp.text & CHR(13)
    End If
    If cboPayType.text <> "" Then
    str = str & CHR(13) & "‰Ê⁄ «·”œ«œ : " & cboPayType.text & CHR(13)
    End If
      If DcbHemiaSalary.text <> "" Then
    str = str & CHR(13) & "ﬂÊœ Õ„«Ì… «·«ÃÊ— : " & DcbHemiaSalary.text & CHR(13)
    End If
    xReport.ParameterFields(6).AddCurrentValue str
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.dcproject.BoundText <> "" Then
            '   xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & dcproject.text
            Else
            '   xReport.ParameterFields(48).AddCurrentValue "  " & dcproject.text
            End If
       xReport.ParameterFields(48).AddCurrentValue "  " '& dcproject.text
       
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = StrFileName
    FrmReport.CRViewer.viewReport
    ' FrmReport.Show
  
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
   Sendkeys "{RIGHT}"

End Sub

Private Sub TxtMonthHours_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMonthHours.text, 1)
End Sub
Private Sub GetAdvanceValues(IntMonth As Integer, _
                             IntYear As Integer)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.Grid
        rs.MoveFirst
        'Sayemen707 .Cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = Round(rs("CCC").value, 0)
                End If
            End If

            rs.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub

Private Sub GetEmpBalance()
If SystemOptions.ShowBalanceOfEmpInSalary = True Then
    Dim i As Integer
    Dim Balance As String
    Dim StrTempAccountCode As String
   ToDate.value = DateAdd("d", -1, DTP_Date)
    With Me.Grid
        For i = 1 To .rows - 1
          If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 Then
                    StrTempAccountCode = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")     'C?C??? C???E??E
                    WriteCustomerBalPublic StrTempAccountCode, Balance, , , , , , , ToDate, 1
                    .TextMatrix(i, .ColIndex("EmpBalance")) = Balance
                    .TextMatrix(i, .ColIndex("EmpReaminB")) = val(.TextMatrix(i, .ColIndex("EmpBalance"))) - val(.TextMatrix(i, .ColIndex("TotalAdvance")))
           End If
        Next i
    End With
End If
End Sub

Private Sub TxtSearchCode_KeyUp(KeyCode As Integer, Shift As Integer)
Dim EmpID As Integer
  If KeyCode = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        Me.DcEmp.BoundText = EmpID
    End If
End Sub

