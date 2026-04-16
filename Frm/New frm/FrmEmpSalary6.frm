VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalary6 
   BackColor       =   &H0000FF00&
   Caption         =   "»Ì«‰«  «·”œ«œ"
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   HelpContextID   =   580
   Icon            =   "FrmEmpSalary6.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11400
   ScaleWidth      =   15120
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
      Height          =   11400
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15120
      _cx             =   26670
      _cy             =   20108
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
      GridRows        =   2
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmEmpSalary6.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   9630
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   14925
         _cx             =   26326
         _cy             =   16986
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
         Caption         =   "»Ì«‰«  «·”œ«œ| ’œÌ— ··»‰ﬂ|»Ì«‰«  «·”œ«œ"
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
         Picture(0)      =   "FrmEmpSalary6.frx":03E0
         Flags(0)        =   2
         Picture(1)      =   "FrmEmpSalary6.frx":077A
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9165
            Index           =   1
            Left            =   -15780
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   45
            Width           =   14835
            _cx             =   26167
            _cy             =   16166
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
               Height          =   855
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   0
               Visible         =   0   'False
               Width           =   1275
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
               Height          =   750
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   -105
               Visible         =   0   'False
               Width           =   4500
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
                  MICON           =   "FrmEmpSalary6.frx":0B14
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
               Height          =   645
               Left            =   9330
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   150
               Width           =   5505
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
                  MICON           =   "FrmEmpSalary6.frx":0B30
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
                  Format          =   241565699
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
                  Visible         =   0   'False
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
               Height          =   435
               Left            =   2895
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   180
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CheckBox Check16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ÊﬁÌ⁄"
               Height          =   285
               Left            =   -105
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.CheckBox Check15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ì"
               Height          =   285
               Left            =   1170
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.CheckBox Check14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì 2"
               Height          =   285
               Left            =   1935
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.CheckBox Check13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ã“«¡« "
               Height          =   285
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”·›"
               Height          =   285
               Left            =   3735
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.CheckBox Check11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì1"
               Height          =   285
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.CheckBox Check10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ê·« "
               Height          =   285
               Left            =   5025
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.CheckBox Check9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬂ«›√ "
               Height          =   285
               Left            =   5820
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox Check8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷«›Ì"
               Height          =   300
               Left            =   7290
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   105
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.CheckBox Check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ—Ï"
               Height          =   285
               Left            =   7215
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.CheckBox Check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ⁄«„"
               Height          =   285
               Left            =   8235
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " „Ê«’·« "
               Height          =   285
               Left            =   8820
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ· ”ﬂ‰"
               Height          =   285
               Left            =   9630
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—« » «”«”Ì"
               Height          =   285
               Left            =   10620
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   420
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   240
               Left            =   11820
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   1035
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2760
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„ÊŸ›"
               Height          =   225
               Left            =   13155
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   1500
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   7995
               Left            =   -4005
               TabIndex        =   60
               Top             =   765
               Width           =   5685
               _cx             =   10028
               _cy             =   14102
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
               Cols            =   65
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEmpSalary6.frx":0B4C
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
               Height          =   330
               Left            =   4875
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   150
               Width           =   3330
            End
            Begin VB.Image ImgFavorites 
               Height          =   495
               Left            =   4980
               Picture         =   "FrmEmpSalary6.frx":133F
               Stretch         =   -1  'True
               Top             =   150
               Width           =   405
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9165
            Index           =   2
            Left            =   -15480
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   14835
            _cx             =   26167
            _cy             =   16166
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
            Height          =   9165
            Index           =   4
            Left            =   45
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   45
            Width           =   14835
            _cx             =   26167
            _cy             =   16166
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
            Begin VB.CheckBox Check22 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   330
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   585
               Width           =   1050
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               ClipControls    =   0   'False
               Height          =   1935
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   -150
               Width           =   9285
               Begin VB.ComboBox cboPayType 
                  Height          =   315
                  ItemData        =   "FrmEmpSalary6.frx":4FA7
                  Left            =   3000
                  List            =   "FrmEmpSalary6.frx":4FA9
                  TabIndex        =   125
                  Top             =   1125
                  Width           =   1665
               End
               Begin MSDataListLib.DataCombo Dcemp2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   108
                  Top             =   765
                  Width           =   4905
                  _ExtentX        =   8652
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
               Begin MSDataListLib.DataCombo dcproject2 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   109
                  Top             =   885
                  Width           =   2865
                  _ExtentX        =   5054
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
               Begin MSDataListLib.DataCombo Dcdep2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   110
                  Top             =   435
                  Width           =   4905
                  _ExtentX        =   8652
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCGroupID2 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   111
                  Top             =   525
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcBranch2 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   112
                  Top             =   150
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcempcontract2 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   124
                  Top             =   1245
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbHemiaSalary 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   126
                  Top             =   1125
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ Õ„«Ì… «·«ÃÊ—"
                  Height          =   240
                  Index           =   25
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   1125
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·”œ«œ"
                  Height          =   345
                  Index           =   24
                  Left            =   3960
                  TabIndex        =   123
                  Top             =   1125
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊŸ› "
                  DataField       =   "Õœœ"
                  Height          =   360
                  Index           =   23
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   780
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·„‘—Ê⁄"
                  Height          =   360
                  Index           =   22
                  Left            =   7980
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   885
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   255
                  Index           =   21
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   435
                  Width           =   1050
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
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
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   120
                  Width           =   2820
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   360
                  Index           =   20
                  Left            =   7980
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   525
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   360
                  Index           =   19
                  Left            =   7980
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   150
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
                  Height          =   240
                  Index           =   18
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   1245
                  Width           =   930
               End
            End
            Begin VB.CommandButton Command3 
               Caption         =   " ’œÌ— »Ì«‰«  «·„ÊŸ›Ì‰"
               Height          =   465
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   150
               Width           =   1710
            End
            Begin VB.CommandButton Command4 
               Caption         =   " ’œÌ—  »Ì«‰«  «·„ ⁄ÂœÌ‰"
               Height          =   465
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   150
               Width           =   1710
            End
            Begin VB.CheckBox Check21 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   330
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   585
               Width           =   1290
            End
            Begin VB.CheckBox Check20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   330
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   585
               Width           =   1290
            End
            Begin VB.CheckBox Check19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ ﬂ· «·„ÊŸ›Ì‰"
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   585
               Width           =   1290
            End
            Begin VB.CheckBox Check18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   330
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   585
               Width           =   1290
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   3525
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   1440
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.ComboBox CboPaymentType 
               Height          =   315
               ItemData        =   "FrmEmpSalary6.frx":4FAB
               Left            =   3525
               List            =   "FrmEmpSalary6.frx":4FBE
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   885
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   360
               Left            =   13410
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   -465
               Width           =   990
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   7065
               TabIndex        =   35
               Top             =   -465
               Visible         =   0   'False
               Width           =   1905
               _ExtentX        =   3360
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
               Left            =   7065
               TabIndex        =   36
               Top             =   -465
               Width           =   1905
               _ExtentX        =   3360
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
               Height          =   420
               Left            =   3120
               TabIndex        =   42
               Top             =   105
               Visible         =   0   'False
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   741
               _Version        =   393216
               Format          =   241172481
               CurrentDate     =   39614
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   705
               Left            =   105
               TabIndex        =   46
               Top             =   0
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   1244
               BTYPE           =   2
               TX              =   "”œ«œ "
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
               MICON           =   "FrmEmpSalary6.frx":4FE5
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   390
               Left            =   10725
               TabIndex        =   64
               Top             =   -915
               Visible         =   0   'False
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   688
               _Version        =   393216
               Format          =   241172481
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DCAccount 
               Height          =   315
               Left            =   7065
               TabIndex        =   66
               Top             =   -465
               Width           =   1905
               _ExtentX        =   3360
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
               Height          =   540
               Left            =   0
               TabIndex        =   90
               Top             =   2340
               Visible         =   0   'False
               Width           =   7695
               _cx             =   13573
               _cy             =   952
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
               Cols            =   68
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":5001
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
            Begin ALLButtonS.ALLButton ALLButton6 
               Height          =   555
               Left            =   105
               TabIndex        =   94
               Top             =   0
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   979
               BTYPE           =   2
               TX              =   "”œ«œ "
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
               MICON           =   "FrmEmpSalary6.frx":5842
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
               Height          =   555
               Left            =   135
               TabIndex        =   100
               Top             =   0
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   979
               BTYPE           =   2
               TX              =   "”œ«œ "
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
               MICON           =   "FrmEmpSalary6.frx":585E
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
               Height          =   555
               Left            =   105
               TabIndex        =   102
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   979
               BTYPE           =   2
               TX              =   "”œ«œ "
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
               MICON           =   "FrmEmpSalary6.frx":587A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ImpulseButton.ISButton Cmdd 
               Height          =   405
               Left            =   -105
               TabIndex        =   104
               Top             =   885
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   714
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
               ButtonImage     =   "FrmEmpSalary6.frx":5896
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
               Height          =   6705
               Left            =   0
               TabIndex        =   98
               Top             =   2100
               Visible         =   0   'False
               Width           =   14760
               _cx             =   26035
               _cy             =   11827
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
               BackColorAlternate=   16777088
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":2F4B8
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
               Begin MSComctlLib.ProgressBar ProgressBar2 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   99
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
               Height          =   6825
               Left            =   0
               TabIndex        =   120
               Top             =   2040
               Visible         =   0   'False
               Width           =   14760
               _cx             =   26035
               _cy             =   12039
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
               BackColorAlternate=   16777088
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
               Cols            =   19
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":2F6AD
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
               Begin MSComctlLib.ProgressBar ProgressBar3 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   121
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid1 
               Height          =   6705
               Left            =   -30
               TabIndex        =   61
               Top             =   1920
               Visible         =   0   'False
               Width           =   14910
               _cx             =   26300
               _cy             =   11827
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
               Cols            =   72
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":2F955
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   6825
               Left            =   0
               TabIndex        =   91
               Top             =   1860
               Visible         =   0   'False
               Width           =   14760
               _cx             =   26035
               _cy             =   12039
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
               BackColorAlternate=   16777088
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
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":3025D
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
               Begin MSComctlLib.ProgressBar ProgressBar1 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   92
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
               Height          =   6840
               Left            =   7800
               TabIndex        =   101
               Top             =   2220
               Visible         =   0   'False
               Width           =   6930
               _cx             =   12224
               _cy             =   12065
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
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmEmpSalary6.frx":30551
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
            Begin VB.Shape Shape2 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Height          =   750
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   1185
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
               Height          =   345
               Index           =   14
               Left            =   1140
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   150
               Width           =   1230
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               Height          =   345
               Index           =   13
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·”œ«œ"
               Height          =   345
               Index           =   12
               Left            =   11685
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   -885
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
               Height          =   345
               Index           =   10
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   150
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   300
               Index           =   9
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1500
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ“Ì‰…"
               Height          =   345
               Index           =   8
               Left            =   8865
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   -585
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   345
               Index           =   7
               Left            =   4845
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   885
               Visible         =   0   'False
               Width           =   735
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1695
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   9675
         Visible         =   0   'False
         Width           =   15060
         _cx             =   26564
         _cy             =   2990
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
         Begin MSComDlg.CommonDialog CD 
            Left            =   3840
            Top             =   960
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1515
            Index           =   3
            Left            =   7605
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   135
            Width           =   2070
            _cx             =   3651
            _cy             =   2672
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
               Top             =   300
               Width           =   1770
            End
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   675
               Width           =   1770
            End
            Begin ImpulseButton.ISButton CmdOk 
               Height          =   405
               Left            =   90
               TabIndex        =   88
               Top             =   1065
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   714
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
               ButtonImage     =   "FrmEmpSalary6.frx":30669
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
               Top             =   2250
               Width           =   1770
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
               Top             =   2325
               Width           =   1770
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   465
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   510
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Õ–› «·ﬁÌœ"
            Enabled         =   0   'False
            Height          =   465
            Left            =   4455
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1050
            Visible         =   0   'False
            Width           =   1560
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
            ItemData        =   "FrmEmpSalary6.frx":30A03
            Left            =   90
            List            =   "FrmEmpSalary6.frx":30A1C
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   645
            Visible         =   0   'False
            Width           =   3105
         End
         Begin MSDataListLib.DataCombo Dcemp 
            Height          =   315
            Left            =   9795
            TabIndex        =   9
            Top             =   1350
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   480
            Left            =   30
            TabIndex        =   2
            Top             =   1260
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   847
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
            ButtonImage     =   "FrmEmpSalary6.frx":30A8F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   465
            Left            =   4455
            TabIndex        =   8
            Top             =   510
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   820
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
            ButtonImage     =   "FrmEmpSalary6.frx":30E29
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcproject 
            Height          =   315
            Left            =   14160
            TabIndex        =   11
            Top             =   1350
            Width           =   2985
            _ExtentX        =   5265
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
            Height          =   645
            Left            =   8895
            TabIndex        =   17
            Top             =   -630
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1138
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
            ButtonImage     =   "FrmEmpSalary6.frx":311C3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   465
            Left            =   9945
            TabIndex        =   18
            Top             =   -600
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   820
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
            ButtonImage     =   "FrmEmpSalary6.frx":3155D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   540
            Left            =   0
            TabIndex        =   32
            Top             =   -105
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   953
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
            MICON           =   "FrmEmpSalary6.frx":318F7
            PICN            =   "FrmEmpSalary6.frx":31913
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
            Height          =   480
            Left            =   105
            TabIndex        =   48
            Top             =   1860
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   847
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
            ButtonImage     =   "FrmEmpSalary6.frx":31DBF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   480
            Left            =   1020
            TabIndex        =   49
            Top             =   675
            Visible         =   0   'False
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   847
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
            ButtonImage     =   "FrmEmpSalary6.frx":32159
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   465
            Left            =   3225
            TabIndex        =   50
            Top             =   1620
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   820
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
            ButtonImage     =   "FrmEmpSalary6.frx":324F3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   465
            Left            =   6060
            TabIndex        =   72
            Top             =   1050
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   820
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
            MICON           =   "FrmEmpSalary6.frx":3288D
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
            Left            =   9795
            TabIndex        =   79
            Top             =   510
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCGroupID 
            Height          =   315
            Left            =   14145
            TabIndex        =   81
            Top             =   915
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcempcontract 
            Height          =   315
            Left            =   9795
            TabIndex        =   83
            Top             =   915
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcBranch1 
            Height          =   315
            Left            =   14145
            TabIndex        =   87
            Top             =   450
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·ÕÊŸ… : «÷€ÿ ⁄·Ï «”„ «·„ÊŸ› ·„‘«Âœ… „·›…"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   0
            Width           =   3450
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
            Height          =   270
            Index           =   17
            Left            =   13065
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   915
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   16
            Left            =   17490
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   450
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Êﬁ⁄"
            Height          =   270
            Index           =   15
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   915
            Width           =   1065
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
            Height          =   360
            Left            =   15675
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   105
            Width           =   2820
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ««·ﬁ”„"
            Height          =   600
            Index           =   4
            Left            =   12810
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ ‰„Ê–Ã"
            Height          =   285
            Index           =   6
            Left            =   3255
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   675
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„‘—Ê⁄"
            Height          =   420
            Index           =   5
            Left            =   17055
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   1350
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ÊŸ› „Õœœ"
            DataField       =   "Õœœ"
            Height          =   435
            Index           =   3
            Left            =   12945
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1365
            Width           =   1020
         End
      End
      Begin VB.Shape Shape1 
         Height          =   9630
         Left            =   30
         Top             =   30
         Width           =   14925
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
      ButtonImage     =   "FrmEmpSalary6.frx":328A9
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalary6"
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
Dim FixedOrChanged(40) As Integer
Dim AddOrDiscount(40) As Integer
Dim ViewComp(40) As Boolean
Dim Account_code(40) As String
Dim Account_code1(40) As String

Dim ZmamAccount(40) As String
Public ClearSalary As Boolean
Public ClearPayment As Boolean
Public ClearPayment1 As Boolean
Dim AdvPaymentdAccount(40) As String
Public Row1 As Long
Dim componentname(40) As String
Dim firstrun As Boolean
Dim rsBranch As New ADODB.Recordset
Public PayDes As String
Public OrderSupplerDes As String
Public OrderSupplerDes1 As String
Public PayDes1 As String

Public empDes As String
Public empDes1 As String
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
                .cell(flexcpBackColor, i, 1, i, 60) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 60) = vbWhite
            End If

        Next i

    End With

    With GRID1

        For i = .FixedRows To .rows - 2
        
            If i Mod 2 = 0 Then
     '           .Cell(flexcpBackColor, i, 1, i, 62) = &HFFFFC0
            Else
     '           .Cell(flexcpBackColor, i, 1, i, 62) = vbWhite
            End If

        Next i

    End With
 
     With GRID2

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

Function GetNotesSerials(year As String, Month As String, NoteType As Integer) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select NoteSerial from notes where salary=" & val(year) & Month & " and  NoteType=" & NoteType & " and branch_no=" & Current_branch
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        GetNotesSerials = ""
    Else
        GetNotesSerials = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
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
    If SystemOptions.UserInterface = ArabicInterface Then
     ss = "»Ì«‰ »«”„«¡ «·„ÊŸ›Ì‰ «·–Ì‰ ·œÌÂ„ „‘«ﬂ·  "
     Else
    ss = "Statement the names of employees who have problems "
    End If
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
                   If SystemOptions.UserInterface = ArabicInterface Then
                   error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡    ÕœÌœ «·›—⁄ «· «»⁄ ·Â"
        Else
         error_string = error_string + "  Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "It is not the creation of a branch"
       End If
        
                check_employee_accounts = False
                   End If
                   
                   
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")

            If Employee_account = "" Or (Employee_account) = Null Then
            If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» –„ …"
                Else
                 error_string = error_string + "  Employee No  :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "It is not created discharged Account"
            
                End If
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» –„ … ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
                Else
                 error_string = error_string + "  Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & "  Account has been deleted discharged from the accounts manually guide  " & vbCrLf
       
                End If
       
                check_employee_accounts = False
            End If
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
                    If Employee_account = "" Or (Employee_account) = Null Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» «·«ÃÊ— «·„” Õﬁ…"
                Else
                error_string = error_string + "   Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "Is not the creation wages due account"
        
                  End If
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
                Else
                error_string = error_string + "  Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & "  Manually delete the wages due account of the chart of accounts  " & vbCrLf
       End If
       
                check_employee_accounts = False
            End If
            
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3")
                    If Employee_account = "" Or (Employee_account) = Null Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«»   «·„œ›Ê⁄«  «·„ﬁœ„…"
        Else
         error_string = error_string + "   Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "Is not Create an account payments"
        
        End If
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«»    «·„œ›Ê⁄«  «·„ﬁœ„… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
                Else
                  error_string = error_string + "  Employee No :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Name " & .TextMatrix(i, .ColIndex("Emp_Name")) & "   Manually delete the foreground of payments accounts of the chart of accounts " & vbCrLf
       
                End If
       
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
    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("Â·  —Ìœ › Õ «·„·› ··„—«Ã⁄Â", vbCritical + vbYesNo, "ÌÊÃœ Œÿ√ ›Ì Õ”«»«  «·„ÊŸ›Ì‰  Ì„ﬂ‰ „—«Ã⁄ … ›Ì „·› «·«Œÿ«¡")
Else
     X = MsgBox("Do you want to open the file for review", vbCritical + vbYesNo, "There is an error in the accounts staff can review the error file")

End If
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
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
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
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
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

Function GetComponentValuePerBranch(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
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
    Dim CValue As Double
    Dim Branch As Integer
    Dim projectId As Integer
    
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
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
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
    
            If .TextMatrix(i, .ColIndex("total1")) > 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                
                
                
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("total1")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                    GoTo ErrTrap
                End If
                
                
                
                
                

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
                    If val(.TextMatrix(.rows - 1, .ColIndex(ColumnName))) > 0 Then
                    
                               Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        '
                       
                                  If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex(ColumnName))), 0, Msg & "  " & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
                            line_no = line_no + 1
                        
                        If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), val(.TextMatrix(i, .ColIndex(ColumnName))), 1, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
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
                       
             
                            
                            
                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                               Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
                                                 If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                            GoTo ErrTrap
                                        End If
        
                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                                                        
                                  
                                line_no = line_no + 1
                                
                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                            GoTo ErrTrap
                                        End If
        
                                line_no = line_no + 1
                            End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
            
            
            
                                      Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
                                                 If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 0, Msg & " ”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                            GoTo ErrTrap
                                        End If
                                        
                         line_no = line_no + 1
                                        
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
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
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                                                GoTo ErrTrap
                                                            End If
                        
                                                line_no = line_no + 1
                                            End If
                 
                 Else
                 
                 
                 
                 
                                           If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                           
                                           
                                                       
                                      Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
                                                 If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & " ”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
                                            GoTo ErrTrap
                                        End If
                                        
                                        line_no = line_no + 1
                                        
                                             Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                                StrAccountCode = Employee_account
                                     
                                     
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId")))) = False Then
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

Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId"
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
             projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
    rs("Note_Value").value = net_value

    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex + 1
    rs("NoteType").value = 66
    rs("NoteDate").value = DTP_Date.value
    rs("UserID").value = user_id
    rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    rs("sanad_year").value = year(DTP_Date.value)
    rs("sanad_month").value = Month(DTP_Date.value)
    rs.update
   ''/////////////////
   '''''////////////
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
    Dim BranchID As Integer
    Dim CValue As Double
    Dim Branch As Integer
    Dim projectId As Integer
    
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
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               
                        If CValue > 0 Then
                        Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
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
                             
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            Else
                            
                                     If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & "   Œ’Ê„«  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            
                            
End If
                            line_no = line_no + 1
                        End If
                                    
                        rsBranch.MoveNext
                    Next Branch
                                    
                End If
            End If
    
        Next j

        For i = .FixedRows To .rows - 2
    
            If .TextMatrix(i, .ColIndex("EmpTotalNet")) > 0 Then '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
     
                 If .TextMatrix(i, .ColIndex("EmpTotalNet")) < 0 Then '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(.TextMatrix(i, .ColIndex("EmpTotalNet"))), 0, Msg, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
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
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 Then '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
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
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                                                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
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
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId"
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
             projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
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

'sql = " SELECT  dbo.TblEmployee.Emp_ID,    dbo.TblEmployee.BranchId,  dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.EmpSalaryComponent.[Value], dbo.mofrad.Account_Code,"
'sql = sql & "                       dbo.MOFRAD.Account_code1"
'sql = sql & "  FROM         dbo.mofrad INNER JOIN"
'sql = sql & "                       dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN"
'sql = sql & "                       dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode INNER JOIN"
'sql = sql & "                       dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID"
'sql = sql & "  Where (dbo.MOFRAD.acc = 1)"
'sql = sql & "  AND dbo.EmpSalaryComponent.emp_ID in"
'sql = sql & "  ("
' sql = sql & "  Select  Emp_ID"
'sql = sql & "  From TblEmployee"
'sql = sql & "  Where dbo.TblEmployee.WorkState  in ( "
'sql = sql & "  SELECT     id"
'sql = sql & "  From dbo.jopstatus"
'sql = sql & " Where  Insurances = 1)"
'sql = sql & "  and dbo.TblEmployee.BignDateWork<" & SQLDate(DTP_Date.value, True) & " "
'sql = sql & " )"
'sql = sql & "  ORDER BY (  dbo.TblEmployee.fullcode)"
'
sql = "  SELECT     TOP 100 PERCENT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, " & Chr(13)
sql = sql & "                       dbo.TblEmployee.Emp_Name, SUM(dbo.EmpSalaryComponent.[Value]) AS value, dbo.mofrad.Account_Code, dbo.mofrad.Account_code1" & Chr(13)
sql = sql & "  FROM         dbo.mofrad INNER JOIN" & Chr(13)
                      sql = sql & "   dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN" & Chr(13)
sql = sql & "                         dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode INNER JOIN" & Chr(13)
                      sql = sql & "   dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID" & Chr(13)
sql = sql & "    WHERE     (dbo.mofrad.acc = 1) AND (dbo.EmpSalaryComponent.emp_ID IN" & Chr(13)
sql = sql & "                             (SELECT     Emp_ID" & Chr(13)
sql = sql & "                                From TblEmployee" & Chr(13)
sql = sql & "                                WHERE     dbo.TblEmployee.WorkState IN" & Chr(13)
sql = sql & "                                                          (SELECT     id" & Chr(13)
sql = sql & "                                                             From dbo.jopstatus" & Chr(13)
sql = sql & "                                                             WHERE     Insurances = 1) AND dbo.TblEmployee.BignDateWork <" & SQLDate(DTP_Date.value, True) & ")) AND" & Chr(13)
sql = sql & "                         (year(dbo.EmpSalaryComponent.EntIncresDataM)<year( " & SQLDate(DTP_Date.value, True) & ") OR" & Chr(13)
                      sql = sql & "   dbo.EmpSalaryComponent.EntIncresDataM IS NULL) AND (dbo.mofrad.acc = 1)" & Chr(13)
sql = sql & "   GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.EmpSalaryComponent.AccountName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name," & Chr(13)
sql = sql & "                         dbo.MOFRAD.Account_Code , dbo.MOFRAD.Account_code1, dbo.TblEmployee.fullcode" & Chr(13)
sql = sql & "    ORDER BY dbo.TblEmployee.Fullcode" & Chr(13)






    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     mofradAccount1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
     
     Balance = IIf(IsNull(rs("Value").value), 0, rs("Value").value)
      
     mofradname = IIf(IsNull(rs("AccountName").value), "", rs("AccountName").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     Emp_id = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
      
      emp_Name = IIf(IsNull(rs("emp_Name").value), "", rs("emp_Name").value)
                             If mofradAccount <> "" And mofradAccount1 <> "" And Balance > 0 Then
                                   
                                  
                                   If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "   √Ì‰« - " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID, , , , , , , , Emp_id) = False Then
                                            GoTo ErrTrap
                                        End If
                        
                                        line_no = line_no + 1
                                        
                                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount1, Balance, 1, Msg & mofradname & "  √Ì‰«  -" & "  " & emp_Name, val(notes_id), , , , DTP_Date.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , Emp_id) = False Then
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

Private Sub ALLButton2_Click()


'On Error Resume Next

   
    
    
    DCEmP.text = ""
    Dcdep.text = ""
    DCproject.text = ""

    'FillGridWithData
    DoEvents
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text
    DTP_Date.value = MonthLastDay(CDate(str))

    If Grid.rows = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Õœœ ‘Â— «Ê·«", vbCritical
        Else
            MsgBox " Specify Month Firstly", vbCritical
        End If
        
        Exit Sub
    End If

    If detect_employee_work_type = 1 Then



  If SystemOptions.ProjectEmployeeGV = True Then
        If Create_dev3 = False Then
                      Exit Sub
                End If
  Else
  
       If Create_dev2 = False Then
                      Exit Sub
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
    FillGridWithData2
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

    LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & Chr(13) & " «·‘Â—     " & CmbMonth.text & Chr(13) & "  «·”‰…   " & CboYear.text & Chr(13) & " «· «—ÌŒ " & DTP_Date.value
                     
    LogTexte = ""
       AddToLogFile CInt(user_id), 66, Date, Time, LogTextA, LogTexte, Me.Name, "N", "", , val(TxtNoteSerial), ""
       
 
    
    
End Sub

Private Sub ALLButton3_Click()
    On Error Resume Next

FrmPayments.XPTxtVal.text = (lbl(14).Caption)
FrmPayments.empDes.text = empDes
FrmPayments.empDes1.text = empDes1
FrmPayments.CboYear1.text = CboYear.text
FrmPayments.CmbMonth1.text = CmbMonth.text
Me.Hide


      Exit Sub
      
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
                total_value = total_value + Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), 2)
            
                depit_side = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_Code1")
                CURRENT_LINE = setfoxy_Line

                If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), 2), 0, Msg, val(notes_id), , , , DTPicker1.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_id")))) = False Then
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
        X = MsgBox(" „ «‰‘«¡ ”‰œ «·”œ«œ —ﬁ„ «·ﬁÌœ ÂÊ " & Chr(13) & notes_serial & " Â·  —Ìœ ⁄—÷ «·ﬁÌœ ‰⁄„ «„ ·«", vbInformation + vbYesNo)

    Else
        X = MsgBox("   Voucher Created " & Chr(13) & notes_serial & "  Show GE", vbInformation + vbYesNo)
    End If

    If X = vbYes Then
        ShowGL_cc notes_serial, , 200
    End If
        
        
            LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & Chr(13) & " «·‘Â—     " & CmbMonth.text & Chr(13) & "  «·”‰…   " & CboYear.text & Chr(13) & " «· «—ÌŒ " & DTP_Date.value
                     
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

Private Sub ALLButton6_Click()
    On Error Resume Next
FrmPayments.XPTxtVal.text = (lbl(14).Caption)
FrmPayments.PayDes.text = PayDes
FrmPayments.empDes1.text = PayDes1
FrmPayments.CalCulteVAT 1
Me.Hide
      Exit Sub
End Sub

Private Sub ALLButton7_Click()
    On Error Resume Next

FrmPayments.XPTxtVal.text = (lbl(14).Caption)
FrmPayments.TxtNoSupplerDes.text = OrderSupplerDes
Me.Hide
      Exit Sub
End Sub

Private Sub ALLButton8_Click()
'FrmPayments.XPTxtVal.text = (lbl(14).Caption)
With FrmPayments.GRID1
.TextMatrix(Row1, .ColIndex("StrQest")) = PayDes
.TextMatrix(Row1, .ColIndex("InstalValue")) = val((lbl(14).Caption))
FrmPayments.Reline2
End With
Me.Hide
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



Private Sub cboPayType_KeyPress(KeyAscii As Integer)
    If cboPayType.text = "" Or val(cboPayType.ListIndex) = -1 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
    FillGridWithData2
    End If
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

    Me.lbl(14).Caption = val(Calculate_TotalSelected) ' Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
 
 
 
 End Sub

Private Sub Check18_Click()
    Dim i As Integer

    If Check18.value = vbChecked Then

        With Me.VSFlexGrid1
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
    
            Me.lbl(14).Caption = val(Calculate_TotalSelected2)

End Sub

Private Sub Check19_Click()
    Dim i As Integer

    If Check19.value = vbChecked Then

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
    '    Me.lbl(14).Caption = val(Calculate_TotalSelected2)

    End If
       Me.lbl(14).Caption = val(Calculate_TotalSelected)
       
End Sub

Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = True

    End If

End Sub

Private Sub Check20_Click()
    Dim i As Integer

    If Check20.value = vbChecked Then

        With Me.VSFlexGrid2
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid2

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
    
            Me.lbl(14).Caption = val(Calculate_TotalSelected3)
End Sub

Private Sub Check21_Click()
    Dim i As Integer

    If Check21.value = vbChecked Then

        With Me.VSFlexGrid3
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid3

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
    
            Me.lbl(14).Caption = val(Calculate_TotalSelectedQest)
End Sub

Private Sub Check22_Click()
  Dim i As Integer

    If Check22.value = vbChecked Then

        With Me.VSFlexGrid4
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid4

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
    
            Me.lbl(14).Caption = val(Calculate_TotalSelected16)
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




Private Sub Cmdd_Click()
FrmPayments.XPTxtVal.text = (lbl(14).Caption)
Me.Hide
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

Private Sub CmdOk_Click()
    On Error Resume Next

'firstrun = False
     If getTitlesName = True Then
   
   End If
    
    
    If firstrun = True Then
 
 '       Exit Sub
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
  '  FillGridWithData
    Me.TxtNoteSerial.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 66)
    Me.TxtNoteSerial2.text = GetNotesSerials(CboYear.text, CmbMonth.ListIndex + 1, 555)

    DoEvents
    cProgress.StopProgess
    cProgress.FinishProgress
   
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
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
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
'
Next i
End With



End Sub

Function create_report_data()
    On Error Resume Next
    Dim StrSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim ColumnName As String
    
    
'      StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'" ' & " and Branchid=" & Current_branch
StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'" ' '& " and Branchid=" & Current_branch

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
            
 
            rs("Emp_Name").value = .TextMatrix(i, .ColIndex("Emp_Name"))

            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True Then
                    rs(ColumnName).value = val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If
    
            Next j

            ' rs("Emp_Salary").value = .TextMatrix(i, .ColIndex("Emp_Salary"))
            ' rs("Emp_Salary_sakn").value = .TextMatrix(i, .ColIndex("Emp_Salary_sakn"))
            ' rs("Emp_Salary_bus").value = .TextMatrix(i, .ColIndex("Emp_Salary_bus"))
            ' rs("Emp_Salary_food").value = .TextMatrix(i, .ColIndex("Emp_Salary_food"))
            ' rs("Emp_Salary_mob").value = .TextMatrix(i, .ColIndex("Emp_Salary_mob"))
            ' rs("Emp_Salary_mang").value = .TextMatrix(i, .ColIndex("Emp_Salary_mang"))
            ' rs("Emp_Salary_others").value = .TextMatrix(i, .ColIndex("Emp_Salary_others"))
            ' rs("OverTimePrice").value = .TextMatrix(i, .ColIndex("OverTimePrice"))
            rs("Mokafea").value = .TextMatrix(i, .ColIndex("Mokafea"))
            rs("TotalAdvance").value = .TextMatrix(i, .ColIndex("TotalAdvance"))
            rs("TotalDiscount").value = .TextMatrix(i, .ColIndex("TotalDiscount"))
            rs("SalesCom").value = .TextMatrix(i, .ColIndex("SalesCom"))
            rs("total1").value = .TextMatrix(i, .ColIndex("total1"))
            rs("total2").value = .TextMatrix(i, .ColIndex("total2"))
            rs("EmpTotalNet").value = .TextMatrix(i, .ColIndex("EmpTotalNet"))
            rs("m_year").value = CboYear.text
            rs("m_month").value = CmbMonth.text
            rs("DepartmentID").value = .TextMatrix(i, .ColIndex("dep"))
            rs("project_id").value = .TextMatrix(i, .ColIndex("project"))
            rs("sgn").value = CboYear.text & CmbMonth.ListIndex + 1
 
            ',,
    
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

Private Sub Combo1_Click()

    If Combo1.ListIndex > -1 Then
                         If Combo1.ListIndex = 0 Then
                             ISButton2_Click
                         ElseIf Combo1.ListIndex = 1 Then
                             ISButton3_Click
                         ElseIf Combo1.ListIndex = 2 Then
                             ISButton4_Click
                         ElseIf Combo1.ListIndex = 3 Then
                             ISButton5_Click
                         ElseIf Combo1.ListIndex = 4 Then
                             ISButton6_Click
                         
                        ElseIf Combo1.ListIndex = 5 Then
                             ShowReports (5)
   ElseIf Combo1.ListIndex = 6 Then
                             ShowReports (6)
                             
                        End If
    End If

End Sub
Function ShowReports(indexs As Integer)
Dim FileName As String

Select Case indexs
Case 5
    FileName = App.path & "\reports\emp\REPORT10project.rpt"
Case 6
    FileName = App.path & "\reports\emp\REPORT10emp.rpt"


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


My_SQL = "SELECT     dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblBranchesData.*, dbo.emp_salary.*, dbo.projects.Fullcode"
My_SQL = My_SQL & "    FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL & "    dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
My_SQL = My_SQL & "                          dbo.projects ON dbo.emp_salary.project_id = dbo.projects.id"
 
My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"

    If Dcdep.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DCEmP.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_id=" & val(Me.DCEmP.BoundText)
    End If

    '
        If Me.DCproject.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and project_id=" & val(Me.DCproject.BoundText)
    End If
    
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(FileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
    
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.DCproject.BoundText <> "" Then
               xReport.ParameterFields(48).AddCurrentValue " «·„‘—Ê⁄ : " & DCproject.text
            Else
               xReport.ParameterFields(48).AddCurrentValue "  " & DCproject.text
            End If
    Dim j As Integer
    Dim ColumnName As String

    For j = 1 To 40
        ColumnName = "Comp" & j
        xReport.ParameterFields(6 + j).AddCurrentValue componentname(j)
    
    Next j

    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
Select Case indexs
Case 5
    FrmReport.txtpath = FileName
Case 6
    FrmReport.txtpath = FileName

End Select
    FrmReport.CRViewer.viewReport
    ' FrmReport.Show
  
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Function
Private Sub Command1_Click()
    Dim X As Integer
    Dim Msg As String
    Dim StrSQL  As String

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




   StrSQL = "Delete  Notes  where salary=" & CboYear.text & CmbMonth.ListIndex + 1 '& " and Branch_no=" & Current_branch
        Cn.Execute StrSQL, , adExecuteNoRecords
       
        ' StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        StrSQL = "Delete   emp_salary where SGN='" & CboYear.text & CmbMonth.ListIndex + 1 & "'" '& " and BranchId=" & Current_branch
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        
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

    LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & Chr(13) & " «·‘Â—     " & CmbMonth.text & Chr(13) & "  «·”‰…   " & CboYear.text & Chr(13) & " «· «—ÌŒ " & DTP_Date.value
                     
    LogTexte = ""
       AddToLogFile CInt(user_id), 66, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtNoteSerial), ""
       
 
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\Payrolll.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
    Me.Grid.saveGrid StrFileName, flexFileExcel, True
    OpenFile StrFileName
End Sub

Private Sub Command3_Click()
 
 Dim StrFileName As String
 
        On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "PaymentEmp"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.GRID1.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
     
     
 End Sub

Private Sub Command4_Click()
 Dim StrFileName As String
 
        On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Payemnt"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.VSFlexGrid2.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
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
    CmdOk_Click
End Sub

 

Private Sub DcbHemiaSalary_KeyPress(KeyAscii As Integer)
    If DcbHemiaSalary.text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub DcBranch1_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcBranch1.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
End Sub

Private Sub DcBranch2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Dcbranch2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub dcdep_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If Dcdep.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If

End Sub

Private Sub Dcdep2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Dcdep2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyReturn Then
        CmdOk_Click
    End If

End Sub

Private Sub Dcemp2_KeyUp(KeyCode As Integer, Shift As Integer)
    If DCEmp2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub dcempcontract_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcempcontract.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
End Sub

Private Sub dcempcontract2_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcempcontract2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub DCGroupID_KeyUp(KeyCode As Integer, Shift As Integer)
    If DCGroupID.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
    
End Sub

Private Sub DCGroupID2_KeyUp(KeyCode As Integer, Shift As Integer)
    If DCGroupID2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
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

Function getTitlesName() As Boolean
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
   If DCproject.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        firstrun = False
        CmdOk_Click
    End If
    
    
End Sub

Private Sub dcproject2_KeyUp(KeyCode As Integer, Shift As Integer)
    If dcproject2.text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
    FillGridWithData2
    End If
End Sub

Private Sub DTP_Date_Change()
    TxtNoteSerial.text = ""
End Sub

Private Sub Form_Activate()
If FrmPayments.TxtModFlg.text = "N" Or FrmPayments.TxtModFlg.text = "R" Then
Cmdd.Visible = False
Else
Cmdd.Visible = True

End If

End Sub

Private Sub Form_Load()
    Dim My_SQL As String
 'C1Tab1.CurrTab = 0
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    firstrun = True
If FrmPayments.TxtModFlg.text = "N" Or FrmPayments.TxtModFlg.text = "R" Then
Cmdd.Visible = False
Else
Cmdd.Visible = True

End If

    DTP_Date.value = Date
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
    My_SQL = "select Emp_id,Emp_Name From TblEmployee  order by  Emp_Name"
    fill_combo DCEmP, My_SQL

    My_SQL = "select DeparmentID,DepartmentName From TblEmpDepartments  order by DepartmentName "
    fill_combo Dcdep, My_SQL

    My_SQL = " select id,Project_name from projects order by Project_name"
    fill_combo DCproject, My_SQL

    My_SQL = "SELECT  (branch_id) From TblBranchesData"
   
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCmboEmp, True
    Dcombos.GetEmployees Me.DCEmp2, True
    Set cSearchDCombo = New clsDCboSearch
    Set cSearchDCombo.Client = DCmboEmp
    Dcombos.GetBranches Me.Dcbranch2
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.dcBranch1
    Dcombos.GetEmpSalaryCode Me.DcbHemiaSalary
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetAccountingCodes Me.DCAccount
    Dcombos.GetEmpLocations Me.DCGroupID
    Dcombos.GetEmpDepartments Me.Dcdep2
    Dcombos.GetEmpLocations Me.DCGroupID2
    Dcombos.Getemp_Contract_type Me.dcempcontract
    Dcombos.Getemp_Contract_type Me.dcempcontract2
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
    lbl(11).Caption = "Date"
    lbl(12).Caption = "Date"
    lbl(17).Caption = "Payment Type"
 Check19.RightToLeft = False
 Check19.Caption = "Select All"
 Check21.Caption = "Select All"
 Check20.RightToLeft = False
 Check18.RightToLeft = False
 Check18.Caption = "Select All"
 Check20.Caption = "Select All"
 Check22.Caption = "Select All"
 Check22.RightToLeft = False
    Me.Caption = "Monthly Payroll"
    ALLButton2.Caption = "Salary Allocation JV"
    ALLButton3.Caption = "Salary Payment JV"
    ALLButton6.Caption = "Payment"
    ALLButton7.Caption = "Payment"
    Me.C1Tab1.TabCaption(0) = " Allocation "
    Me.C1Tab1.TabCaption(2) = " Payment"
ALLButton8.Caption = "Payment"
    Ele(3).Caption = "Select Date"
    lbl(0).Caption = "Month"
    lbl(2).Caption = "Year"
    Fra.Caption = "Work Hours"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Emp Name"
    lbl(4).Caption = "Departement"
    lbl(5).Caption = "Project"
    lbl(6).Caption = "Select Report"
    lbl(7).Caption = "Payment Type"
    lbl(8).Caption = "Box"
    lbl(9).Caption = "Cheque No."
    lbl(10).Caption = "Due Date"
Cmdd.Caption = "Exit"
lbl(19).Caption = "Branch"
lbl(21).Caption = "Department"
lbl(18).Caption = "Contract Type"
lbl(23).Caption = "Employee"
lbl(20).Caption = "Location"
lbl(22).Caption = "Project"
Command3.Caption = "Export Data Employee"
Command4.Caption = "Export Data Contractors"
    ALLButton1.Caption = "Change Screen"
    CmdPrint.Caption = "Print"
    CmdExit.Caption = "Exit"
    Command1.Caption = "Delete JL"
Label6.Caption = "Press Eneter"
    Check17.Caption = "Select All"
With VSFlexGrid3
.TextMatrix(0, .ColIndex("ch")) = "Select"
.TextMatrix(0, .ColIndex("Inst_No")) = "No. installment"
.TextMatrix(0, .ColIndex("Due_Date")) = "Date"
.TextMatrix(0, .ColIndex("Value")) = "Value"
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
End With
    With Me.Grid
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

        .TextMatrix(0, .ColIndex("total1")) = "Total Add. "
        .TextMatrix(0, .ColIndex("total2")) = "Total Discount. "

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

    End With
    With VSFlexGrid4
      .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("ch")) = "Select"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
    .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
    .TextMatrix(0, .ColIndex("MordValue")) = "Value"
    .TextMatrix(0, .ColIndex("RecDate1")) = "Date"
    .TextMatrix(0, .ColIndex("name")) = "Name"
    End With
        With VSFlexGrid2
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("ch")) = "Select"
    .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
    .TextMatrix(0, .ColIndex("CusName")) = "Name"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
    .TextMatrix(0, .ColIndex("RecordNo")) = "Record .No"
    .TextMatrix(0, .ColIndex("IBAN")) = "IBAN"
    .TextMatrix(0, .ColIndex("BoardNO")) = "Car "
    .TextMatrix(0, .ColIndex("net")) = "Value"
    
    End With
    
    With VSFlexGrid1
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("ch")) = "Select"
    '.TextMatrix(0, .ColIndex("name")) = ""
    .TextMatrix(0, .ColIndex("name")) = "Name"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
    .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
    .TextMatrix(0, .ColIndex("Account_Name1")) = "Acc Expenditure "
    .TextMatrix(0, .ColIndex("Account_Name")) = "Acc Payable "
    .TextMatrix(0, .ColIndex("Valu")) = "Value"
    .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
    Frame1.Caption = "JV Data"
    Label1.Caption = "JV NO."
    ALLButton4.Caption = "Print JV"
lbl(13).Caption = "Total"
    With Me.GRID1
  
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("payed")) = "Select"
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("Emp_id")) = "Emp.ID"
        .TextMatrix(0, .ColIndex("emp_code")) = "Emp.Code"
        .TextMatrix(0, .ColIndex("emp_name")) = "Emp.Name"
        .TextMatrix(0, .ColIndex("Mokafea")) = "Additional"
        .TextMatrix(0, .ColIndex("TotalAdvance")) = "Advances"
        .TextMatrix(0, .ColIndex("TotalDiscount")) = "Discounts"
        .TextMatrix(0, .ColIndex("SalesCom")) = "Sales Com."
        .TextMatrix(0, .ColIndex("NetValue")) = "Net "
        .TextMatrix(0, .ColIndex("OldValue")) = "Prepaid "
        .TextMatrix(0, .ColIndex("EmpTotalNet")) = "Paid Value "
        .TextMatrix(0, .ColIndex("RemainValue")) = "Remaining  "
        
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
        .TextMatrix(0, .ColIndex("GroupName1")) = "Location"
        .TextMatrix(0, .ColIndex("DepartmentName1")) = "Department"
        .TextMatrix(0, .ColIndex("Project_name1")) = "Project"
        .TextMatrix(0, .ColIndex("name1")) = "Contract Type"

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
 
Public Sub FillGridWithData()
    Dim i As Integer
    Dim j As Integer

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
    
    My_SQL = " Select  lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE" & Chr(13)
  My_SQL = My_SQL + "  From (" & Chr(13)

  My_SQL = My_SQL + "  SELECT     TOP 100 PERCENT dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & Chr(13)
  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE" & Chr(13)
  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & Chr(13)
  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & Chr(13)
  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & Chr(13)
  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & Chr(13)

        If DCEmP.text <> "" Then
            My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.Emp_id=" & val(DCEmP.BoundText) ' & "'"
        Else

            If Dcdep.text <> "" Then
    
                If DCproject.BoundText = "" Then
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & val(Dcdep.BoundText) & "'"
                Else
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & Dcdep.BoundText & "' and dbo.TblEmployee.project_id='" & Me.DCproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If DCproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.project_id='" & Me.DCproject.BoundText & "'"
                    Else
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                    End If
    
                Else
    
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                End If
            End If
        End If

        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork<" & SQLDate(DTP_Date.value, True)
   If val(DCGroupID.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and   dbo.TblEmployee.workstate=1 and dbo.TblEmployee.GroupID=" & val(DCGroupID.BoundText)
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
 
 
My_SQL = My_SQL + "  GROUP BY dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, " & Chr(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus," & Chr(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang," & Chr(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName," & Chr(13)
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name," & Chr(13)
My_SQL = My_SQL + "                      dbo.Projects.Project_nameE" & Chr(13)
My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Fullcode" & Chr(13)

My_SQL = My_SQL + "  )XTable"



 Else
        FrstDay = "1-" & CmbMonth.ListIndex + 1 & "-" & year(Date)
        LstDay = DateAdd("d", -1, "1-" & CmbMonth.ListIndex + 2 & "-" & year(Date))

        My_SQL = "select Emp_ID,Emp_Name,Emp_Salary ,sum(TotalDiscount) as TotalDiscount," & "sum(Mokafea) as Mokafea  From QryEmpAllValues where TransDate >=#" & Format(FrstDay, "mm/dd/yyyy") & "# and TransDate<=#" & Format(LstDay, "mm/dd/yyyy") & "# " & StrWhere & " GROUP BY Emp_ID, Emp_Name, " & "Emp_Salary  "
    End If





    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
Dim CountDays As Double
Dim countFlag As Double
Dim MonthDayNo  As Double

MonthDayNo = daysInMonth(DTP_Date.value)

            For i = 1 To .rows - 1
         countFlag = 0
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(rs.Fields("BignDateWork").value), "", rs.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(rs.Fields("lastHolidaydate").value), "", rs.Fields("lastHolidaydate").value)

           
           
           If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("BignDateWork"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("BignDateWork"))) Then
           'CountDays
           countFlag = 1
           CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("BignDateWork")), DTP_Date.value)
           .TextMatrix(i, .ColIndex("CountDays")) = CountDays + 1
           Else
           countFlag = 0
            .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
           End If
           
           If IsDate(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           
                      If year(DTP_Date.value) = year(.TextMatrix(i, .ColIndex("lastHolidaydate"))) And Month(DTP_Date.value) = Month(.TextMatrix(i, .ColIndex("lastHolidaydate"))) Then
           'CountDays
           countFlag = 1
           CountDays = DateDiff("D", .TextMatrix(i, .ColIndex("lastHolidaydate")), DTP_Date.value)
           .TextMatrix(i, .ColIndex("CountDays")) = CountDays + 1
           Else
           countFlag = 0
            .TextMatrix(i, .ColIndex("CountDays")) = MonthDayNo
           End If
           
           
           End If
           
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
                '                .TextMatrix(i, .ColIndex("Comp1")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", Round(rs.Fields("Emp_Salary").value, Decimal_Places))
            
                
                      If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeName").value), "", rs.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0

                For j = 1 To 40
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DTP_Date.value)
                                           If countFlag = 1 Then
                                           
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, 2)
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1)
                                                     
                        End If
                    End If
    
                Next j
    
                '   .TextMatrix(I, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), _
                    "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                '      .TextMatrix(i, .ColIndex("total1"))
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, Decimal_Places))
             
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Round(rs.Fields("TotalMokafea").value, Decimal_Places))
              
                rs.MoveNext
            
            Next

            rs.Close
        End If

        GetAdvanceValues IntMonth, IntYear
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
Function GetExchangReq(Optional ID As Double = 0) As String
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
GetExchangReq = ""
If ID <> 0 Then
sql = " SELECT      AllID"
sql = sql & " From dbo.TblExchangeRequest"
sql = sql & " Where (id = " & ID & ")"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
GetExchangReq = IIf(IsNull(Rs9("AllID").value), "", Rs9("AllID").value)
Else
GetExchangReq = ""
End If
End If
End Function
Sub FillGrid5(Optional AllID As String)
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset

'''/////////////////

'''///////////////
sql = " SELECT     dbo.TblAttributionContract.IDMC, dbo.TblAttributionContract.ProcessNo, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.Dif, "
sql = sql & "                      dbo.TblAttributionContract.Depend, dbo.TblAttributionContract.SchoolYear, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
sql = sql & "                      dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblCustemers.CusName,"
sql = sql & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.BankAddress,"
sql = sql & "                      dbo.TblCustemers.IBAN, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.RecordNo, dbo.TblCustemers.CustGID,"
sql = sql & "                      dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.Account_Code,"
sql = sql & "                      dbo.TblCustemers.CusID, dbo.TblAttributionContract.IDAC, dbo.TblAttributionInstallmentDivided.TotalValue, dbo.TblAttributionInstallmentDivided.ID,"
sql = sql & "                      dbo.TblAttributionInstallmentDivided.BoardNO , dbo.TblAttributionInstallmentDivided.PayMentPayed"
sql = sql & " FROM         dbo.TblAttributionInstallmentDivided RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblAttributionContract ON dbo.TblAttributionInstallmentDivided.IDAC = dbo.TblAttributionContract.IDAC LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID"
sql = sql & "   WHERE     (dbo.TblAttributionInstallmentDivided.ID IN ( " & AllID & "))"

''''''''''''''''''''''/////////
    


   
'****************
If OrderSupplerDes = "" Then OrderSupplerDes = 0

'My_SQL = My_SQL + "   or      (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
'***************
   
  
If FrmPayments.TxtModFlg.text = "N" Then
ALLButton7.Enabled = True
Check20.Enabled = True
sql = sql + " and (dbo.TblAttributionInstallmentDivided.PayMentPayed Is Null)"
   
   VSFlexGrid1.Editable = flexEDKbdMouse

ElseIf FrmPayments.TxtModFlg.text = "R" Then
  
  sql = sql + "   AND  (dbo.TblAttributionInstallmentDivided.PayMentPayed =1)  AND (dbo.TblAttributionInstallmentDivided.ID in(" & OrderSupplerDes & ")) "
          
         Check20.Enabled = False
          VSFlexGrid2.Editable = flexEDNone
ALLButton7.Enabled = False

ElseIf FrmPayments.TxtModFlg.text = "E" Then
Check20.Enabled = True
ALLButton7.Enabled = True
    If ClearPayment = True Then 'new
    sql = sql + " and (dbo.TblAttributionInstallmentDivided.PayMentPayed Is Null)"
    VSFlexGrid2.Editable = flexEDKbdMouse

    
    Else: 'View
  '   Sql = Sql + "       AND ((dbo.TblPripaidExpensesDet.PaymentPayed IS NULL) OR"
'Sql = Sql + "                      (dbo.TblPripaidExpensesDet.PaymentPayed = 0)) "
'sql = sql + "   AND  (dbo.TblPripaidExpensesDet.PaymentPayed =1)  AND (dbo.TblPripaidExpensesDet.ID in(" & PayDes & ")) "
 sql = sql + "   AND  (dbo.TblAttributionInstallmentDivided.PayMentPayed =1)  AND (dbo.TblAttributionInstallmentDivided.ID in(" & OrderSupplerDes & ")) "
          
                
          VSFlexGrid2.Editable = flexEDNone
    End If


End If
sql = sql & " order by  TblAttributionInstallmentDivided.IDAC,BoardNO"
'''''''''''////////////////////
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
With VSFlexGrid2
 .rows = 2
.rows = .rows + Rs8.RecordCount - 1
Rs8.MoveFirst

For k = .FixedRows To Rs8.RecordCount
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("IDAC")) = IIf(IsNull(Rs8("IDAC").value), "", Rs8("IDAC").value)

.TextMatrix(k, .ColIndex("InsID")) = IIf(IsNull(Rs8("ID").value), "", Rs8("ID").value)
.TextMatrix(k, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchID").value), "", Rs8("BranchID").value)
.TextMatrix(k, .ColIndex("CusID")) = IIf(IsNull(Rs8("CusID").value), 0, Rs8("CusID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(k, .ColIndex("CusName")) = IIf(IsNull(Rs8("CusName").value), "", Rs8("CusName").value)
Else
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(k, .ColIndex("CusName")) = IIf(IsNull(Rs8("CusNamee").value), "", Rs8("CusNamee").value)
End If
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(k, .ColIndex("RecordNo")) = IIf(IsNull(Rs8("RecordNo").value), "", Rs8("RecordNo").value)
.TextMatrix(k, .ColIndex("net")) = IIf(IsNull(Rs8("TotalValue").value), 0, Rs8("TotalValue").value)
.TextMatrix(k, .ColIndex("IBAN")) = IIf(IsNull(Rs8("BankIBAN").value), "", Rs8("BankIBAN").value)
.TextMatrix(k, .ColIndex("BoardNO")) = IIf(IsNull(Rs8("BoardNO").value), "", Rs8("BoardNO").value)

.TextMatrix(k, .ColIndex("Account_Code")) = IIf(IsNull(Rs8("Account_Code").value), "", Rs8("Account_Code").value)
Rs8.MoveNext
Next k
.AutoSize 0, .Cols - 1, False
End With
End If
Reline2
End Sub
Sub FillGrid6(Optional ID As Double = 0)
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset

sql = "SELECT     QestID, Ind, [Value], Due_Date, Remarks, Inst_No, FlgPaye"
sql = sql & " From dbo.TblQestFexed"
sql = sql & " Where (ind = " & ID & ") "

If PayDes = "" Then PayDes = 0

If FrmPayments.TxtModFlg.text = "N" Then
sql = sql + " and  FlgPaye Is Null"
   Check21.Enabled = True
    VSFlexGrid3.Editable = flexEDKbdMouse
ALLButton8.Enabled = True
ElseIf FrmPayments.TxtModFlg.text = "R" Then
  ALLButton8.Enabled = False
  Check21.Enabled = False
  sql = sql + "   AND  (FlgPaye =1)  AND (QestID in(" & PayDes & ")) "
          
         
          VSFlexGrid3.Editable = flexEDNone


ElseIf FrmPayments.TxtModFlg.text = "E" Then
Check21.Enabled = True
ALLButton8.Enabled = True
    If ClearPayment1 = True Then 'new
    sql = sql + "       AND FlgPaye Is Null "
    VSFlexGrid3.Editable = flexEDKbdMouse

    
    Else: 'View

sql = sql + "   AND  (FlgPaye =1)  AND (QestID in(" & PayDes & ")) "
                
          VSFlexGrid3.Editable = flexEDNone
    End If


End If
'''''''''''////////////////////
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
With VSFlexGrid3
.rows = .rows + Rs8.RecordCount - 1
Rs8.MoveFirst

For k = .FixedRows To Rs8.RecordCount
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("QestID")) = IIf(IsNull(Rs8("QestID").value), 0, Rs8("QestID").value)
.TextMatrix(k, .ColIndex("Ind")) = IIf(IsNull(Rs8("Ind").value), 0, Rs8("Ind").value)
.TextMatrix(k, .ColIndex("Value")) = IIf(IsNull(Rs8("Value").value), 0, Rs8("Value").value)
.TextMatrix(k, .ColIndex("Due_Date")) = IIf(IsNull(Rs8("Due_Date").value), "", Rs8("Due_Date").value)
.TextMatrix(k, .ColIndex("Remarks")) = IIf(IsNull(Rs8("Remarks").value), "", Rs8("Remarks").value)
.TextMatrix(k, .ColIndex("Inst_No")) = IIf(IsNull(Rs8("Inst_No").value), 0, Rs8("Inst_No").value)

Rs8.MoveNext
Next k
.AutoSize 0, .Cols - 1, False
End With
End If
RelineQest
End Sub
Sub FillGrid4()
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset

'''/////////////////

'''///////////////

sql = "SELECT     dbo.TblPripaidExpensesDet.ID, dbo.TblPripaidExpensesDet.Name, dbo.TblPripaidExpensesDet.NameE, dbo.TblPripaidExpensesDet.TypeExpens, "
sql = sql & "                       dbo.TblPripaidExpensesDet.EmpID, dbo.TblPripaidExpensesDet.HistoryDate, dbo.TblPripaidExpensesDet.FromDate, dbo.TblPripaidExpensesDet.ToDate,"
sql = sql & "                       dbo.TblPripaidExpensesDet.Valu, dbo.TblPripaidExpensesDet.Remark2, dbo.TblPripaidExpensesDet.Distribution, dbo.TblPripaidExpensesDet.ProofID,"
sql = sql & "                       dbo.TblPripaidExpensesDet.Paye, dbo.TblPripaidExpensesDet.Account_Code, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_Serial,"
sql = sql & "                       ACCOUNTS_1.Account_NameEng, dbo.TblPripaidExpensesDet.Account_Code1, ACCOUNTS_1.Account_Name AS ExpAccount_Name,"
sql = sql & "                       ACCOUNTS_1.Account_Serial AS ExpAccount_Serial, ACCOUNTS_1.Account_NameEng AS ExpAccount_NameE, dbo.TblEmployee.Emp_Name,"
sql = sql & "                       dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblPripaidExpensesDet.BranchID, dbo.TblPripaidExpensesDet.PaymentPayed,"
sql = sql & "                       dbo.TblPripaidExpensesDet.ProfExpID , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
sql = sql & "  FROM         dbo.TblPripaidExpensesDet LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblPripaidExpensesDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblPripaidExpensesDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
sql = sql & "                       dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblPripaidExpensesDet.Account_Code = ACCOUNTS_2.Account_Code"
sql = sql & "  Where (dbo.TblPripaidExpensesDet.Paye = 1)"

''''''''''''''''''''''/////////
    


   
'****************
If PayDes = "" Then PayDes = 0

'My_SQL = My_SQL + "   or      (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
'***************
   
  
If FrmPayments.TxtModFlg.text = "N" Then
sql = sql + "       AND ((dbo.TblPripaidExpensesDet.PaymentPayed IS NULL) OR"
sql = sql + "                      (dbo.TblPripaidExpensesDet.PaymentPayed = 0)) "
   Check18.Enabled = True
    VSFlexGrid1.Editable = flexEDKbdMouse
ALLButton6.Enabled = True
ElseIf FrmPayments.TxtModFlg.text = "R" Then
  
  sql = sql + "   AND  (dbo.TblPripaidExpensesDet.PaymentPayed =1)  AND (dbo.TblPripaidExpensesDet.ID in(" & PayDes & ")) "
          
         Check18.Enabled = False
          VSFlexGrid1.Editable = flexEDNone

ALLButton6.Enabled = False
ElseIf FrmPayments.TxtModFlg.text = "E" Then
ALLButton6.Enabled = True
Check18.Enabled = True

    If ClearPayment = True Then 'new
    sql = sql + "       AND ((dbo.TblPripaidExpensesDet.PaymentPayed IS NULL) OR"
sql = sql + "                      (dbo.TblPripaidExpensesDet.PaymentPayed = 0)) "
    VSFlexGrid1.Editable = flexEDKbdMouse

    
    Else: 'View
  '   Sql = Sql + "       AND ((dbo.TblPripaidExpensesDet.PaymentPayed IS NULL) OR"
'Sql = Sql + "                      (dbo.TblPripaidExpensesDet.PaymentPayed = 0)) "
sql = sql + "   AND  (dbo.TblPripaidExpensesDet.PaymentPayed =1)  AND (dbo.TblPripaidExpensesDet.ID in(" & PayDes & ")) "
                
          VSFlexGrid1.Editable = flexEDNone
    End If


End If
With VSFlexGrid1
.rows = 1
End With
'''''''''''////////////////////
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
With VSFlexGrid1
.rows = 3
.rows = .rows + Rs8.RecordCount - 1
Rs8.MoveFirst

For k = .FixedRows To Rs8.RecordCount
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("MainID")) = IIf(IsNull(Rs8("ID").value), "", Rs8("ID").value)
.TextMatrix(k, .ColIndex("BranchId")) = IIf(IsNull(Rs8("BranchID").value), "", Rs8("BranchID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("Name").value), "", Rs8("Name").value)
.TextMatrix(k, .ColIndex("Account_Name1")) = IIf(IsNull(Rs8("ExpAccount_Name").value), "", Rs8("ExpAccount_Name").value)
.TextMatrix(k, .ColIndex("Account_Name")) = IIf(IsNull(Rs8("Account_Name").value), "", Rs8("Account_Name").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("NameE").value), "", Rs8("NameE").value)
.TextMatrix(k, .ColIndex("Account_Name1")) = IIf(IsNull(Rs8("ExpAccount_NameE").value), "", Rs8("ExpAccount_NameE").value)
.TextMatrix(k, .ColIndex("Account_Name")) = IIf(IsNull(Rs8("Account_NameEng").value), "", Rs8("Account_NameEng").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
End If
.TextMatrix(k, .ColIndex("TypeExpens")) = IIf(IsNull(Rs8("TypeExpens").value), 0, Rs8("TypeExpens").value)
.TextMatrix(k, .ColIndex("Valu")) = IIf(IsNull(Rs8("Valu").value), 0, Rs8("Valu").value)
.TextMatrix(k, .ColIndex("Account_Code")) = IIf(IsNull(Rs8("Account_Code").value), "", Rs8("Account_Code").value)
.TextMatrix(k, .ColIndex("ExpAccount_Code")) = IIf(IsNull(Rs8("Account_Code1").value), "", Rs8("Account_Code1").value)
.TextMatrix(k, .ColIndex("Account_Serial")) = IIf(IsNull(Rs8("Account_Serial").value), "", Rs8("Account_Serial").value)
.TextMatrix(k, .ColIndex("Account_Serial1")) = IIf(IsNull(Rs8("ExpAccount_Serial").value), "", Rs8("ExpAccount_Serial").value)
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(k, .ColIndex("EmpID")) = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
.TextMatrix(k, .ColIndex("IDD")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
Rs8.MoveNext
Next k
.AutoSize 0, .Cols - 1, False
End With
End If
Reline
End Sub
Sub FillGrid16()
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid4.rows = 2
sql = "SELECT     dbo.TblApproveCompoYearDet.ID, dbo.TblApproveCompoYearDet.PaymentPayed, dbo.TblApproveCompoYearDet.MofrdID, dbo.mofrdat.mofrad_type, "
sql = sql & "                       dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrad.Account_Code, dbo.TblApproveCompoYearDet.EmpID, dbo.TblEmployee.Emp_Name,"
sql = sql & "                       dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblApproveCompoYearDet.BrnchID1, dbo.TblBranchesData.branch_name,"
sql = sql & "                       dbo.TblBranchesData.branch_namee, dbo.TblApproveCompoYearDet.DeptID, dbo.TblApproveCompoYearDet.ProjectID, dbo.TblApproveCompoYearDet.TypeMofrd,"
sql = sql & "                       dbo.TblApproveCompoYearDet.StFunction, dbo.TblApproveCompoYearDet.RecDate1, dbo.TblApproveCompoYearDet.MordValue,"
sql = sql & "                       dbo.TblApproveCompoYearDet.CompYerID"
sql = sql & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblApproveCompoYearDet ON dbo.TblBranchesData.branch_id = dbo.TblApproveCompoYearDet.BrnchID1 LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblApproveCompoYearDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.mofrad RIGHT OUTER JOIN"
sql = sql & "                       dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.TblApproveCompoYearDet.MofrdID = dbo.mofrdat.mofrad_code"
sql = sql & " where 1<>-1"
''''''''''''''''''''''/////////
'****************
If PayDes = "" Then PayDes = 0
If FrmPayments.TxtModFlg.text = "N" Then
sql = sql + "       AND ((dbo.TblApproveCompoYearDet.PaymentPayed IS NULL) OR"
sql = sql + "                      (dbo.TblApproveCompoYearDet.PaymentPayed = 0)) "
   Check22.Enabled = True
    VSFlexGrid4.Editable = flexEDKbdMouse
ALLButton6.Enabled = True
ElseIf FrmPayments.TxtModFlg.text = "R" Then
  
  sql = sql + "   AND  (dbo.TblApproveCompoYearDet.PaymentPayed =1)  AND (dbo.TblApproveCompoYearDet.ID in(" & PayDes & ")) "
          
         Check22.Enabled = False
          VSFlexGrid4.Editable = flexEDNone

ALLButton6.Enabled = False
ElseIf FrmPayments.TxtModFlg.text = "E" Then
ALLButton6.Enabled = True
Check22.Enabled = True

    If ClearPayment = True Then 'new
    sql = sql + "       AND ((dbo.TblApproveCompoYearDet.PaymentPayed IS NULL) OR"
sql = sql + "                      (dbo.TblApproveCompoYearDet.PaymentPayed = 0)) "
    VSFlexGrid1.Editable = flexEDKbdMouse
    Else:
sql = sql + "   AND  (dbo.TblApproveCompoYearDet.PaymentPayed =1)  AND (dbo.TblApproveCompoYearDet.ID in(" & PayDes & ")) "
                
          VSFlexGrid4.Editable = flexEDNone
    End If
End If
'''''''''''////////////////////
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
With VSFlexGrid4
.rows = .rows + Rs8.RecordCount - 1
Rs8.MoveFirst

For k = .FixedRows To Rs8.RecordCount
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("ID")) = IIf(IsNull(Rs8("ID").value), "", Rs8("ID").value)
.TextMatrix(k, .ColIndex("BrnchID1")) = IIf(IsNull(Rs8("BrnchID1").value), "", Rs8("BrnchID1").value)
.TextMatrix(k, .ColIndex("EmpID")) = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
.TextMatrix(k, .ColIndex("MofrdID")) = IIf(IsNull(Rs8("MofrdID").value), 0, Rs8("MofrdID").value)
.TextMatrix(k, .ColIndex("mofrad_type")) = IIf(IsNull(Rs8("mofrad_type").value), 0, Rs8("mofrad_type").value)
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(k, .ColIndex("DeptID")) = IIf(IsNull(Rs8("DeptID").value), 0, Rs8("DeptID").value)
.TextMatrix(k, .ColIndex("ProjectID")) = IIf(IsNull(Rs8("ProjectID").value), 0, Rs8("ProjectID").value)
.TextMatrix(k, .ColIndex("Account_Code")) = IIf(IsNull(Rs8("Account_Code").value), "", Rs8("Account_Code").value)
.TextMatrix(k, .ColIndex("TypeMofrd")) = IIf(IsNull(Rs8("TypeMofrd").value), 0, Rs8("TypeMofrd").value)
.TextMatrix(k, .ColIndex("MordValue")) = IIf(IsNull(Rs8("MordValue").value), 0, Rs8("MordValue").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("mofrad_name").value), "", Rs8("mofrad_name").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("mofrad_namee").value), "", Rs8("mofrad_namee").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
End If

.TextMatrix(k, .ColIndex("RecDate1")) = IIf(IsNull(Rs8("RecDate1").value), "", Rs8("RecDate1").value)
.TextMatrix(k, .ColIndex("CompYerID")) = IIf(IsNull(Rs8("CompYerID").value), "", Rs8("CompYerID").value)
Rs8.MoveNext
Next k
.AutoSize 0, .Cols - 1, False
End With
End If
Reline16
End Sub
Public Sub FillGridWithData3()
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
My_SQL = My_SQL + "   WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "


        If DCEmP.text <> "" Then
            My_SQL = My_SQL + "  and  emp_salary.emp_code='" & DCEmP.BoundText & "'"
        Else

            If Dcdep.text <> "" Then
    
                If DCproject.BoundText = "" Then
                    My_SQL = My_SQL + "  and  emp_salary.DepartmentID='" & Dcdep.BoundText & "'"
                Else
                    My_SQL = My_SQL + "   and  emp_salary.DepartmentID='" & Dcdep.BoundText & "' and  project_id='" & Me.DCproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If DCproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + "  and  emp_salary.project_id='" & Me.DCproject.BoundText & "'"
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
   
 
   
        If val(CboPayMentType.ListIndex) <> -1 Then
   
   My_SQL = My_SQL + "  and ( dbo.TblEmployee.PayType is null or  dbo.TblEmployee.PayType=" & val(CboPayMentType.ListIndex) & ")"
   End If
   
   
        If val(dcempcontract.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
   End If
   
   
   
        My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
        '  My_SQL = My_SQL + " order by   LPAD(Emp_code,6,'0') ASC"
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.GRID2
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
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
            
                    .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
                   ' .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
                     .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("NumEkama").value), "", rs.Fields("NumEkama").value)
                    
                             If Trim(.TextMatrix(i, .ColIndex("Emp_Code"))) = "" Then
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("NumPoket").value), "", rs.Fields("NumPoket").value)
                    End If
                    
                    
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)

            '        .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
            '        .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            '
            '        .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
            '        .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    .TextMatrix(i, .ColIndex("Emp_Namee")) = IIf(IsNull(rs.Fields("Emp_Namee").value), IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value), rs.Fields("Emp_Namee").value)
                    If Trim(.TextMatrix(i, .ColIndex("Emp_Namee"))) = "" Then
                    .TextMatrix(i, .ColIndex("Emp_Namee")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    End If
                    
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
                   .TextMatrix(i, .ColIndex("EmpTotalNet")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, 2))

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
Sub RetrivSalaryPayed(Optional EmpID As Double, Optional ByRef PaymentValue As Double, Optional ByRef netvalue As Double, Optional ByRef RemainValue As Double, Optional ByRef OldValue As Double)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     PaymentValue, NetValue, RemainValue, OldValue"
sql = sql & " From dbo.TblSalaryNotesPayment"
sql = sql & "  Where (EmpID = " & EmpID & ") And (YearID = " & val(CboYear.text) & ") And (MonthID = " & val(CmbMonth.ListIndex) & ") And (TransID = " & val(FrmPayments.XPTxtID.text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
PaymentValue = IIf(IsNull(Rs3("PaymentValue").value), 0, Rs3("PaymentValue").value)
netvalue = IIf(IsNull(Rs3("NetValue").value), 0, Rs3("NetValue").value)
RemainValue = IIf(IsNull(Rs3("RemainValue").value), 0, Rs3("RemainValue").value)
OldValue = IIf(IsNull(Rs3("OldValue").value), 0, Rs3("OldValue").value)
Else
OldValue = 0
RemainValue = 0
netvalue = 0
PaymentValue = 0
End If
End Sub
Function GetSalaryPayed(Optional EmpID As Double, Optional TransID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     SUM(PaymentValue) AS SumValue"
sql = sql & " From dbo.TblSalaryNotesPayment"
sql = sql & " Where (EmpID = " & EmpID & ") And (YearID = " & val(CboYear.text) & ") And (MonthID = " & val(CmbMonth.ListIndex) & ") and TransID<>" & TransID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetSalaryPayed = IIf(IsNull(Rs3("SumValue").value), 0, Rs3("SumValue").value)
Else
GetSalaryPayed = 0
End If
End Function
Public Sub FillGridWithData2()
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
    Dim netvalue As Double
    Dim OldValue As Double
    Dim RemainValue As Double
    Dim PaymentValue As Double
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

'My_SQL = "SELECT     *"
'My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
'My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'''''''''''''''My_SQL = My_SQL + "   WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

My_SQL = " SELECT   TblEmployee.BranchId AS EMPBRANCHID, TblEmployee.Emp_Name, *, dbo.EmpGroupDep.GroupName AS GroupName1, dbo.EmpGroupDep.Ename AS Ename1, dbo.TblEmpDepartments.DepartmentName AS DepartmentName1,"
My_SQL = My_SQL + "                       dbo.TblEmpDepartments.DepartmentNamee AS DepartmentNamee1, dbo.projects.Project_name AS Project_name1, dbo.projects.Project_nameE AS Project_nameE1,"
My_SQL = My_SQL + "                       dbo.emp_contract_type.name AS name1, dbo.emp_contract_type.NameE AS NameE1 , dbo.emp_salary.id AS IDEmp ,dbo.TblEmployee.SalaryCode"
My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.emp_contract_type ON dbo.TblEmployee.ContractID = dbo.emp_contract_type.id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.projects ON dbo.emp_salary.project_id = dbo.projects.id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID"
My_SQL = My_SQL + "                       "
My_SQL = My_SQL + "   WHERE     ( 1=1) "
        If DCEmP.text <> "" Then
            My_SQL = My_SQL + "  and  emp_salary.emp_code='" & DCEmP.BoundText & "'"
        Else

            If Dcdep.text <> "" Then
    
                If DCproject.BoundText = "" Then
                    My_SQL = My_SQL + "  and  emp_salary.DepartmentID='" & Dcdep.BoundText & "'"
                Else
                    My_SQL = My_SQL + "   and  emp_salary.DepartmentID='" & Dcdep.BoundText & "' and  project_id='" & Me.DCproject.BoundText & "'"
                End If

            Else

                If Dcdep.text = "" Then
    
                    If DCproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + "  and  emp_salary.project_id='" & Me.DCproject.BoundText & "'"
                    End If
    
                End If
            End If
        End If

  '      If SystemOptions.usertype <> UserAdminAll Then
  '          My_SQL = My_SQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
            
  '      End If
    If val(dcproject2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.emp_salary.project_id=" & val(dcproject2.BoundText)
   End If
    
      If val(DCGroupID2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.TblEmployee.GroupID=" & val(DCGroupID2.BoundText)
   End If
   
    If val(DCEmp2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.emp_salary.emp_id=" & val(DCEmp2.BoundText)
   End If
   
    If val(dcempcontract2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.TblEmployee.ContractID=" & val(dcempcontract2.BoundText)
   End If
   
    If val(Dcdep2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.emp_salary.DepartmentID=" & val(Dcdep2.BoundText)
   End If
      If val(Dcbranch2.BoundText) <> 0 Then
   
       My_SQL = My_SQL + " and dbo.emp_salary.BranchId=" & val(Dcbranch2.BoundText)
   End If
   
   
              If val(dcBranch1.BoundText) <> 0 Then
   
   My_SQL = My_SQL + " and dbo.emp_salary.BranchId=" & val(dcBranch1.BoundText)
   End If
   
 
   
        If val(CboPayMentType.ListIndex) <> -1 Then
   
   My_SQL = My_SQL + "  and ( dbo.TblEmployee.PayType is null or  dbo.TblEmployee.PayType=" & val(CboPayMentType.ListIndex) & ")"
   End If
   
   
    If val(cboPayType.ListIndex) <> -1 Then
   My_SQL = My_SQL + " and dbo.TblEmployee.PayType=" & val(cboPayType.ListIndex)
   End If
 If Me.DcbHemiaSalary.text <> "" And DcbHemiaSalary.BoundText <> "" Then
          My_SQL = My_SQL + " and dbo.TblEmployee.SalaryCode='" & DcbHemiaSalary.BoundText & "' "
    End If
    
     If val(dcempcontract.BoundText) <> 0 Then
   My_SQL = My_SQL + " and dbo.TblEmployee.ContractID=" & val(dcempcontract.BoundText)
   End If
'****************
If empDes = "" Then empDes = 0

'My_SQL = My_SQL + "   or      (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
'***************
   
  
If FrmPayments.TxtModFlg.text = "N" Then
My_SQL = My_SQL + " AND (payed =0)  AND     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')  "

    My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
    GRID1.Editable = flexEDKbdMouse
Check19.Enabled = True
ALLButton3.Enabled = True
ElseIf FrmPayments.TxtModFlg.text = "R" Then
ALLButton3.Enabled = False
  Check19.Enabled = False
' My_SQL = My_SQL + "   AND  (payed =1) and     (sgn = '" & Me.CboYear.Text & CmbMonth.ListIndex + 1 & "')"
 My_SQL = My_SQL + "   AND      (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
          
          My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
          GRID1.Editable = flexEDNone


ElseIf FrmPayments.TxtModFlg.text = "E" Then
ALLButton3.Enabled = True
Check19.Enabled = True
    If ClearSalary = True Then 'new
    My_SQL = My_SQL + "  AND     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "') AND  (payed =0) "
    'AND (payed =0)
    My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
    GRID1.Editable = flexEDKbdMouse

    
    Else: 'View
     My_SQL = My_SQL + "   AND  (payed =1) and     (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
      ' My_SQL = My_SQL + "   and     (sgn = '" & Me.CboYear.Text & CmbMonth.ListIndex + 1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
        My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "
          GRID1.Editable = flexEDNone
    End If


End If
Dim k As Integer
  Dim Emp_id As Double
        '  My_SQL = My_SQL + " order by   LPAD(Emp_code,6,'0') ASC"
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.GRID1
            .rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
               ' .Rows = rs.RecordCount + 1
               .rows = 1
                rs.MoveFirst
i = 0
                For k = 1 To rs.RecordCount
                OldValue = 0
                netvalue = 0
                      netvalue = IIf(IsNull(rs.Fields("EmpTotalNet").value), 0, Round(rs.Fields("EmpTotalNet").value, 2))
                      Emp_id = IIf(IsNull(rs.Fields("Emp_id").value), 0, rs.Fields("Emp_id").value)
                    If FrmPayments.TxtModFlg.text = "N" Or FrmPayments.TxtModFlg.text = "E" Then
                    OldValue = GetSalaryPayed(Emp_id, val(FrmPayments.XPTxtID.text))
                    
                    End If
                    
                  If netvalue <> OldValue And netvalue <> 0 Then
                  .rows = .rows + 1
                    i = i + 1
                    .TextMatrix(i, .ColIndex("Ser")) = i
                    .TextMatrix(i, .ColIndex("NetValue")) = netvalue
            
                    '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
                       "", Rs.Fields("ID").value)
            
          '  .TextMatrix(i, .ColIndex("payed")) = IIf(IsNull(rs.Fields("payed").value), .Cell(flexcpChecked, i, .ColIndex("payed")) = flexUnchecked, .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked)
            
             .TextMatrix(i, .ColIndex("payed")) = IIf(IsNull(rs.Fields("payed").value), 0, rs.Fields("payed").value)
              
              
                        If .TextMatrix(i, .ColIndex("payed")) = True Then
                .cell(flexcpBackColor, i, 1, i, 62) = &HFF00&
            Else
                .cell(flexcpBackColor, i, 1, i, 62) = vbWhite
            End If
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("IDEmp").value), "", rs.Fields("IDEmp").value)
           
                   ' .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("EMPBRANCHID").value), "", rs.Fields("EMPBRANCHID").value)
           'sa MsgBox .TextMatrix(i, .ColIndex("id"))
           .TextMatrix(i, .ColIndex("Emp_id")) = Emp_id
                    
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
                    .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            .TextMatrix(i, .ColIndex("TotalVacValue")) = IIf(IsNull(rs.Fields("TotalVacValue").value), "", rs.Fields("TotalVacValue").value)
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  .TextMatrix(i, .ColIndex("GroupName1")) = IIf(IsNull(rs.Fields("GroupName1").value), "", rs.Fields("GroupName1").value)
                  .TextMatrix(i, .ColIndex("DepartmentName1")) = IIf(IsNull(rs.Fields("DepartmentName1").value), "", rs.Fields("DepartmentName1").value)
                  .TextMatrix(i, .ColIndex("Project_name1")) = IIf(IsNull(rs.Fields("Project_name1").value), "", rs.Fields("Project_name1").value)
                  .TextMatrix(i, .ColIndex("name1")) = IIf(IsNull(rs.Fields("name1").value), "", rs.Fields("name1").value)
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  
                  Else
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  .TextMatrix(i, .ColIndex("GroupName1")) = IIf(IsNull(rs.Fields("Ename1").value), "", rs.Fields("Ename1").value)
                  .TextMatrix(i, .ColIndex("DepartmentName1")) = IIf(IsNull(rs.Fields("DepartmentNamee1").value), "", rs.Fields("DepartmentNamee1").value)
                  .TextMatrix(i, .ColIndex("Project_name1")) = IIf(IsNull(rs.Fields("Project_nameE1").value), "", rs.Fields("Project_nameE1").value)
                  .TextMatrix(i, .ColIndex("name1")) = IIf(IsNull(rs.Fields("NameE1").value), "", rs.Fields("NameE1").value)
                  End If

               
                    '            .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), _
                                 "", rs.Fields("Emp_Salary").value)
            
                    .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                    .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("Mokafea").value), "", Round(rs.Fields("Mokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                    .TextMatrix(i, .ColIndex("TotalAdvance")) = IIf(IsNull(rs.Fields("TotalAdvance").value), "", Round(rs.Fields("TotalAdvance").value))
            
                    .TextMatrix(i, .ColIndex("SalesCom")) = IIf(IsNull(rs.Fields("SalesCom").value), "", Round(rs.Fields("SalesCom").value))
                    
                    .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(rs.Fields("total1").value), "", Round(rs.Fields("total1").value, 2))
            
                    .TextMatrix(i, .ColIndex("total2")) = IIf(IsNull(rs.Fields("total2").value), "", Round(rs.Fields("total2").value, 2))
            
                    '.TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, 2))
                    .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, 2))
                    If val(.TextMatrix(i, .ColIndex("NetValue"))) < 0 Then
                      .TextMatrix(i, .ColIndex("NetValue")) = 0
                    End If
                    If FrmPayments.TxtModFlg.text = "N" Or FrmPayments.TxtModFlg.text = "E" Then
                    .TextMatrix(i, .ColIndex("OldValue")) = GetSalaryPayed(val(.TextMatrix(i, .ColIndex("Emp_id"))), val(FrmPayments.XPTxtID.text))
                    .TextMatrix(i, .ColIndex("RemainValue")) = val(.TextMatrix(i, .ColIndex("NetValue"))) - val(.TextMatrix(i, .ColIndex("OldValue")))
                     .TextMatrix(i, .ColIndex("RemainValue")) = Round(val(.TextMatrix(i, .ColIndex("RemainValue"))), 2)
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("RemainValue")))
                    Else
                    RetrivSalaryPayed val(.TextMatrix(i, .ColIndex("Emp_id"))), PaymentValue, netvalue, RemainValue, OldValue
                    .TextMatrix(i, .ColIndex("OldValue")) = Round(OldValue, 2)
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = Round(PaymentValue, 2)
                    '.TextMatrix(I, .ColIndex("NetValue")) = Round(netvalue, 2)
                    .TextMatrix(i, .ColIndex("RemainValue")) = Round(RemainValue, 2)
                    .TextMatrix(i, .ColIndex("payed")) = 1
                    End If

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
            
                    
            End If
            rs.MoveNext
                Next k

                rs.Close
            End If
    
            GetAdvanceValues IntMonth, IntYear
            GetWorkHours
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
    RelinSalaryPayed
'rs.Close
Set rs = Nothing
    Coloring
   ' FillGridWithData3
    
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
                    TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + TotalDiscount
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))

            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub
Sub Reline2()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid2
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                IntCounter = IntCounter + 1
           '     .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           End If
           Next i
  
    End With
  Me.lbl(14).Caption = val(Calculate_TotalSelected3)
End Sub
Sub Reline16()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid4
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                IntCounter = IntCounter + 1
           Sm = Sm + val(.TextMatrix(i, .ColIndex("MordValue")))
           End If
           Next i
  
    End With
   lbl(14).Caption = Sm
End Sub
Sub Reline()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                IntCounter = IntCounter + 1
           '     .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("Valu")))
           End If
           Next i
  
    End With
   lbl(14).Caption = Sm
End Sub
Sub RelineQest()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid3
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
         
           End If
           Next i
  
    End With
   lbl(14).Caption = val(Calculate_TotalSelectedQest)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FrmPayments.TxtModFlg.text = "N" Or FrmPayments.TxtModFlg.text = "E" Then
'MsgBox "«÷€ÿ “— «·”œ«œ ", vbCritical
Shape2.Visible = True
Cancel = True
Else
 Shape2.Visible = False

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmPayments.XPTxtVal.text = (lbl(14).Caption)

    SaveMySetting
    rsBranch.Close
End Sub

Private Sub Grid_Click()
 
     Static lNoteRow&, lNoteCol&, r&, c&

    With Me.Grid
 
        r = .row
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

Private Sub Grid1_AfterEdit(ByVal row As Long, _
                            ByVal Col As Long)
    'Me.lbl(14).Caption = Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    RelinSalaryPayed
      
End Sub
Sub RelinSalaryPayed()
Dim i As Integer
With GRID1
For i = 1 To .rows - 1
If GRID1.cell(flexcpChecked, i, GRID1.ColIndex("payed")) = flexChecked Then
If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("OldValue"))) > val(.TextMatrix(i, .ColIndex("NetValue"))) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„”œœÂ «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "The paid value is greater than the residual value"
End If
.TextMatrix(i, .ColIndex("EmpTotalNet")) = 0
Exit Sub
End If
If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) < 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… »«·”«·»"
Else
MsgBox "Value can not be negative"
End If
.TextMatrix(i, .ColIndex("EmpTotalNet")) = 0
Exit Sub
End If

End If
Next i
End With
Me.lbl(14).Caption = val(Calculate_TotalSelected)
End Sub

Function Calculate_TotalSelected3() As Double
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
OrderSupplerDes = ""
OrderSupplerDes1 = ""
    If VSFlexGrid2.rows = 1 Then Exit Function
    Calculate_TotalSelected3 = 0

    For i = 1 To VSFlexGrid2.rows
        
        If VSFlexGrid2.cell(flexcpChecked, i, VSFlexGrid2.ColIndex("ch")) = flexChecked And ((VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("InsID")))) <> "" Then
            
            Calculate_TotalSelected3 = Calculate_TotalSelected3 + val(VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("net")))

OrderSupplerDes = ((VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("InsID")))) + "," + OrderSupplerDes
OrderSupplerDes1 = (VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("InsID"))) + "#" + (VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("net"))) + "," + OrderSupplerDes1
        End If

    Next i
    If Len(OrderSupplerDes) > 0 Then
    OrderSupplerDes = mId(OrderSupplerDes, 1, Len(OrderSupplerDes) - 1)
    End If
    
        If Len(OrderSupplerDes1) > 0 Then
    OrderSupplerDes1 = mId(OrderSupplerDes1, 1, Len(OrderSupplerDes1) - 1)
    End If
End Function
Function Calculate_TotalSelectedQest() As Double
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
With VSFlexGrid3
PayDes = ""
PayDes1 = ""
    If .rows = 1 Then Exit Function
    Calculate_TotalSelectedQest = 0

    For i = 1 To .rows - 1
        
        If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked And val((.TextMatrix(i, .ColIndex("QestID")))) <> 0 Then
            
            Calculate_TotalSelectedQest = Calculate_TotalSelectedQest + val(.TextMatrix(i, .ColIndex("Value")))

PayDes = ((.TextMatrix(i, .ColIndex("QestID")))) + "," + PayDes
PayDes1 = (.TextMatrix(i, .ColIndex("QestID"))) + "#" + (.TextMatrix(i, .ColIndex("Value"))) + "," + PayDes1
        End If

    Next i
    If Len(PayDes) > 0 Then
    PayDes = mId(PayDes, 1, Len(PayDes) - 1)
    End If
    
        If Len(PayDes1) > 0 Then
    PayDes1 = mId(PayDes1, 1, Len(PayDes1) - 1)
    End If
 End With
End Function
Function Calculate_TotalSelected16() As Double
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
PayDes = ""
PayDes1 = ""
    If VSFlexGrid4.rows = 1 Then Exit Function
    Calculate_TotalSelected16 = 0

    For i = 1 To VSFlexGrid4.rows - 1
        If VSFlexGrid4.cell(flexcpChecked, i, VSFlexGrid4.ColIndex("ch")) = flexChecked And val(VSFlexGrid4.TextMatrix(i, VSFlexGrid4.ColIndex("BrnchID1"))) <> 0 Then
            
            Calculate_TotalSelected16 = Calculate_TotalSelected16 + val(VSFlexGrid4.TextMatrix(i, VSFlexGrid4.ColIndex("MordValue")))
'            branchs_nos = val(Grid1.TextMatrix(i, Grid1.ColIndex("EmpTotalNet"))) + "," + branchs_nos
PayDes = (VSFlexGrid4.TextMatrix(i, VSFlexGrid4.ColIndex("ID"))) + "," + PayDes
PayDes1 = (VSFlexGrid4.TextMatrix(i, VSFlexGrid4.ColIndex("CompYerID"))) + "," + PayDes1

        End If

    Next i
    If Len(PayDes) > 0 Then
    PayDes = mId(PayDes, 1, Len(PayDes) - 1)
    End If
    
        If Len(PayDes1) > 0 Then
    PayDes1 = mId(PayDes1, 1, Len(PayDes1) - 1)
    End If
    
 
End Function
Function Calculate_TotalSelected2() As Double
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
PayDes = ""
PayDes1 = ""
    If VSFlexGrid1.rows = 1 Then Exit Function
    Calculate_TotalSelected2 = 0

    For i = 1 To VSFlexGrid1.rows - 1
        
        If VSFlexGrid1.cell(flexcpChecked, i, VSFlexGrid1.ColIndex("ch")) = flexChecked And val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("BranchId"))) <> 0 Then
            
            Calculate_TotalSelected2 = Calculate_TotalSelected2 + val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("Valu")))
'            branchs_nos = val(Grid1.TextMatrix(i, Grid1.ColIndex("EmpTotalNet"))) + "," + branchs_nos
PayDes = (VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("MainID"))) + "," + PayDes
PayDes1 = (VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("MainID"))) + "#" + (VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("Valu"))) + "," + PayDes1

        End If

    Next i
    If Len(PayDes) > 0 Then
    PayDes = mId(PayDes, 1, Len(PayDes) - 1)
    End If
    
        If Len(PayDes1) > 0 Then
    PayDes1 = mId(PayDes1, 1, Len(PayDes1) - 1)
    End If
    
 
End Function
Function Calculate_TotalSelected() As Double
    Dim i As Integer
    On Error Resume Next
'Dim branchs_nos As String
empDes = ""
empDes1 = ""
    If GRID1.rows = 1 Then Exit Function
    Calculate_TotalSelected = 0

    For i = 1 To GRID1.rows - 1
        
        If GRID1.cell(flexcpChecked, i, GRID1.ColIndex("payed")) = flexChecked Then
            If val(GRID1.TextMatrix(i, GRID1.ColIndex("EmpTotalNet"))) > 0 Then
            Calculate_TotalSelected = Calculate_TotalSelected + val(GRID1.TextMatrix(i, GRID1.ColIndex("EmpTotalNet")))
            Else
            
            End If
'            branchs_nos = val(Grid1.TextMatrix(i, Grid1.ColIndex("EmpTotalNet"))) + "," + branchs_nos
empDes = (GRID1.TextMatrix(i, GRID1.ColIndex("Emp_id"))) + "," + empDes
empDes1 = (GRID1.TextMatrix(i, GRID1.ColIndex("Emp_id"))) + "#" + (GRID1.TextMatrix(i, GRID1.ColIndex("EmpTotalNet"))) + "#" + (GRID1.TextMatrix(i, GRID1.ColIndex("TotalVacValue"))) + "," + empDes1

        End If

    Next i
    If Len(empDes) > 0 Then
    empDes = mId(empDes, 1, Len(empDes) - 1)
    End If
    
        If Len(empDes1) > 0 Then
    empDes1 = mId(empDes1, 1, Len(empDes1) - 1)
    End If
    
    
  ' FillGridWithData3
   
End Function

Private Sub Grid1_BeforeEdit(ByVal row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With GRID1
  Select Case .ColKey(Col)
  Case "Payed"
  .ComboList = ""
  Case "project"
  Cancel = True
  Case "branch_name"
  Cancel = True
  Case "Emp_Code"
  Cancel = True
  Case "Emp_Name"
  Cancel = True
  Case "NetValue"
  Cancel = True
  Case "OldValue"
  Cancel = True
  Case "RemainValue"
  Cancel = True
  Case "EmpTotalNet"
  If GRID1.cell(flexcpChecked, row, GRID1.ColIndex("payed")) = flexChecked Then
  .ComboList = ""
  Else
  Cancel = True
  End If
  Case "DepartmentName1"
  Cancel = True
  Case "Project_name1"
  Cancel = True
  Case "GroupName1"
  Cancel = True
  Case "name1"
  Cancel = True
  End Select
         
    End With

End Sub

Private Sub Grid1_DblClick()
  '   Static lNoteRow&, lNoteCol&, r&, c&

  '  With Me.Grid1
 
  '      r = .Row
  '      c = .Col
'
'        If .TextMatrix(r, .ColIndex("Emp_ID")) <> "" Then
'            FrmEmployee.show
'            FrmEmployee.Retrive val(.TextMatrix(r, .ColIndex("Emp_ID")))
'        End If
'
'    End With
End Sub

Private Sub Grid2_DblClick()
     Static lNoteRow&, lNoteCol&, r&, c&

    With Me.GRID2
 
        r = .row
        c = .Col

        If .TextMatrix(r, .ColIndex("Emp_ID")) <> "" Then
            FrmEmployee.show
            FrmEmployee.Retrive val(.TextMatrix(r, .ColIndex("Emp_ID")))
        End If
    
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

'    My_SQL = " SELECT     *"
'    My_SQL = My_SQL & " FROM         dbo.emp_salary INNER JOIN"
'    My_SQL = My_SQL & "  dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
'    My_SQL = My_SQL & "     where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"


'My_SQL = "SELECT     *"
'My_SQL = My_SQL & "  FROM         dbo.emp_salary INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID"
 
My_SQL = "SELECT     *"
My_SQL = My_SQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.emp_salary ON dbo.TblEmployee.Emp_ID = dbo.emp_salary.emp_id LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
                      
My_SQL = My_SQL & " where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"

 

    If Dcdep.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + " and emp_salary.DepartmentID=" & val(Dcdep.BoundText)
        '  Else
        '   My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
        '  My_SQL = "SELECT * from emp_salary where sgn='" & CboYear.text & (CmbMonth.ListIndex + 1) & "'"
    End If
    
    If Me.DCEmP.BoundText <> "" Then
        '    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
        My_SQL = My_SQL + "  and emp_salary.emp_id=" & val(Me.DCEmP.BoundText)
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
  
  My_SQL = My_SQL + "  order by TblEmployee.Fullcode"
 '
 
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim StrFileName As String
StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"

    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Dim str As String
    
    If dcBranch1.text <> "" Then
    str = "«·›—⁄ : " & dcBranch1.text & Chr(13)
    End If
     
        If DCGroupID.text <> "" Then
    str = str & Chr(13) & "«·„Êﬁ⁄ : " & DCGroupID.text & Chr(13)
    End If
      
        If DCproject.text <> "" Then
    str = str & Chr(13) & "«·„‘—Ê⁄ : " & DCproject.text & Chr(13)
    End If
            
           If Dcdep.text <> "" Then
    str = str & Chr(13) & "«·ﬁ”„ : " & Dcdep.text & Chr(13)
    End If
      
     
           If dcempcontract.text <> "" Then
    str = str & Chr(13) & "‰Ê⁄ «· ⁄«ﬁœ : " & dcempcontract.text & Chr(13)
    End If
           
           If DCEmP.text <> "" Then
    str = str & Chr(13) & "«·„ÊŸ› : " & DCEmP.text & Chr(13)
    End If
    
           
    xReport.ParameterFields(6).AddCurrentValue str
    xReport.ParameterFields(47).AddCurrentValue DCGroupID.text
             If Me.DCproject.BoundText <> "" Then
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
    FrmReport.txtpath = StrFileName
    FrmReport.CRViewer.viewReport
    ' FrmReport.Show
  
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
    FrmReport.txtpath = App.path & "\reports\emp\REPORT11.rpt"
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"
End Sub

Private Sub ISButton4_Click()
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
Dim FileName As String
FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"
    'Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT10.rpt")
    Set xReport = xApp.OpenReport(FileName)
    ' App.path & "\Special\" & SystemOptions.Reportpath & "\REPORT10.rpt"
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.txtpath = FileName
    FrmReport.show
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
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
    FrmReport.txtpath = App.path & "\reports\emp\REPORT11.rpt"
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Sub

Private Sub ISButton6_Click()
    FillGridWithData

    DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

    If Dcdep.BoundText <> "" Then
        My_SQL = "SELECT * from emp_salary where payed=0 and m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "' and DepartmentID=" & Dcdep.BoundText
    Else
        My_SQL = "SELECT * from emp_salary where payed=0 and m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    End If
    
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\emp\report10Not.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.show
    FrmReport.txtpath = App.path & "\reports\emp\REPORT10not.rpt"
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
    xReport.ParameterFields(6).AddCurrentValue Dcdep.text
     
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
        .cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance")) = 0

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

Private Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
lbl(14).Caption = val(Calculate_TotalSelected2)
'Reline
End Sub

Private Sub VSFlexGrid1_Click()
lbl(14).Caption = val(Calculate_TotalSelected2)
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal row As Long, ByVal Col As Long)
Reline2
End Sub

Private Sub VSFlexGrid2_Click()
Reline2
End Sub

Private Sub VSFlexGrid3_AfterEdit(ByVal row As Long, ByVal Col As Long)
lbl(14).Caption = val(Calculate_TotalSelectedQest)
End Sub

Private Sub VSFlexGrid4_Click()
lbl(14).Caption = val(Calculate_TotalSelected16)
End Sub
