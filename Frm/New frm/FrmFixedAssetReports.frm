VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFixedAsseteports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·«’Ê· «·À«» …"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13290
   HelpContextID   =   470
   Icon            =   "FrmFixedAssetReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   13290
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   4920
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13290
      _cx             =   23442
      _cy             =   8678
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
      _GridInfo       =   $"FrmFixedAssetReports.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton Cmd 
         Height          =   480
         Left            =   30
         TabIndex        =   2
         Top             =   4410
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
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
      Begin C1SizerLibCtl.C1Elastic EleMain 
         Height          =   4365
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   13230
         _cx             =   23336
         _cy             =   7699
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
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   2
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   -1  'True
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
         Begin C1SizerLibCtl.C1Tab MainTab 
            CausesValidation=   0   'False
            Height          =   4185
            Left            =   90
            TabIndex        =   3
            Top             =   90
            Width           =   13050
            _cx             =   23019
            _cy             =   7382
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
            FrontTabForeColor=   16711680
            Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·«’Ê· «·À«» …"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   1
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   -1  'True
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic ElcContainer 
               CausesValidation=   0   'False
               Height          =   3810
               Index           =   0
               Left            =   45
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   45
               Width           =   12960
               _cx             =   22860
               _cy             =   6720
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
               Frame           =   0
               FrameStyle      =   5
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
                  Height          =   3615
                  Index           =   2
                  Left            =   90
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   12870
                  _cx             =   22701
                  _cy             =   6376
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
                  Frame           =   0
                  FrameStyle      =   5
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«” »⁄«œ« "
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   11
                     Left            =   10800
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   1800
                     Width           =   1860
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ—ﬂ… «·«’·"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   10
                     Left            =   10560
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   1440
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ› «·«’Ê· Õ Ï  «—ÌŒ „⁄Ì‰"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   9
                     Left            =   10200
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   720
                     Value           =   -1  'True
                     Width           =   2460
                  End
                  Begin VB.TextBox TxtAssesetCode 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0E0FF&
                     Height          =   315
                     Left            =   8460
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   33
                     Top             =   1440
                     Width           =   675
                  End
                  Begin VB.CheckBox Check1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·«’· „⁄Ì‰"
                     Height          =   375
                     Left            =   9240
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   1440
                     Width           =   1335
                  End
                  Begin VB.CheckBox chkEmp 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " „ÊŸ› „⁄Ì‰"
                     Height          =   375
                     Left            =   9240
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   2160
                     Width           =   1335
                  End
                  Begin VB.CheckBox chkGroup 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„Ã„Ê⁄Â „⁄Ì‰Â"
                     Height          =   375
                     Left            =   9240
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   1800
                     Width           =   1335
                  End
                  Begin VB.CheckBox ChkMain 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ «· ﬁ—Ì— ÿ»ﬁ« ··ÊÕœ… «·ﬂ»—Ï"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     ToolTipText     =   " ” Œœ„ ··„ƒ””«  «· Ì  »Ì⁄ »ÊÕœ… Ê«Õœ… ›ﬁÿ ··’„› «·Ê«Õœ"
                     Top             =   -840
                     Width           =   2535
                  End
                  Begin VB.CheckBox Chekopt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·›—⁄ „⁄Ì‰"
                     Height          =   375
                     Left            =   9240
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   1080
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ› «·«’Ê· «·À«» …"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   8
                     Left            =   10800
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   1080
                     Width           =   1860
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«‰ «Ã „Ã„⁄ Œ·«· ﬁ —… „⁄Ì‰…"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   7
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   -240
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   6
                     Left            =   14520
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   3000
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   5
                     Left            =   13920
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   3000
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   4200
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.TextBox Txt_order_no 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   3840
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÿ·»Ì«  «· Ì ·„  ”·„ Õ Ï «·«‰"
                     Height          =   195
                     Index           =   4
                     Left            =   14880
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   2760
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  ”·Ì„ «·ÿ·»Ì« "
                     Height          =   195
                     Index           =   3
                     Left            =   14880
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   3345
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«’‰«› «·„‰ Ã…  „‰ ÿ·»Ì… „⁄Ì‰"
                     Height          =   195
                     Index           =   2
                     Left            =   14640
                     RightToLeft     =   -1  'True
                     TabIndex        =   8
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· ﬂ«·Ì› «·«‰ «ÃÌ… ·√„— «‰ «Ã „⁄Ì‰"
                     Height          =   195
                     Index           =   1
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   7
                     Top             =   4680
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì—    «·«’‰«› «·„ Ã… Œ·«· › —… ÿ»ﬁ« ·”‰œ«  «” ·«„ «·«‰ «Ã «· «„"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   0
                     Left            =   -120
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   -1260
                     Width           =   5820
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1065
                     Index           =   1
                     Left            =   1200
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Top             =   2520
                     Visible         =   0   'False
                     Width           =   2355
                     _cx             =   4154
                     _cy             =   1879
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
                     Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
                     Align           =   0
                     AutoSizeChildren=   0
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
                     Style           =   1
                     TagSplit        =   2
                     PicturePos      =   4
                     CaptionStyle    =   0
                     ResizeFonts     =   0   'False
                     GridRows        =   0
                     GridCols        =   0
                     Frame           =   0
                     FrameStyle      =   5
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   ""
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin MSComCtl2.DTPicker DTPickerAccFrom 
                        BeginProperty DataFormat 
                           Type            =   1
                           Format          =   "dd/MM/yyyy"
                           HaveTrueFalseNull=   0
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   11265
                           SubFormatType   =   3
                        EndProperty
                        Height          =   345
                        Left            =   90
                        TabIndex        =   12
                        ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                        Top             =   240
                        Width           =   1500
                        _ExtentX        =   2646
                        _ExtentY        =   609
                        _Version        =   393216
                        CalendarBackColor=   -2147483624
                        CalendarTitleBackColor=   10383715
                        CheckBox        =   -1  'True
                        CustomFormat    =   "yyyy/M/d"
                        Format          =   247070723
                        CurrentDate     =   37357
                     End
                     Begin MSComCtl2.DTPicker DTPickerAccTo 
                        BeginProperty DataFormat 
                           Type            =   1
                           Format          =   "dd/MM/yyyy"
                           HaveTrueFalseNull=   0
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   11265
                           SubFormatType   =   3
                        EndProperty
                        Height          =   345
                        Left            =   90
                        TabIndex        =   13
                        ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                        Top             =   600
                        Width           =   1500
                        _ExtentX        =   2646
                        _ExtentY        =   609
                        _Version        =   393216
                        CalendarBackColor=   -2147483624
                        CalendarTitleBackColor=   10383715
                        CheckBox        =   -1  'True
                        CustomFormat    =   "yyyy/M/d"
                        Format          =   247070723
                        CurrentDate     =   37357
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„‰"
                        Height          =   285
                        Index           =   4
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   15
                        Top             =   285
                        Width           =   555
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈·Ï"
                        Height          =   285
                        Index           =   2
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   14
                        Top             =   600
                        Width           =   555
                     End
                  End
                  Begin ImpulseButton.ISButton CmdAccount 
                     Height          =   405
                     Left            =   120
                     TabIndex        =   17
                     Top             =   3120
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   714
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
                     ButtonImage     =   "FrmFixedAssetReports.frx":040B
                     ColorButton     =   14871017
                     ColorHoverText  =   16777215
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16777215
                  End
                  Begin MSDataListLib.DataCombo dcGroups 
                     CausesValidation=   0   'False
                     Height          =   315
                     Left            =   240
                     TabIndex        =   29
                     Top             =   1800
                     Width           =   8895
                     _ExtentX        =   15690
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcEmployee 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   30
                     Top             =   2160
                     Width           =   8895
                     _ExtentX        =   15690
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcFixedAssets 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   34
                     Top             =   1440
                     Width           =   8235
                     _ExtentX        =   14526
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   12640511
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboStores 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   35
                     Top             =   1080
                     Width           =   8895
                     _ExtentX        =   15690
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker1 
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "dd/MM/yyyy"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   11265
                        SubFormatType   =   3
                     EndProperty
                     Height          =   315
                     Left            =   8640
                     TabIndex        =   37
                     ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                     Top             =   720
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   556
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CheckBox        =   -1  'True
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   247070723
                     CurrentDate     =   37357
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   0
                     Left            =   4680
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   4200
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·ÿ·»Ì…"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   17
                     Left            =   4680
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   3840
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label LblAccountName 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0C8C0&
                     Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·«’Ê· «·À«» …"
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
                     Height          =   405
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   195
                     Width           =   12510
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmFixedAsseteports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String

Private Sub ChangeLang()
    'Label1.Caption = "Des"

    Me.Caption = "Fixed Assets Reports"
    Me.MainTab.TabCaption(0) = Me.Caption
    LblAccountName.Caption = Me.Caption
     OptAccount(8).Caption = Me.Caption
     OptAccount(9).Caption = "Fixed Assets Reports"
     OptAccount(10).Caption = "Trans.Fixed Assets Reports"
     Check1.Caption = "Fixed"
Chekopt.Caption = "Branch"
 chkEmp.Caption = "Employee"
 chkGroup.Caption = "Group"
    Ele(1).Caption = "In"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
 
    Cmd.Caption = "Exit"

End Sub

Private Sub Check1_Click()
    If Check1.value = vbUnchecked Then
        DcFixedAssets.text = ""
        DcFixedAssets.BoundText = 0
        DcFixedAssets.Enabled = False
          TxtAssesetCode.text = ""
        TxtAssesetCode.Enabled = False
    Else
        DcFixedAssets.Enabled = True
          TxtAssesetCode.Enabled = True
    End If
End Sub

Private Sub Chekopt_Click()

    If Chekopt.value = vbUnchecked Then
        DcboStores.text = ""
        DcboStores.BoundText = ""
        DcboStores.Enabled = False
    Else
        DcboStores.Enabled = True
    End If

End Sub

Private Sub chkEmp_Click()

    If chkEmp.value = vbUnchecked Then
        DCEmployee.text = ""
        DCEmployee.BoundText = 0
        DCEmployee.Enabled = False
    Else
        DCEmployee.Enabled = True
    End If

End Sub

Private Sub chkGroup_Click()

    If chkGroup.value = vbUnchecked Then
        dcGroups.text = ""
        dcGroups.BoundText = 0
        dcGroups.Enabled = False
    Else
        dcGroups.Enabled = True
    End If

End Sub

Private Sub Cmd_Click()
    Unload Me
End Sub

Private Sub CmdAccount_Click()
    Dim i As Integer
    Dim cAccountReport As ClsAccReports

    For i = 0 To Me.OptAccount.count - 1

        If Me.OptAccount(i).value = True Then Exit For
    Next i
 
    Select Case i

        Case 0
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
                                
            If chkGroup.value = vbChecked Then
                  
                If Me.dcGroups.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Ã„Ê⁄Â...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcGroups.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            
            End If
    
            If Chekopt.value = vbChecked Then
        
                If Me.DcboStores.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboStores.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
    
            cAccountReport.ShowProductItems ChkMain.value, val(dcGroups.BoundText), dcGroups.text, val(DcboStores.BoundText), DcboStores.text
            Set cAccountReport = Nothing
                
        Case 8
            Dim BalanceReport As New ClsOpeningBalanceReport
            'Dim X As Integer
            '      X = MsgBox(" Õ·Ì· «· ﬁ—Ì— ÿ»ﬁ« ··«›—⁄", vbYesNo + vbInformation)
            '      Set BalanceReport = New ClsOpeningBalanceReport
            '
            '     If X = vbNo Then
            '      BalanceReport.ShowFixedAssets 3, False 'Short View
            '      Else
            '      BalanceReport.ShowFixedAssets 3, True 'Short View
      
            '      End If
            If Chekopt.value = vbChecked Then
                
                If Me.DcboStores.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·›—⁄  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Branch  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboStores.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
   
            If chkGroup.value = vbChecked Then
                
                If Me.dcGroups.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„Ã„Ê⁄Â  ...!!" & CHR(13)
                    Else
                    Msg = "  Please Select Group  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcGroups.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
   
            If chkEmp.value = vbChecked Then
                
                If Me.DCEmployee.BoundText = "" Or val(DCEmployee.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„ÊŸ›  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Employee  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DCEmployee.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
         If Me.Check1.value = vbChecked Then
                If Me.DcFixedAssets.text = "" And val(Me.DcFixedAssets.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·«’·  ...!!" & CHR(13)
                    Else
                     Msg = "Please Select Fixed Assest  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcFixedAssets.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
            
            Screen.MousePointer = 11
            BalanceReport.ShowFixedAssets 3, True, val(Me.DcboStores.BoundText), val(dcGroups.BoundText), val(DCEmployee.BoundText), val(Me.DcFixedAssets.BoundText)  'Short View
       Case 9
       If IsNull(DTPicker1.value) Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ «Œ Ì«— «· «—ÌŒ"
       Else
       MsgBox "Please Select Date"
       End If
       Exit Sub
       End If
             If chkEmp.value = vbChecked Then
                
                If Me.DCEmployee.BoundText = "" Or val(DCEmployee.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„ÊŸ›  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Employee  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DCEmployee.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
            
            If Chekopt.value = vbChecked Then
                
                If Me.DcboStores.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·›—⁄  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Branch  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboStores.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
   
            If chkGroup.value = vbChecked Then
                
                If Me.dcGroups.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„Ã„Ê⁄Â  ...!!" & CHR(13)
                    Else
                    Msg = "  Please Select Group  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcGroups.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If

         If Me.Check1.value = vbChecked Then
                If Me.DcFixedAssets.text = "" And val(Me.DcFixedAssets.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·«’·  ...!!" & CHR(13)
                    Else
                     Msg = "Please Select Fixed Assest  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcFixedAssets.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
            
print_report
       Case 10

            If Chekopt.value = vbChecked Then
                
                If Me.DcboStores.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·›—⁄  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Branch  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboStores.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
   
            If chkGroup.value = vbChecked Then
                
                If Me.dcGroups.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„Ã„Ê⁄Â  ...!!" & CHR(13)
                    Else
                    Msg = "  Please Select Group  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcGroups.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If

         If Me.Check1.value = vbChecked Then
                If Me.DcFixedAssets.text = "" And val(Me.DcFixedAssets.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·«’·  ...!!" & CHR(13)
                    Else
                     Msg = "Please Select Fixed Assest  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcFixedAssets.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
            
print_report2
Case 11 '**********************************
 
            If Chekopt.value = vbChecked Then
                
                If Me.DcboStores.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·›—⁄  ...!!" & CHR(13)
                    Else
                    Msg = "Please Select Branch  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboStores.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
   
            If chkGroup.value = vbChecked Then
                
                If Me.dcGroups.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·„Ã„Ê⁄Â  ...!!" & CHR(13)
                    Else
                    Msg = "  Please Select Group  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcGroups.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If

         If Me.Check1.value = vbChecked Then
                If Me.DcFixedAssets.text = "" And val(Me.DcFixedAssets.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «·«’·  ...!!" & CHR(13)
                    Else
                     Msg = "Please Select Fixed Assest  ...!!" & CHR(13)
                    End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcFixedAssets.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    
            End If
            
print_report3
 


        Case 1

            If Not IsNumeric(Text1.text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·«„—": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowProductOrderExpenses val(Text1.text)
            Set cAccountReport = Nothing

        Case 2

            If Not IsNumeric(TXT_order_no.text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·ÿ·»Ì…": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersiTEMS TXT_order_no.text
            Set cAccountReport = Nothing

        Case 3
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersStatus TXT_order_no.text
            Set cAccountReport = Nothing

        Case 7

            If checkApility("FrmProductionReport") = False Then
                Exit Sub
            End If
        
            If SystemOptions.TypicalProduction = True Then
                Screen.MousePointer = 11
                Set cAccountReport = New ClsAccReports
        
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowProductionSummury
                Set cAccountReport = Nothing
            Else

                If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then MsgBox "Õœœ › —…", vbCritical: Exit Sub
                Screen.MousePointer = 11
                Set cAccountReport = New ClsAccReports
                CreateReportForProduction Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowProductionSummury2
                Set cAccountReport = Nothing
            End If
                
    End Select

    CuurentLogdata

End Sub
  Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.FixedAssets.NoOfInstallments, SUM(dbo.FixedAssetInstallmentsDetails.InstallmentProduct) AS EXEInstallments,FixedAssets.Quantity,FixedAssets.price, "
MySQL = MySQL & "                      SUM(dbo.FixedAssetInstallmentsDetails.InstallmentValue) AS AccDepreciation, dbo.FixedAssetInstallmentsDetails.FixedAssetID, dbo.FixedAssets.Fullcode,"
MySQL = MySQL & "                      dbo.FixedAssets.Name, FixedAssets.namee,dbo.FixedAssets.PurchasePrice, dbo.FixedAssetsGroup.GroupName, dbo.FixedAssetsGroup.GroupNamee,"
MySQL = MySQL & "                      dbo.FixedAssets.InstallmentValue AS currentInstallemntValue, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.KhordaPrice, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.FixedAssetsGroup.GroupID, dbo.TblBranchesData.branch_id,"
MySQL = MySQL & "                      dbo.GetInstalAddValueByDate(dbo.FixedAssetInstallmentsDetails.FixedAssetID, " & SQLDate(DTPicker1.value, True) & ") AS adValue, dbo.FixedAssets.Emp_id, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee,dbo.GetInstalDiscuValueByDate(dbo.FixedAssetInstallmentsDetails.FixedAssetID, " & SQLDate(DTPicker1.value, True) & ") AS DiscuntValue"
MySQL = MySQL & " FROM         dbo.FixedAssetInstallmentsDetails INNER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.FixedAssetInstallmentsDetails.FixedAssetID = dbo.FixedAssets.id INNER JOIN"
MySQL = MySQL & "                      dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID INNER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "  WHERE 1=1"
If DCEmployee.text <> "" And val(DCEmployee.BoundText) <> 0 Then
MySQL = MySQL & "  and     (dbo.FixedAssetsGroup.Emp_id = " & val(DCEmployee.BoundText) & ")"
End If

If DcboStores.text <> "" And val(DcboStores.BoundText) <> 0 Then
MySQL = MySQL & "  and     (dbo.TblBranchesData.branch_id = " & val(DcboStores.BoundText) & ")"
End If
If DcFixedAssets.text <> "" And val(DcFixedAssets.BoundText) <> 0 Then
MySQL = MySQL & "  AND (dbo.FixedAssetInstallmentsDetails.FixedAssetID = " & val(DcFixedAssets.BoundText) & ") "
End If
If dcGroups.text <> "" And val(dcGroups.BoundText) <> 0 Then
MySQL = MySQL & " AND (dbo.FixedAssetsGroup.GroupID = " & val(dcGroups.BoundText) & ") "
End If

If Not IsNull(DTPicker1.value) Then
MySQL = MySQL & " and  (dbo.FixedAssetInstallmentsDetails.InstallmentDate <= " & SQLDate(DTPicker1.value, True) & ")"
End If
MySQL = MySQL & " GROUP BY dbo.FixedAssetInstallmentsDetails.FixedAssetID, dbo.FixedAssets.Fullcode, dbo.FixedAssets.Name,FixedAssets.namee, dbo.FixedAssets.NoOfInstallments, "
MySQL = MySQL & "                      dbo.FixedAssets.PurchasePrice, dbo.FixedAssetsGroup.GroupName, dbo.FixedAssetsGroup.GroupNamee, dbo.FixedAssets.InstallmentValue,"
MySQL = MySQL & "                      dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.KhordaPrice, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                      dbo.FixedAssetsGroup.GroupID, dbo.TblBranchesData.branch_id, dbo.FixedAssets.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee,FixedAssets.Quantity,FixedAssets.price"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportAudit12.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportAudit12E.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÌÊÃœ »Ì«‰«  "
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

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

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
   Function print_report3(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.notes_all.FAID, dbo.FixedAssets.Fullcode, dbo.FixedAssets.Name, dbo.FixedAssets.group_id, dbo.FixedAssets.namee, dbo.notes_all.NoteSerial ,FixedAssets.Quantity,FixedAssets.price, "
MySQL = MySQL & "                         dbo.notes_all.NoteSerial1, dbo.FixedAssets.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.notes_all.NoteDate,"
MySQL = MySQL & "                         dbo.notes_all.NoteType, TblBranchesData_1.branch_name AS tbranch_name, TblBranchesData_1.branch_namee AS tbranch_namee,"
MySQL = MySQL & "                         dbo.FixedAssetsGroup.GroupName, dbo.FixedAssetsGroup.GroupNamee, dbo.notes_all.UserID, dbo.TblUsers.UserName, dbo.notes_all.PurchasePrice,"
MySQL = MySQL & "                         dbo.notes_all.LoseProfitValue , dbo.notes_all.AccDepre, dbo.notes_all.currentvalue"
MySQL = MySQL & "                  FROM         dbo.TblUsers INNER JOIN"
MySQL = MySQL & "                                        dbo.notes_all ON dbo.TblUsers.UserID = dbo.notes_all.UserID INNER JOIN"
MySQL = MySQL & "                                        dbo.FixedAssets INNER JOIN"
MySQL = MySQL & "                                        dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID ON dbo.notes_all.FAID = dbo.FixedAssets.id INNER JOIN"
MySQL = MySQL & "                                        dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id INNER JOIN"
MySQL = MySQL & "                                        dbo.TblBranchesData TblBranchesData_1 ON dbo.notes_all.branch_no = TblBranchesData_1.branch_id"
 MySQL = MySQL & "    WHERE 1=1"

If DcboStores.text <> "" And val(DcboStores.BoundText) <> 0 Then
MySQL = MySQL & "  and     (dbo.notes_all.branch_no= " & val(DcboStores.BoundText) & ")"
End If
If DcFixedAssets.text <> "" And val(DcFixedAssets.BoundText) <> 0 Then
MySQL = MySQL & "  AND (dbo.notes_all.FAID= " & val(DcFixedAssets.BoundText) & ") "
End If
If dcGroups.text <> "" And val(dcGroups.BoundText) <> 0 Then
MySQL = MySQL & " AND (dbo.FixedAssets.group_id = " & val(dcGroups.BoundText) & ") "
End If

If Not IsNull(DTPickerAccFrom.value) Then
MySQL = MySQL & " and  (dbo.notes_all.NoteDate >= " & SQLDate(DTPickerAccFrom.value, True) & ")"
End If
If Not IsNull(DTPickerAccTo.value) Then
MySQL = MySQL & " and  (dbo.notes_all.NoteDate <= " & SQLDate(DTPickerAccTo.value, True) & ")"
End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportTransection1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportTransectionE1.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault0
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÌÊÃœ »Ì«‰« "
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

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

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL
    

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
  Function print_report2(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.FixedAssetInstallmentsDetails.CurrentValue, dbo.FixedAssetInstallmentsDetails.InstallmentValue, dbo.FixedAssetInstallmentsDetails.InstallmentDate ,FixedAssets.Quantity,FixedAssets.price, "
MySQL = MySQL & "                      dbo.FixedAssetInstallmentsDetails.AccDepreciation, dbo.FixedAssetInstallmentsDetails.RemainInstallments,"
MySQL = MySQL & "                      dbo.FixedAssetInstallmentsDetails.FixedAssetInstallmentsid, dbo.FixedAssetInstallmentsDetails.InstallmentProduct, dbo.FixedAssetInstallmentsDetails.InstallmentID,"
MySQL = MySQL & "                       dbo.FixedAssetInstallmentsDetails.FixedAssetID, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.FixedAssets.Fullcode, dbo.FixedAssets.group_id,"
MySQL = MySQL & "                      dbo.FixedAssetsGroup.GroupName, dbo.FixedAssetsGroup.Fullcode AS GroupFullcode, dbo.FixedAssetsGroup.GroupNamee, dbo.FixedAssetInstallments.BranchId,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
MySQL = MySQL & " FROM         dbo.FixedAssetInstallments LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.FixedAssetInstallments.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssetInstallmentsDetails ON"
MySQL = MySQL & "                      dbo.FixedAssetInstallments.FixedAssetInstallmentsid = dbo.FixedAssetInstallmentsDetails.FixedAssetInstallmentsid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssetsGroup RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.FixedAssetsGroup.GroupID = dbo.FixedAssets.group_id ON dbo.FixedAssetInstallmentsDetails.FixedAssetID = dbo.FixedAssets.id"
MySQL = MySQL & "  WHERE 1=1"

If DcboStores.text <> "" And val(DcboStores.BoundText) <> 0 Then
MySQL = MySQL & "  and     (dbo.FixedAssetInstallments.BranchId= " & val(DcboStores.BoundText) & ")"
End If
If DcFixedAssets.text <> "" And val(DcFixedAssets.BoundText) <> 0 Then
MySQL = MySQL & "  AND (dbo.FixedAssetInstallmentsDetails.FixedAssetID= " & val(DcFixedAssets.BoundText) & ") "
End If
If dcGroups.text <> "" And val(dcGroups.BoundText) <> 0 Then
MySQL = MySQL & " AND (dbo.FixedAssetsGroup.GroupID = " & val(dcGroups.BoundText) & ") "
End If

If Not IsNull(DTPickerAccFrom.value) Then
MySQL = MySQL & " and  (dbo.FixedAssetInstallmentsDetails.InstallmentDate >= " & SQLDate(DTPickerAccFrom.value, True) & ")"
End If
If Not IsNull(DTPickerAccTo.value) Then
MySQL = MySQL & " and  (dbo.FixedAssetInstallmentsDetails.InstallmentDate <= " & SQLDate(DTPickerAccTo.value, True) & ")"
End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportTransection.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "FixedASSETReportTransectionE.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = " ·«ÌÊÃœ »Ì«‰« "
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

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

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub DcFixedAssets_Change()
   Dim AsseCode1 As String
If val(DcFixedAssets.BoundText) <> 0 Then
GetAsseteCode_ID val(DcFixedAssets.BoundText), AsseCode1, 0
TxtAssesetCode.text = AsseCode1
End If
End Sub
Sub GetAsseteCode_ID(Optional ByRef ID As Double = 0, Optional ByRef fullcode As String = "", Optional Typ As Integer = 0)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If Typ = 0 Then
sql = "select Fullcode  from FixedAssets where id=" & ID & " "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
fullcode = IIf(IsNull(Rs7("Fullcode").value), "", Rs7("Fullcode").value)
Else
fullcode = ""
End If
Else
sql = "select ID  from FixedAssets where Fullcode='" & fullcode & "' "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ID = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
Else
ID = 0
End If
End If
End Sub
Private Sub DcFixedAssets_Click(Area As Integer)
DcFixedAssets_Change
End Sub

Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyF3 Then
        FixedAssetsSearch.RetrunType = 13
        FixedAssetsSearch.show vbModal
  
    End If


End Sub

Private Sub Form_Load()
    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
 
    ScreenNameArabic = "  ‹‹ﬁ‹‹«—Ì‹‹‹— «·«‰‰«Ã  "
    ScreenNameEnglish = "  Production Report "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches DcboStores
 Dcombos.GetFixedAssets Me.DcFixedAssets
    Dcombos.GetFixedAssetsGroup Me.dcGroups
    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo
    SetDtpickerDate Me.DTPicker1
    
 If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "  select Emp_ID,Emp_name  from TblEmployee order by Emp_name   "
Else
StrSQL = "  select Emp_ID,Emp_nameE  from TblEmployee order by Emp_name   "
End If
    fill_combo DCEmployee, StrSQL

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
    Dim i As Integer
  
    LogTextA = "    ‘«‘… " & ScreenNameArabic & "   ⁄—÷  ﬁ—Ì— "

    For i = 0 To 7

        If OptAccount(i).value = True Then
            LogTextA = LogTextA & OptAccount(i).Caption
        End If
 
    Next i
 
    LogTextA = LogTextA & "    «·› —… „‰  " & DTPickerAccFrom.value & "   «·Ï  " & DTPickerAccTo.value
  
    LogTexte = "    Screen " & ScreenNameEnglish & "   View Report   "

    For i = 0 To 7

        If OptAccount(i).value = True Then
            LogTexte = LogTextA & OptAccount(i).Caption
        End If
 
    Next i
 
    LogTexte = LogTexte & "    From " & DTPickerAccFrom.value & "   To  " & DTPickerAccTo.value
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "V"
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "V"
    End If
    
End Function

Private Sub OptAccount_Click(Index As Integer)
DTPicker1.Visible = False
Ele(1).Visible = False
chkEmp.Enabled = True
If OptAccount(9).value = True Then
DTPicker1.Visible = True
ElseIf OptAccount(10).value = True Or OptAccount(11).value = True Then
Ele(1).Visible = True
chkEmp.Enabled = False
chkEmp.value = vbUnchecked
chkEmp_Click
End If
End Sub

Private Sub TxtAssesetCode_KeyPress(KeyAscii As Integer)
Dim AsseID As Double
If TxtAssesetCode.text <> "" Then
GetAsseteCode_ID AsseID, TxtAssesetCode.text, 1
DcFixedAssets.BoundText = AsseID
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub


