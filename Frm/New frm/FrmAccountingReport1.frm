VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAccountingReport1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·‹‹‹Õ‹‹”‹‹«»‹« "
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   HelpContextID   =   470
   Icon            =   "FrmAccountingReport1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   12510
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
      Height          =   7755
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12510
      _cx             =   22066
      _cy             =   13679
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
      AutoSizeChildren=   0
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
      Begin VB.TextBox txt_mod_flag 
         Alignment       =   1  'Right Justify
         Height          =   480
         Left            =   5970
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   7245
         Visible         =   0   'False
         Width           =   3240
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   480
         Left            =   30
         TabIndex        =   1
         Top             =   7245
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
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   480
         Left            =   1650
         TabIndex        =   2
         Top             =   7245
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   847
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
      End
      Begin C1SizerLibCtl.C1Elastic EleMain 
         Height          =   6240
         Left            =   150
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   150
         Width           =   12450
         _cx             =   21960
         _cy             =   11007
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
         AutoSizeChildren=   0
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   4575
            Begin ImpulseAniLabel.ISAniLabel TxtDept 
               Height          =   315
               Left            =   2280
               TabIndex        =   50
               Top             =   600
               Width           =   2160
               _ExtentX        =   3810
               _ExtentY        =   556
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "FrmAccountingReport1.frx":038A
               BackColor       =   14871017
               Alignment       =   2
               Caption         =   "0"
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin ImpulseAniLabel.ISAniLabel TxtCredit 
               Height          =   315
               Left            =   120
               TabIndex        =   51
               Top             =   600
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   556
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "FrmAccountingReport1.frx":04EC
               BackColor       =   14871017
               Alignment       =   2
               Caption         =   "0"
               ColorHover      =   16711680
               RightToLeft     =   -1  'True
               ImageCount      =   0
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   1335
            End
         End
         Begin C1SizerLibCtl.C1Tab MainTab 
            CausesValidation=   0   'False
            Height          =   540
            Left            =   90
            TabIndex        =   6
            Top             =   -750
            Width           =   6060
            _cx             =   10689
            _cy             =   952
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
            Caption         =   "«·ﬁÊ«∆„ «·„«·Ì…"
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
               Height          =   165
               Index           =   0
               Left            =   45
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   45
               Width           =   5970
               _cx             =   10530
               _cy             =   291
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
               AutoSizeChildren=   2
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
                  Height          =   795
                  Index           =   3
                  Left            =   5160
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   6000
                  Visible         =   0   'False
                  Width           =   2970
                  _cx             =   5239
                  _cy             =   1402
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
                  BackColor       =   14653050
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
                  Begin VB.TextBox TxtEhlak 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0FFFF&
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "2222/22/22"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   11265
                        SubFormatType   =   0
                     EndProperty
                     Height          =   345
                     Left            =   435
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   330
                     Width           =   2085
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00DF967A&
                     Caption         =   "„‰ ›÷·ﬂ √œŒ· ﬁÌ„… ≈Â·«ﬂ«  «·› —…"
                     Height          =   240
                     Index           =   3
                     Left            =   165
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   450
                     Width           =   2745
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   30
                  Index           =   2
                  Left            =   90
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   5790
                  _cx             =   10213
                  _cy             =   53
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
                     Caption         =   "ÿ»«⁄… «·œ·Ì· «·„Õ«”»Ï"
                     Height          =   240
                     Index           =   6
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   1020
                     Width           =   3045
                  End
                  Begin VB.TextBox TxtSearch 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     Height          =   345
                     Left            =   1410
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     ToolTipText     =   "√ﬂ » ﬂÊœ «·Õ”«» «·„—«œ «·»ÕÀ ⁄‰Â"
                     Top             =   90
                     Visible         =   0   'False
                     Width           =   1650
                  End
                  Begin VB.CommandButton CmdSeach 
                     BackColor       =   &H00C0C8C0&
                     Caption         =   "»ÕÀ"
                     Height          =   345
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     Style           =   1  'Graphical
                     TabIndex        =   29
                     ToolTipText     =   "»œ¡ ⁄„·Ì… «·»ÕÀ"
                     Top             =   90
                     Visible         =   0   'False
                     Width           =   405
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹…"
                     Height          =   195
                     HelpContextID   =   520
                     Index           =   5
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   3180
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Ì‹“«‰‹Ì‹…  "
                     Height          =   150
                     Index           =   4
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   3510
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ«∆„… «·œŒ· "
                     Height          =   240
                     Index           =   3
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   2910
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— „ «Ã—…"
                     Height          =   30
                     Index           =   2
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   2685
                     Visible         =   0   'False
                     Width           =   1260
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«»"
                     Height          =   210
                     HelpContextID   =   480
                     Index           =   0
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   2415
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "√” «– ⁄«„ ·‹‹ ...."
                     Height          =   195
                     Index           =   1
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   2655
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»«⁄… «·ﬁÌÊœ «·ÌÊ„Ì…"
                     Height          =   180
                     Index           =   7
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   1320
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»«⁄… «·ﬁÌÊœ «·ÌÊ„Ì…  Õ·Ì·Ì"
                     Height          =   195
                     Index           =   8
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   1560
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ·„—ﬂ“ «· ﬂ·›…"
                     Height          =   195
                     Index           =   9
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   1905
                     Width           =   3045
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ·„‘—Ê⁄ „⁄Ì‰"
                     Height          =   195
                     Index           =   10
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   2160
                     Width           =   3045
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   660
                     Left            =   90
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   18
                     Top             =   3840
                     Width           =   4770
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ÕœÌœ „‘—Ê⁄"
                     Height          =   1095
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   4560
                     Visible         =   0   'False
                     Width           =   3270
                     Begin MSDataListLib.DataCombo dcprojects 
                        Bindings        =   "FrmAccountingReport1.frx":064E
                        Height          =   315
                        Left            =   0
                        TabIndex        =   15
                        Top             =   240
                        Width           =   3015
                        _ExtentX        =   5318
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
                     Begin MSDataListLib.DataCombo Dcdetails 
                        Bindings        =   "FrmAccountingReport1.frx":0663
                        Height          =   315
                        Left            =   0
                        TabIndex        =   16
                        Top             =   600
                        Width           =   2055
                        _ExtentX        =   3625
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   16777215
                        ListField       =   ""
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
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "»‰œ/⁄„·Ì…"
                        Height          =   255
                        Left            =   2040
                        RightToLeft     =   -1  'True
                        TabIndex        =   17
                        Top             =   600
                        Width           =   1095
                     End
                  End
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ÕœÌœ „—ﬂ“ «· ﬂ·›…"
                     Height          =   1095
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   4560
                     Visible         =   0   'False
                     Width           =   3270
                     Begin MSDataListLib.DataCombo DcCostCenter 
                        Bindings        =   "FrmAccountingReport1.frx":0678
                        Height          =   315
                        Left            =   0
                        TabIndex        =   13
                        Top             =   240
                        Width           =   3015
                        _ExtentX        =   5318
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
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1095
                     Index           =   1
                     Left            =   0
                     TabIndex        =   32
                     TabStop         =   0   'False
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2205
                     _cx             =   3889
                     _cy             =   1931
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
                        TabIndex        =   33
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
                        Format          =   100073475
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
                        TabIndex        =   34
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
                        Format          =   100073475
                        CurrentDate     =   37357
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈·Ï"
                        Height          =   285
                        Index           =   2
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   36
                        Top             =   600
                        Width           =   555
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„‰"
                        Height          =   285
                        Index           =   4
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   35
                        Top             =   285
                        Width           =   555
                     End
                  End
                  Begin ImpulseButton.ISButton CmdAccount 
                     Height          =   450
                     Left            =   90
                     TabIndex        =   37
                     Top             =   5715
                     Width           =   1260
                     _ExtentX        =   2223
                     _ExtentY        =   794
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
                     ButtonImage     =   "FrmAccountingReport1.frx":068D
                     ColorButton     =   14871017
                     ColorHoverText  =   16777215
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16777215
                  End
                  Begin VB.Label LblAccountName 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0C8C0&
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
                     Height          =   465
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   495
                     Width           =   5535
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·‘—Õ"
                     Height          =   255
                     Left            =   4800
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   3840
                     Width           =   570
                  End
               End
               Begin MSComCtl2.DTPicker DtpSheet 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   40
                  Top             =   6480
                  Visible         =   0   'False
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   556
                  _Version        =   393216
                  CalendarBackColor=   -2147483624
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   100073475
                  CurrentDate     =   37958
               End
               Begin MSComctlLib.ImageList ImgLstChartTree 
                  Left            =   4500
                  Top             =   1350
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  ImageWidth      =   16
                  ImageHeight     =   16
                  MaskColor       =   12632256
                  _Version        =   393216
                  BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                     NumListImages   =   5
                     BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountingReport1.frx":0A27
                        Key             =   "Expanded_Node"
                     EndProperty
                     BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountingReport1.frx":1879
                        Key             =   "Root"
                     EndProperty
                     BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountingReport1.frx":1C13
                        Key             =   "Open_Node"
                     EndProperty
                     BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountingReport1.frx":1FAD
                        Key             =   "Closed_Node"
                     EndProperty
                     BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountingReport1.frx":2347
                        Key             =   "Item"
                     EndProperty
                  EndProperty
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "›Ï"
                  Height          =   285
                  Index           =   5
                  Left            =   5265
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   6000
                  Visible         =   0   'False
                  Width           =   255
               End
            End
         End
         Begin MSComctlLib.TreeView TrvAccounts 
            Height          =   5700
            Left            =   4890
            TabIndex        =   42
            Top             =   90
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   10054
            _Version        =   393217
            Indentation     =   18
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            SingleSel       =   -1  'True
            ImageList       =   "ImgLstChartTree"
            Appearance      =   1
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   4
            Left            =   120
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   4485
            _cx             =   7911
            _cy             =   1296
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
            Begin MSComCtl2.DTPicker Enddate 
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
               TabIndex        =   45
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   100073475
               CurrentDate     =   47484
            End
            Begin MSComCtl2.DTPicker Bdate 
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
               Left            =   2250
               TabIndex        =   44
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
               Format          =   100073475
               CurrentDate     =   36526
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   285
               Index           =   1
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   0
               Left            =   3630
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   285
               Width           =   555
            End
         End
      End
   End
End
Attribute VB_Name = "FrmAccountingReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String
'Option Explicit
'Dim RPTCompany_Name_Arabic  As String
'Dim RPTComment_Arabic       As String
'Dim RPTCompany_Name_Eng     As String
'Dim RPTComment_Eng          As String
'Dim RPTCurrency
'
'Private Sub Cmd_Click()
'Unload Me
'End Sub
'
'Private Sub CmdSeach_Click()
''Me.LblAccountName.Caption = StartSearch(Me.TreeView2, Me.TxtSearch.text, True)
'End Sub
'
'Private Sub Form_Load()
'Dim RsOpt                   As New ADODB.Recordset
''Disable the Redram of the Tree Control to fast load
''Call SendMessage(Me.TreeView2.hwnd, WM_SETREDRAW, 0, 0)
'Set Me.TreeView2.ImageList = FrmSystemTrees.TreeView2.ImageList
''Load the Tree Accounting
'LoadTreeAccount Me.TreeView2
'If SystemOptions.UserInterface = EnglishInterface Then
'    SetInterface Me
'    ChangeLang
'End If
''Enaable the Redraw of the control
''Call SendMessage(Me.TreeView2.hwnd, WM_SETREDRAW, -1, 0)
'
'Call open_rs("select OPTIONS.Company_Name_Arabic, OPTIONS.Comment_Arabic, OPTIONS.Company_Name_Eng, OPTIONS.currency_unite, OPTIONS.Comment_Eng From OPTIONS", RsOpt)
'RPTCompany_Name_Arabic = IIf(IsNull(RsOpt!Company_Name_Arabic), "", RsOpt!Company_Name_Arabic)   'rs!Company_Name_Arabic
'RPTComment_Arabic = IIf(IsNull(RsOpt!Comment_Arabic), "", RsOpt!Comment_Arabic)    'rs!Comment_Arabic
'RPTCompany_Name_Eng = IIf(IsNull(RsOpt!Company_Name_Eng), "", RsOpt!Company_Name_Eng)   'rs!Company_Name_Eng
'RPTComment_Eng = IIf(IsNull(RsOpt!Comment_Eng), "", RsOpt!Comment_Eng)   'rs!Comment_Eng
'RPTCurrency = IIf(IsNull(RsOpt!currency_unite), "", RsOpt!currency_unite)
'RsOpt.Close
'Set RsOpt = Nothing
''==========================initial Setting For Controls
'Me.DtpSheet.Value = Date
'Me.DTPickerAccFrom.Value = Date
'Me.DTPickerAccTo.Value = Date
''Hide this Tab at this monent
'Me.MainTab.TabVisible(1) = False
'Me.left = (MDIFrmamin.ScaleWidth - Me.ScaleWidth) / 2
'Me.top = (MDIFrmamin.ScaleHeight - Me.ScaleHeight) / 2
'
'End Sub
'
'
'
'Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error Resume Next
'Me.LblAccountName.Caption = Me.TreeView2.SelectedItem.text
'End Sub
'
'Private Sub TxtEhlak_KeyPress(KeyAscii As Integer)
'If KeyAscii = 8 Then Exit Sub
'If CBool(InStr(1, ".", Chr(KeyAscii))) And CBool(InStr(1, Me.TxtEhlak, Chr(KeyAscii))) Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub TreeView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'If InStr(Me.TreeView2.SelectedItem.Tag, "last") Then
'    If Me.OptAccount(0).Value = True Then Me.CmdAccount.Enabled = True
'    If Me.OptAccount(1).Value = True Then Me.CmdAccount.Enabled = False
'    If Button = 2 Then
'        MDIFrmamin.SubmasterMnu(0).Enabled = True
'        MDIFrmamin.SubmasterMnu(1).Enabled = True
'        MDIFrmamin.SubmasterMnu(2).Enabled = False
'        MDIFrmamin.PopupMenu MDIFrmamin.reportMnu
'    End If
'Else
'    If Me.OptAccount(1).Value = True Then Me.CmdAccount.Enabled = True
'    If Me.OptAccount(0).Value = True Then Me.CmdAccount.Enabled = False
'    If Button = 2 Then   'And Me.OptAccount(1).Value = True
'        MDIFrmamin.SubmasterMnu(0).Enabled = False
'        MDIFrmamin.SubmasterMnu(1).Enabled = False
'        MDIFrmamin.SubmasterMnu(2).Enabled = True
'        MDIFrmamin.PopupMenu MDIFrmamin.reportMnu
'    End If
'End If
'End Sub
'Private Sub OptAccount_Click(Index As Integer)
'
'Select Case Index
'    Case 0
'
'        Me.eLE(2).Visible = True
'        Me.TxtEhlak.text = ""
'        Me.eLE(3).Visible = False
'    Case 1
'
'        Me.eLE(2).Visible = False
'        Me.TxtEhlak.text = ""
'        Me.eLE(3).Visible = False
'    Case 2
'
'        Me.eLE(2).Visible = True
'        Me.TxtEhlak.text = ""
'        Me.eLE(3).Visible = False
'    Case 3
'        'Me.CmdAccount.Enabled = True
'        Me.eLE(2).Visible = True
'        Me.TxtEhlak.text = ""
'        Me.eLE(3).Visible = False
'    Case 4, 5
'        'Me.CmdAccount.Enabled = True
'        Me.eLE(2).Visible = False
'        Me.TxtEhlak.text = ""
'        Me.eLE(3).Visible = False
'    Case 6
'        'Me.CmdAccount.Enabled = True
'        Me.eLE(2).Visible = False
'        Me.eLE(3).Visible = False
'End Select
'If OptAccount(4).Value Or OptAccount(5).Value Then
'    lbl(0).Visible = True
'    DtpSheet.Visible = True
'Else
'    lbl(0).Visible = False
'    DtpSheet.Visible = False
'End If
'End Sub
'
'Public Sub CmdAccount_Click()
''By Nour  25/5/2003
'Dim MySQL As String
'Dim RS1                     As New ADODB.Recordset
'Dim Rs2                     As New ADODB.Recordset  '«·Œ«’ »»Ì«‰«  «·„ «Ã—…
'Dim DEP_VALUE               As Double
'Dim CRED_VALUE              As Double
'Dim open_balance            As Double   'the value of openning balance OR specephic period
'Dim counter_opt As Integer
'Dim HHH As Double, openning_From As Double, purchase_From As Double
'Dim salles_to As Double, purchaseback_to As Double
'Dim sallesback_From As Double, ending_to As Double
'Dim Zoom_Report As Integer
'
''---------------
'Dim RsData As New ADODB.Recordset
'Dim xApp As New CRAXDRT.Application
'Dim xReport As CRAXDRT.Report
'Dim Frm As FrmPrint
'Dim cAccountReport As ClsAccReports
'Dim Msg As String
'On Error GoTo ErrTrap
''----------------------------------
''Dim HHH As Integer
''Dim openning_From As Integer
''If Me.TxtAccFrom.Visible = True Or Me.TxtAccTo.Visible = True Then MsgBox "ÌÃ» ≈Œ Ì«— «· «—ÌŒ „‰ ... Ê≈·Ï ... ", vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, "„œÌ— «· ﬁ«—Ì—  ": Exit Sub
'If Me.DTPickerAccFrom.Value > Me.DTPickerAccTo.Value Then
'    MsgBox "Œÿ√ ›Ì «· «—ÌŒ...." & Chr(13) & " «—ÌŒ »œ«Ì… «·› —… ·«»œ «‰ Ìﬁ· ⁄‰  «—ÌŒ ‰Â«Ì… «·› —…....", vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, "„œÌ— «· ﬁ«—Ì—"
'    Screen.MousePointer = 0
'    Exit Sub
'End If
'
'Screen.MousePointer = 11
'For counter_opt = 0 To Me.OptAccount.count - 1
'    If Me.OptAccount(counter_opt).Value = True Then Exit For
'Next counter_opt
'
'Select Case counter_opt
'    Case 6
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.ShowChartAccounts
'        Set cAccountReport = Nothing
'    Case 0
'        'Õ”«» «” «– „”«⁄œ
'        If Me.TreeView2.SelectedItem Is Nothing Then
'            Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & _
'            "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.BegineDate = Me.DTPickerAccFrom.Value
'        cAccountReport.EndDate = Me.DTPickerAccTo.Value
'        cAccountReport.ShowLedger Me.TreeView2.SelectedItem.Key, _
'        Me.TreeView2.SelectedItem.text
'        Set cAccountReport = Nothing
'    Case 1
'        ' Õ”«» «” «– ⁄«„
'        If Me.TreeView2.SelectedItem Is Nothing Then
'            Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & _
'            "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï "
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.ShowMaterLedgar _
'            Me.TreeView2.SelectedItem.Key, Me.TreeView2.SelectedItem.text
'        Set cAccountReport = Nothing
'    Case 2  ' ﬁ‹‹—Ì‹‹— «·„ ‹‹«Ã—…
'        '—’Ìœ √Ê· «·„œ…
'        openning_From = 0
'        '«·„‘ —Ì« 
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a2' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            purchase_From = Rs2!SumValue
'        Else
'            purchase_From = 0
'        End If
'        Rs2.Close
'
'        '„—œÊœ«   «·„»Ì⁄« 
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a4' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            sallesback_From = Rs2!SumValue
'        Else
'            sallesback_From = 0
'            End If
'        Rs2.Close
'
'        '«·„»Ì⁄« 
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a1' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            salles_to = Rs2!SumValue
'        Else
'            salles_to = 0
'        End If
'        Rs2.Close
'
'        '„—œÊœ«  «·„‘ —Ì« 
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a3' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            purchaseback_to = Rs2!SumValue
'        Else
'            purchaseback_to = 0
'        End If
'        Rs2.Close
'
'        '—’Ìœ ¬Œ— «·„œ…
'        ending_to = 270000
'        Me.rdc.Refresh
'        'If Me.rdc.Resultset.RowCount = 0 Then
'        '    Screen.MousePointer = 0
'        '    MsgBox " ·«  ÊÃœ √Ï »Ì«‰«  „ÿ«»ﬁ… ·Â–« «·«Œ Ì«—" & vbCrLf & "√Ê ·«Œ Ì«—  «—ÌŒ «· ﬁ—Ì— „‰00 ≈·Ï00      ", vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, " ‰»ÌÂ .."
'        'Else
'            CR.ReportFileName = App.Path & "\Reports\" & "Motagra.rpt"
'            CR.ParameterFields(3) = "report_header;" & "  ﬁ—Ì— »«·„ «Ã—… ›Ì «·› —…" & "(" & headerdate(Me.DTPickerAccFrom) & ")" & " ≈·Ï «·› —… (" & headerdate(Me.DTPickerAccTo) & ")∫" & ";1"
'            CR.ReportTitle = RPTCompany_Name_Arabic
'            CR.ParameterFields(1) = "comment_arabic;" & RPTComment_Arabic & ";1"
'            CR.ParameterFields(0) = "name_english;" & RPTCompany_Name_Eng & ";1"
'            CR.ParameterFields(2) = "comment_english;" & RPTComment_Eng & ";1"
'
'            CR.ParameterFields(4) = "openning;" & openning_From & ";1"
'            CR.ParameterFields(5) = "ending;" & ending_to & ";1"
'            CR.ParameterFields(6) = "purchase;" & purchase_From & ";1"
'            CR.ParameterFields(7) = "sell_back;" & sallesback_From & ";1"
'            CR.ParameterFields(8) = "sells;" & salles_to & ";1"
'            CR.ParameterFields(9) = "purchase_back;" & purchaseback_to & ";1"
'            CR.WindowShowPrintSetupBtn = True
'            CR.WindowShowSearchBtn = True
'            CR.WindowTitle = RPTCompany_Name_Eng
'            CR.WindowState = crptMaximized
'            CR.Action = 1
'            CR.PageZoom (Zoom_Report)
'            Screen.MousePointer = 0
'            CR.Reset
'     Case 3
'        Dim Mogmal_ As String
'        Dim generals_ As String
'        Dim ehlak_ As String
'        Dim discount_From_ As String
'        Dim discount_to_ As String
'        Dim other_income_ As String
'
'        If Me.TxtEhlak.text = "" Then
'            Screen.MousePointer = 0
'            Me.eLE(3).Visible = True
'            TxtEhlak.SetFocus
'            Exit Sub
'        Else
'            Screen.MousePointer = 11
'                        '*************Õ”«» „Ã„· «·—»Õ √Ê «·Œ”«—… („ «Ã—…) 7
'            '—’Ìœ √Ê· «·„œ… ********************
'            openning_From = 0
'            '«·„‘ —Ì« ***********************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a2' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                purchase_From = Rs2!SumValue
'            Else
'                purchase_From = 0
'            End If
'            Rs2.Close
'            '„—œÊœ«   «·„»Ì⁄«  *********************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a4' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                sallesback_From = Rs2!SumValue
'            Else
'                sallesback_From = 0
'                End If
'            Rs2.Close
'            '«·„»Ì⁄«  ***********************8
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a1' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                salles_to = Rs2!SumValue
'            Else
'                salles_to = 0
'            End If
'            Rs2.Close
'            '„—œÊœ«  «·„‘ —Ì«  **************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a3' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                purchaseback_to = Rs2!SumValue
'            Else
'                purchaseback_to = 0
'            End If
'            Rs2.Close
'            '—’Ìœ ¬Œ— «·„œ…' ************
'            ending_to = 270000
'            '„Ã„· —»Õ ÊŒ”«—…
'            Mogmal_ = Val(salles_to) + Val(purchaseback_to) + Val(ending_to) - Val(openning_From) - Val(purchase_From) - Val(sallesback_From)
'
'
'            ''*****************Õ”«» „’—Ê›«  ⁄„Ê„Ì…
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a1' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                generals_ = Rs2!SumValue
'            Else
'                generals_ = 0
'            End If
'            Rs2.Close
'            ''*****************Õ”«» Œ’„ „”„ÊÕ »Â
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a5' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                discount_From_ = Rs2!SumValue
'            Else
'                discount_From_ = 0
'            End If
'            Rs2.Close
'            ''*****************Õ”«»  ≈Ì—«œ«  √Œ—Ï
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a2' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                other_income_ = Rs2!SumValue
'            Else
'                other_income_ = 0
'            End If
'            Rs2.Close
'            ''*****************Õ”«» «·Œ’„ «·„ﬂ ”»
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a4' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                discount_to_ = Rs2!SumValue
'            Else
'                discount_to_ = 0
'            End If
'            Rs2.Close
'            ''********************Õ”«» «·Â«·ﬂ
'            ehlak_ = Val(Me.TxtEhlak)
'
'
'            CR.ReportFileName = App.Path & "\Reports\" & "Gain & Loss.rpt"
'            CR.ParameterFields(3) = "report_header;" & "  ﬁ—Ì— »«·√—»«Õ Ê«·Œ”«∆— ›Ì «·› —…" & "(" & headerdate(Me.DTPickerAccFrom) & ")" & " ≈·Ï «·› —… (" & headerdate(Me.DTPickerAccTo) & ")∫" & ";1"
'            CR.ReportTitle = RPTCompany_Name_Arabic
'            CR.ParameterFields(1) = "comment_arabic;" & RPTComment_Arabic & ";1"
'            CR.ParameterFields(0) = "name_english;" & RPTCompany_Name_Eng & ";1"
'            CR.ParameterFields(2) = "comment_english;" & RPTComment_Eng & ";1"
'
'            CR.ParameterFields(5) = "Mogmal;" & Mogmal_ & ";1"
'            CR.ParameterFields(6) = "generals;" & generals_ & ";1"
'            CR.ParameterFields(7) = "ehlak;" & ehlak_ & ";1"
'            CR.ParameterFields(8) = "discount_From;" & discount_From_ & ";1"
'            CR.ParameterFields(9) = "discount_to;" & discount_to_ & ";1"
'            CR.ParameterFields(4) = "other_income;" & other_income_ & ";1"
'
'            CR.WindowShowPrintSetupBtn = True
'            CR.WindowShowSearchBtn = True
'            CR.WindowTitle = RPTCompany_Name_Eng
'            CR.WindowState = crptMaximized
'            CR.Action = 1
'            CR.PageZoom (Zoom_Report)
'            Screen.MousePointer = 0
'            CR.Reset
'
'        End If
'            Me.TxtEhlak.text = ""
'            Me.eLE(3).Visible = False
'            Screen.MousePointer = 0
'        '==============================================================================
'    Case 4 '          («·„Ì“«‰Ì…)'ﬁ«∆„… «·„—ﬂ“ «·„«·Ï
'        SheetBalance
'    Case 5 '„Ì“«‰ «·„—«Ã⁄…
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.EndDate = Me.DtpSheet.Value
'        cAccountReport.ShowTrialBalance
'        Set cAccountReport = Nothing
'End Select
'Exit Sub
'ErrTrap:
'Screen.MousePointer = vbDefault
'Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
'Msg = Msg & Chr(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
'Msg = Msg & Chr(13) & "—ﬁ„ «·Œÿ« " & Err.Number
'Msg = Msg & Chr(13) & "‰’ «·Œÿ« " & Err.Description
'Msg = Msg & Chr(13) & "„’œ— «·Œÿ« " & Err.Source
'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'End Sub
'Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    Me.LblAccountName.Caption = StartSearch(Me.TreeView2, Me.TxtSearch.text, True)
'End If
'End Sub
'
'Private Sub SheetBalance()
'Dim EqupDep As Double
'Dim EqupCre As Double
'Dim GroundDep As Double
'Dim GroundCre As Double
'Dim BuildingDep As Double
'Dim BuildingCre As Double
'Dim ClientDep As Double
'Dim ClientCre As Double
'Dim BoxDep As Double
'Dim BoxCre As Double
'Dim BankDep As Double
'Dim BankCre As Double
'Dim CashDep As Double
'Dim CashCre As Double
''*******************************
'Dim CapitalDep As Double
'Dim CapitalCre As Double
'Dim AccCurrentDep As Double
'Dim AccCurrentCre As Double
'Dim SuppDep As Double
'Dim SuppCre As Double
'Dim PayNotesDep As Double
'Dim PayNotesCre As Double
'Dim LoanDep As Double
'Dim LoanCre As Double
'Dim OtherCREDITDep As Double
'Dim OtherCREDITCre As Double
'Dim NET As Double
'Dim OtherDEPETDep As Double
'Dim OtherDEPETDCre As Double
'Dim DblItemStock As Double
'Dim StrSQLReport As String
'
'Dim openning_From As Double
'Dim purchase_From As Double
'Dim sallesback_From As Double
'Dim salles_to As Double
'Dim purchaseback_to As Double
'Dim ending_to As Double
'Dim Mogmal_ As Double
'Dim generals_ As Double
'Dim discount_From_ As Double
'Dim other_income_ As Double
'Dim discount_to_ As Double
'Dim ehlak_ As Double
'
'Dim Rs2 As New ADODB.Recordset
'If Me.TxtEhlak.text = "" Then
'    Screen.MousePointer = 0
'    Me.eLE(3).Visible = True
'    TxtEhlak.SetFocus
'    Exit Sub
'Else
'Screen.MousePointer = 11
'
''**********************«·√’Ê·
''√ÃÂ“… Ê„⁄œ«… '
''„œÌ‰
'StrSQLReport = "SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'"FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & _
'"ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'"NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'"WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a1' & '%' AND " & _
'"DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)"
'Call open_rs(StrSQLReport, Rs2)
'
'If IsNull(Rs2!SumValue) Then
'    EqupDep = 0
'Else
'    EqupDep = Rs2!SumValue
'End If
'Rs2.Close
''œ«∆‰
'StrSQLReport = "SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'"FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & _
'"ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'"NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'"WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a1' & '%' AND " & _
'"DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)"
'Call open_rs(StrSQLReport, Rs2)
'If IsNull(Rs2!SumValue) Then
'    EqupCre = 0
'Else
'    EqupCre = Rs2!SumValue
'End If
'Rs2.Close
''√—«÷Ì*********
''„œÌ‰
'StrSQLReport = "SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'"FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & _
'"ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'"NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'"WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a3' & '%' AND " & _
'"DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)"
'Call open_rs(StrSQLReport, Rs2)
'If IsNull(Rs2!SumValue) Then
'    GroundDep = 0
'Else
'    GroundDep = Rs2!SumValue
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a3' & '%' AND " & _
'    " DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 " & _
'    " AND NOTES.Note_Date <= #" & SQLDate(Me.DtpSheet.Value) & _
'    "# AND (NOTES.NotePosted=True)", Rs2)
'If IsNull(Rs2!SumValue) Then
'    GroundCre = 0
'Else
'    GroundCre = Rs2!SumValue
'End If
'Rs2.Close
'
''„»«‰Ì*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    " NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a4' & '%' AND " & _
'    " DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If IsNull(Rs2!SumValue) Then
'    BuildingDep = 0
'Else
'    BuildingDep = Rs2!SumValue
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a1a4' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If IsNull(Rs2!SumValue) Then
'    BuildingCre = 0
'Else
'    BuildingCre = Rs2!SumValue
'End If
'Rs2.Close
'
''⁄„·«¡*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code" & _
'    " Like 'a1a2a3' & '%' AND DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If IsNull(Rs2!SumValue) Then
'    ClientDep = 0
'Else
'    ClientDep = Rs2!SumValue
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & _
'    " ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a3' & '%' AND " & _
'    " DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    ClientCre = Rs2!SumValue
'Else
'    ClientCre = 0
'End If
'Rs2.Close
''Œ“Ì‰…*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    BoxDep = Rs2!SumValue
'Else
'    BoxDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    BoxCre = Rs2!SumValue
'Else
'    BoxCre = 0
'End If
'Rs2.Close
'
''»‰ﬂ*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS  " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    " NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a2' & '%' AND " & _
'    " DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    BankDep = Rs2!SumValue
'Else
'    BankDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a2' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    BankCre = Rs2!SumValue
'Else
'    BankCre = 0
'End If
'Rs2.Close
'
''√Ê—«ﬁ ﬁ»÷*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a4' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    CashDep = Rs2!SumValue
'Else
'    CashDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a2a4' & '%' AND DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    CashCre = Rs2!SumValue
'Else
'    CashCre = 0
'End If
'Rs2.Close
'
''√—’œ… „œÌ‰… √Œ—Ï*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'"FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'"ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON  " & _
'"NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'"WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a3' & '%' AND " & _
'"DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <=#" & _
'SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    OtherDEPETDep = Rs2!SumValue
'Else
'    OtherDEPETDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a1a3' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    OtherDEPETDCre = Rs2!SumValue
'Else
'    OtherDEPETDCre = 0
'End If
'Rs2.Close
''**********«·Œ’Ê„***********************
''  —«” «·„«·*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE " & _
'    "DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a1a1' & '%' AND DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    CapitalDep = Rs2!SumValue
'Else
'    CapitalDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a1a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    CapitalCre = Rs2!SumValue
'Else
'    CapitalCre = 0
'End If
'Rs2.Close
'
''   «·Ã«—Ì*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a1a2' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    AccCurrentDep = Rs2!SumValue  '«·Ã«—Ì
'Else
'    AccCurrentDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a1a2' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    AccCurrentCre = Rs2!SumValue
'Else
'    AccCurrentCre = 0  '
'End If
'Rs2.Close
'
''   „Ê—œÊ‰*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a3a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    SuppDep = Rs2!SumValue  '
'Else
'    SuppDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a3a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    SuppCre = Rs2!SumValue
'Else
'    SuppCre = 0  '
'End If
'Rs2.Close
'
''   √Ê—«ﬁ œ›⁄*********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a3a2' & '%' AND DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    PayNotesDep = Rs2!SumValue  '
'Else
'    PayNotesDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'"FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'"ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'"ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'"WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a3a2' & '%' AND " & _
'"DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    PayNotesCre = Rs2!SumValue
'Else
'    PayNotesCre = 0  '
'End If
'Rs2.Close
''ﬁ—Ê÷ *********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a4a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    LoanDep = Rs2!SumValue  'ﬁ—÷
'Else
'    LoanDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a4a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    LoanCre = Rs2!SumValue
'Else
'    LoanCre = 0  '
'End If
'Rs2.Close
'
''    √—’œ… œ«∆‰… √Œ—Ï *********
''„œÌ‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a5' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    OtherCREDITDep = Rs2!SumValue  '
'Else
'    OtherCREDITDep = 0
'End If
'Rs2.Close
''œ«∆‰
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a5' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 AND NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    OtherCREDITCre = Rs2!SumValue
'Else
'    OtherCREDITCre = 0  '
'End If
'Rs2.Close
'
''***************Õ”«» ’«›Ì —»Õ «·› —…***********************************
''%%%%%%%%%%%$$$$$$$&&&&&&&^^^^^^^^^^@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                '*************Õ”«» „Ã„· «·—»Õ √Ê «·Œ”«—… („ «Ã—…) 7
'    '—’Ìœ √Ê· «·„œ… ********************
'    openning_From = 0
'    '«·„‘ —Ì« ***********************
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    "NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a3a2' & '%' AND " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        purchase_From = Rs2!SumValue
'    Else
'        purchase_From = 0
'    End If
'    Rs2.Close
'    '„—œÊœ«   «·„»Ì⁄«  *********************
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a3a4' & '%' AND  " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        sallesback_From = Rs2!SumValue
'    Else
'        sallesback_From = 0
'        End If
'    Rs2.Close
'    '«·„»Ì⁄«  ***********************8
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a4a1' & '%' AND " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        salles_to = Rs2!SumValue
'    Else
'        salles_to = 0
'    End If
'    Rs2.Close
'    '„—œÊœ«  «·„‘ —Ì«  **************
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a4a3' & '%' AND " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        purchaseback_to = Rs2!SumValue
'    Else
'        purchaseback_to = 0
'    End If
'    Rs2.Close
'    '—’Ìœ ¬Œ— «·„œ…' ************
'    ending_to = 0
'    '„Ã„· —»Õ ÊŒ”«—…
'    Mogmal_ = Val(salles_to) + Val(purchaseback_to) + Val(ending_to) - Val(openning_From) - Val(purchase_From) - Val(sallesback_From)
'
'
'    ''*****************Õ”«» „’—Ê›«  ⁄„Ê„Ì…
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id  " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a3a1' & '%' AND " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        generals_ = Rs2!SumValue
'    Else
'        generals_ = 0
'    End If
'    Rs2.Close
'    ''*****************Õ”«» Œ’„ „”„ÊÕ »Â
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    " ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a3a5' & '%' AND " & _
'    "NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        discount_From_ = Rs2!SumValue
'    Else
'        discount_From_ = 0
'    End If
'    Rs2.Close
'    ''*****************Õ”«»  ≈Ì—«œ«  √Œ—Ï
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    " NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    " WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a4a2' & '%' AND " & _
'    " NOTES.Note_Date <=#" & SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        other_income_ = Rs2!SumValue
'    Else
'        other_income_ = 0
'    End If
'    Rs2.Close
'    ''*****************Õ”«» «·Œ’„ «·„ﬂ ”»
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    " ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
'    " NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a4a4' & '%' AND " & _
'    "NOTES.Note_Date <=#" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'    If Not IsNull(Rs2!SumValue) Then
'        discount_to_ = Rs2!SumValue
'    Else
'        discount_to_ = 0
'    End If
'    Rs2.Close
'    ''********************Õ”«» «·Â«·ﬂ
'    ehlak_ = Val(Me.TxtEhlak)
'    DblItemStock = GetItemEvaluation(Me.DtpSheet.Value)
'    '%%%%%%%%%^^&&**********(Õ”«» ’«›Ì «·—»Õ) **************
'    '_________________________________________________________
'
'    NET = (Val(Mogmal_) + Val(other_income_) + Val(discount_to_)) - (Val(generals_) + Val(ehlak_) + Val(discount_From_))
'
'    CR.ReportFileName = App.Path & "\Reports\" & "Sheet_balance.rpt"
'    CR.ParameterFields(3) = "report_header;" & "  ﬁ—Ì— »ﬁ«∆‹„… «·„—ﬂ“ «·„‹«·‹Ï ›Ì " & "" & headerdate(Me.DtpSheet.Value) & "" & ";1"
'    CR.ReportTitle = RPTCompany_Name_Arabic
'    CR.ParameterFields(1) = "comment_arabic;" & RPTComment_Arabic & ";1"
'    CR.ParameterFields(0) = "name_english;" & RPTCompany_Name_Eng & ";1"
'    CR.ParameterFields(2) = "comment_english;" & RPTComment_Eng & ";1"
'    CR.ParameterFields(4) = "EqupDep_;" & EqupDep & ";1"
'    CR.ParameterFields(5) = "EqupCre_;" & EqupCre & ";1"
'    CR.ParameterFields(6) = "GroundDep_;" & GroundDep & ";1"
'    CR.ParameterFields(7) = "GroundCre_;" & GroundCre & ";1"
'    CR.ParameterFields(8) = "BuildingDep_;" & BuildingDep & ";1"
'    CR.ParameterFields(9) = "BuildingCre_;" & BuildingCre & ";1"
'    CR.ParameterFields(10) = "ClientDep_;" & ClientDep & ";1"
'    CR.ParameterFields(11) = "ClientCre_;" & ClientCre & ";1"
'    CR.ParameterFields(12) = "BoxDep_;" & BoxDep & ";1"
'    CR.ParameterFields(13) = "BoxCre_;" & BoxCre & ";1"
'    CR.ParameterFields(14) = "BankDep_;" & BankDep & ";1"
'    CR.ParameterFields(15) = "BankCre_;" & BankCre & ";1"
'    CR.ParameterFields(16) = "CashDep_;" & CashDep & ";1"
'    CR.ParameterFields(17) = "CashCre_;" & CashCre & ";1"
'    CR.ParameterFields(18) = "CapitalDep_;" & CapitalDep & ";1"
'    CR.ParameterFields(19) = "CapitalCre_;" & CapitalCre & ";1"
'    CR.ParameterFields(20) = "AccCurrentDep_;" & AccCurrentDep & ";1"
'    CR.ParameterFields(21) = "AccCurrentCre_;" & AccCurrentCre & ";1"
'    CR.ParameterFields(22) = "SuppDep_;" & SuppDep & ";1"
'    CR.ParameterFields(23) = "SuppCre_;" & SuppCre & ";1"
'    CR.ParameterFields(24) = "PayNotesDep_;" & PayNotesDep & ";1"
'    CR.ParameterFields(25) = "PayNotesCre_;" & PayNotesCre & ";1"
'    CR.ParameterFields(26) = "LoanDep_;" & LoanDep & ";1"
'    CR.ParameterFields(27) = "LoanCre_;" & LoanCre & ";1"
'    CR.ParameterFields(28) = "OtherCREDITDep_;" & OtherCREDITDep & ";1"
'    CR.ParameterFields(29) = "OtherCREDITCre_;" & OtherCREDITCre & ";1"
'    CR.ParameterFields(30) = "NET_;" & NET & ";1"
'    CR.ParameterFields(31) = "OtherDEPETDep_;" & OtherDEPETDep & ";1"
'    CR.ParameterFields(32) = "OtherDEPETDCre_;" & OtherDEPETDCre & ";1"
'    CR.ParameterFields(33) = "ItemStock;" & DblItemStock & ";1"
'    Call SendCrystalSetting(CR)
'    Screen.MousePointer = 0
'    CR.Reset
'End If
'
'Me.TxtEhlak.text = ""
'Me.eLE(3).Visible = False
'Screen.MousePointer = 0
'End Sub
'
'Private Function GetItemEvaluation(SecondDate As Date, Optional FirstDate As Date = CDate("01/01/1000")) As Double
'Dim Rs As New ADODB.Recordset
'Dim StrSQL As String
'Dim AdCmd As New ADODB.Command
'Dim ParDate1 As New ADODB.Parameter
'Dim ParDate2 As New ADODB.Parameter
'Dim TempDate As Date
'Dim NET As Double
'StrSQL = "SELECT Sum( QryStockNet.StockNet) as ItemsNet" & _
'" FROM QryStockNet INNER JOIN ITEMS ON QryStockNet.Item_ID = ITEMS.Item_ID " & _
'" Where Items.ReEvaluation_Method=3"

'
'Set AdCmd.ActiveConnection = Cn
'TempDate = FirstDate
'Set ParDate1 = AdCmd.CreateParameter("Date1", adDate, adParamInput, , TempDate)
'TempDate = SecondDate
'Set ParDate2 = AdCmd.CreateParameter("Date2", adDate, adParamInput, , TempDate)
'AdCmd.Parameters.Append ParDate1
'AdCmd.Parameters.Append ParDate2
'AdCmd.CommandType = adCmdText
'AdCmd.CommandText = StrSQL
'Rs.CursorType = adOpenStatic
'Rs.Open AdCmd, , adOpenStatic, adLockReadOnly, adCmdText
'If Not (Rs.BOF Or Rs.EOF) Then
'    If Not IsNull(Rs("ItemsNet").Value) Then
'         NET = Rs("ItemsNet").Value
'    End If
'End If
'GetItemEvaluation = NET
'End Function
Private Sub ChangeLang()
    Label1.Caption = "Des"
    Label2.Caption = "OPr/Term"
    Me.Caption = "Accounting Reports"
    Me.MainTab.TabCaption(0) = "Financial Statements"
    'OptAccount(0).Caption = "Subsidiary For ..."
    OptAccount(0).Caption = "Statement Of Account..."
    OptAccount(1).Caption = "General Ledger For..."
    OptAccount(2).Caption = "Trade Report"
    OptAccount(3).Caption = "Income Statement"
    OptAccount(4).Caption = "Balance Sheet"
    OptAccount(5).Caption = "Trial Balance"
    OptAccount(7).Caption = "Print GL"
    OptAccount(8).Caption = "Print GL with Cost Center"
    OptAccount(9).Caption = "Cost Center Transactions"
    OptAccount(10).Caption = "Projects Transactions"
    Frame1.Caption = "Select Cost Center"
    Frame2.Caption = "Select Projects"

    OptAccount(6).Caption = "Print Chart of Accounts"
    Ele(1).Caption = "In"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
    lbl(3).Caption = "Enter Depreciation Value"
    CmdSeach.Caption = "Search"

    ISButton1.Caption = "Search"
    Cmd.Caption = "Exit"

End Sub

Private Sub Cmd_Click()
    Unload Me
End Sub

Private Sub CmdAccount_Click()
    Dim i As Integer
    Dim cAccountReport As ClsAccReports

    Screen.MousePointer = 11

    For i = 0 To Me.OptAccount.count - 1

        If Me.OptAccount(i).value = True Then Exit For
    Next i

    Select Case i

        Case 6
            Set cAccountReport = New ClsAccReports
            cAccountReport.ShowChartAccounts WindowTarget
            Set cAccountReport = Nothing

        Case 0

            If txt_mod_flag.text = "N" Then

                'Õ”«» «” «– „”«⁄œ
                If Me.TrvAccounts.SelectedItem Is Nothing Then
                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                ElseIf InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
                    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If

                StrAccountCode = Me.TrvAccounts.SelectedItem.key
                StrAccountName = Me.TrvAccounts.SelectedItem.text
            End If
            
            If Text1.text = "" Then
                Set cAccountReport = New ClsAccReports
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowLedger StrAccountCode, StrAccountName
                Set cAccountReport = Nothing
            Else
                  
                Set cAccountReport = New ClsAccReports
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, Text1.text
                Set cAccountReport = Nothing
            
            End If

        Case 1

            'Õ”«» «” «– ⁄«„
            If Me.TrvAccounts.SelectedItem Is Nothing Then
                Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Set cAccountReport = New ClsAccReports
            StrAccountCode = Me.TrvAccounts.SelectedItem.key
            StrAccountCode = Mid$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            cAccountReport.ShowMaterLedgar StrAccountCode, Me.TrvAccounts.SelectedItem.text
            Set cAccountReport = Nothing

        Case 3
            'ﬁ«∆„… «·œŒ·
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowIncomeStatment
            Set cAccountReport = Nothing

        Case 4
            '«·„Ì‹“«‰‹Ì‹…
            Dim x As Integer

            If SystemOptions.UserInterface = ArabicInterface Then
                x = val(InputBox("Õœœ «·„” ÊÌ"))
            Else
                x = val(InputBox("Specify Level"))
            End If

            account_level = val(x)
            Set cAccountReport = New ClsAccReports
            cAccountReport.ShowBalanceSheet
            Set cAccountReport = Nothing

        Case 5
            '„Ì“«‰ „—«Ã⁄…
            Set cAccountReport = New ClsAccReports
 
            cAccountReport.ShowTrialBalance
            Set cAccountReport = Nothing

        Case 7
            'ﬁÌÊœ «·ÌÊ„Ì… ⁄‰ › —…
            ShowGL

        Case 8
            'ﬁÌÊœ «·ÌÊ„Ì…  Õ·Ì·Ì »„—«ﬂ“ «· ﬂ·›… ⁄‰ › —…
            ShowGLWITH_Cost_center

        Case 9
            ShowTransactionsWith_Cost_center StrAccountCode, DcCostCenter.BoundText

        Case 10
            ShowGLto_project val(Me.dcprojects.BoundText), Dcdetails.BoundText
    End Select

End Sub
 
Function ShowTransactionsWith_Cost_center(Optional Account_Code As String = "", Optional cost_center_id As String = "")
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = "SELECT     *, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN CC_ValIe * 1 ELSE 0 END, DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN CC_ValIe * 1 ELSE 0 END FROM    GL_CC where not(cost_center_id is null)"
    'MySQL = "Select * From GL_CC where not(cost_center_id is null)"

    If Account_Code = "" And cost_center_id <> "" Then
        MySQL = MySQL + " and cost_center_id='" & cost_center_id & "'"
    ElseIf Account_Code <> "" And cost_center_id = "" Then
        MySQL = MySQL + " and account_code='" & Account_Code & "'"
    ElseIf Account_Code <> "" And cost_center_id <> "" Then
        MySQL = MySQL + " and account_code='" & Account_Code & "' and cost_center_id='" & cost_center_id & "'"
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    Dim x As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        x = MsgBox("Â·  —Ìœ  ﬁ—Ì—  ›’Ì·Ì ‰⁄„ «„ ·«", vbExclamation + vbYesNo)
    Else
        x = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
    End If

    If x = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_centerE.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totals.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalse.rpt"
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
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
Function ShowGLWITH_Cost_center()
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From GL_CC "
 
    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    Dim x As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        x = MsgBox("Â·  —Ìœ ÿ»«⁄Â ﬂ· ﬁÌœ ›Ì ’›Õ…", vbExclamation + vbYesNo)
    Else
        x = MsgBox("Print Each Voucher in seprate Page", vbExclamation + vbYesNo)
    End If

    If x = vbNo Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL_cc.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL_ccE.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL_cc1.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL_ccE1.rpt"
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
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Function ShowGLto_project(project_id As Integer, Optional fullcode As String = "")
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_sub_projects where project_id=" & project_id

    If project_id = 0 Then
        MySQL = "Select * From RptLedger_sub_projects "
    End If

    If fullcode <> "" Then
        MySQL = MySQL + " and opr_fullcode='" & fullcode & "'"
        Dim sql As String
        Dim rsvalue As New ADODB.Recordset
        Dim opr_expected_value As Double
        sql = "select total from projects_des  where fullcode='" & fullcode & "'"
        rsvalue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rsvalue.RecordCount > 0 Then
            opr_expected_value = IIf(IsNull(rsvalue("total").value), 0, rsvalue("total").value)
        Else
            rsvalue.Close
            sql = "select total from terms_operations  where fullcode='" & fullcode & "'"
            rsvalue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rsvalue.RecordCount > 0 Then
                opr_expected_value = IIf(IsNull(rsvalue("total").value), 0, rsvalue("total").value)
            End If
        End If
 
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL _with_projects.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL _with_projectse.rpt"
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
            Msg = "No data to view"
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
   
        StrReportTitle = " ﬁ—Ì— „‘—Ê⁄ " '& StrAccountName

        If fullcode <> "" Then
            If SystemOptions.Items_or_operation = 0 Then
                StrReportTitle = " ﬁ—Ì— „‘—Ê⁄ " + dcprojects + " Œ«’ »«·»‰œ " + Me.Dcdetails.text
            ElseIf SystemOptions.Items_or_operation = 1 Then
                StrReportTitle = " ﬁ—Ì— „‘—Ê⁄ " + dcprojects + " Œ«’ »«·⁄„·Ì… " + Me.Dcdetails.text
            End If
        End If

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue opr_expected_value
        StrReportTitle = ""

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Function ShowGL()
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_Sub "
 
    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    Dim x As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        x = MsgBox("Â·  —Ìœ ÿ»«⁄Â ﬂ· ﬁÌœ ›Ì ’›Õ…", vbExclamation + vbYesNo)
    Else
        x = MsgBox("Print Each Voucher In Seprate Page ", vbExclamation + vbYesNo)
    End If

    If x = vbNo Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL_Eng.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL1.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL1_Eng .rpt"
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
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub dcprojects_Click(Area As Integer)
    Dim StrSQL  As String
    Dim project_id As Integer
    project_id = val(dcprojects.BoundText)

    If SystemOptions.Items_or_operation = 1 Then
        StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
          
    ElseIf SystemOptions.Items_or_operation = 0 Then
        StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
            
    End If

    fill_combo Me.Dcdetails, StrSQL
End Sub

'End Sub
Private Sub Ele_Click(Index As Integer)

End Sub

Private Sub Form_Load()
    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""

    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    Dim StrSQL As String
    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

    StrSQL = "  SELECT id ,Project_name FROM projects"
    fill_combo Me.dcprojects, StrSQL

    ModTree.LoadTreeAccount Me.TrvAccounts
    Me.TrvAccounts.Nodes("r").EnsureVisible
    Me.TrvAccounts.Nodes("r").Expanded = True
    Me.TrvAccounts.Nodes("r").Selected = True

    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub

Private Sub ISButton1_Click()
    txt_mod_flag.text = "S"
 
    Account_search.Show
    Account_search.case_id = 1
End Sub

Public Function Set_account_code(code As String, _
                                 name As String)
    StrAccountCode = code
    StrAccountName = name
    Me.LblAccountName.Caption = name

End Function

Private Sub OptAccount_Click(Index As Integer)

    Select Case Index

        Case 0
            Me.Ele(1).Visible = True
        
        Case 5
            Me.Ele(1).Visible = True

        Case 4
            Me.Ele(1).Visible = True

        Case 3
            Me.Ele(1).Visible = True
    
        Case 1
            Me.Ele(1).Visible = True

        Case 7
            Me.Ele(1).Visible = True

        Case 8
            Me.Ele(1).Visible = True

        Case 9
            Me.Ele(1).Visible = True
            Frame1.Visible = True
            Frame2.Visible = False

        Case 10
            Me.Ele(1).Visible = True
            Frame2.Visible = True
            Frame1.Visible = False
    
    End Select

End Sub

Private Sub TrvAccounts_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)

    'If InStr(Me.TrvAccounts.SelectedItem.Tag, "last") Then
    ''    If Me.OptAccount(0).Value = True Then Me.CmdAccount.Enabled = True
    ''    If Me.OptAccount(1).Value = True Then Me.CmdAccount.Enabled = False
    ''    If Button = 2 Then
    '''        MDIFrmamin.SubmasterMnu(0).Enabled = True
    '''        MDIFrmamin.SubmasterMnu(1).Enabled = True
    '''        MDIFrmamin.SubmasterMnu(2).Enabled = False
    '''        MDIFrmamin.PopupMenu MDIFrmamin.reportMnu
    ''    End If
    ''Else
    ''    If Me.OptAccount(1).Value = True Then Me.CmdAccount.Enabled = True
    ''    If Me.OptAccount(0).Value = True Then Me.CmdAccount.Enabled = False
    ''    If Button = 2 Then   'And Me.OptAccount(1).Value = True
    '''        MDIFrmamin.SubmasterMnu(0).Enabled = False
    '''        MDIFrmamin.SubmasterMnu(1).Enabled = False
    '''        MDIFrmamin.SubmasterMnu(2).Enabled = True
    '''        MDIFrmamin.PopupMenu MDIFrmamin.reportMnu
    ''    End If
    'End If
End Sub

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
    'On Error Resume Next
    Me.LblAccountName.Caption = Me.TrvAccounts.SelectedItem.text
    txt_mod_flag.text = "N"
    Dim StrAccountCode As String
    Dim Depit As Double
    Dim Credit As Double
    StrAccountCode = Me.TrvAccounts.SelectedItem.key

    If InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
        StrAccountCode = Me.TrvAccounts.SelectedItem.key
        StrAccountCode = Mid(StrAccountCode, 1, Len(StrAccountCode) - 1)
    Else
        StrAccountCode = Me.TrvAccounts.SelectedItem.key
    End If

    get_balanceFromGlNew StrAccountCode, , , True, Me.Bdate.value, Me.EndDate.value, Depit, Credit
    Me.TxtDept = Depit
    Me.TxtCredit = Credit
End Sub

Private Sub TxtDept_Click()
    viewReport
End Sub

Function viewReport()

    If Me.TrvAccounts.SelectedItem Is Nothing Then
        Msg = "ÌÃ» ≈Œ Ì«— «”„ «·Õ”«» «·›—⁄Ï" & Chr(13) & "«·„—«œ ⁄—÷ «· ﬁ—Ì— ·Â „‰ Œ·«· «·œ·Ì· «·„Õ«”»Ï"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
                        
    End If

    If InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
        StrAccountCode = Me.TrvAccounts.SelectedItem.key
        StrAccountCode = Mid(StrAccountCode, 1, Len(StrAccountCode) - 1)
    Else
        StrAccountCode = Me.TrvAccounts.SelectedItem.key
    End If
            
    StrAccountName = Me.TrvAccounts.SelectedItem.text
             
    Set cAccountReport = New ClsAccReports
    cAccountReport.BegineDate = Me.DTPickerAccFrom.value
    cAccountReport.EndDate = Me.DTPickerAccTo.value
    cAccountReport.ShowLedger2 StrAccountCode, StrAccountName
    Set cAccountReport = Nothing

End Function

