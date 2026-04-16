VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProjectsReports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹—  «·„‘«—Ì⁄"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   HelpContextID   =   470
   Icon            =   "frmProjectsReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   10485
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
      Height          =   8625
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10485
      _cx             =   18494
      _cy             =   15214
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
      _GridInfo       =   $"frmProjectsReports.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton Cmd 
         Height          =   480
         Left            =   30
         TabIndex        =   2
         Top             =   8115
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
         Height          =   8070
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   10425
         _cx             =   18389
         _cy             =   14235
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
            Height          =   7890
            Left            =   90
            TabIndex        =   3
            Top             =   90
            Width           =   10245
            _cx             =   18071
            _cy             =   13917
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
            Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹—  «·„‘«—Ì⁄"
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
               Height          =   7515
               Index           =   0
               Left            =   45
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   45
               Width           =   10155
               _cx             =   17912
               _cy             =   13256
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
               Begin C1SizerLibCtl.C1Elastic DcboBankName 
                  Height          =   7275
                  Index           =   2
                  Left            =   90
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   10425
                  _cx             =   18389
                  _cy             =   12832
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
                     Caption         =   " ﬁ—Ì— ··„Ê«œ «·„‰’—›… Ê«·„” ·„… ··„‘—Ê⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   23
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   5640
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  ﬁ—Ì— »÷„«‰ «·«⁄„«· ÿ»ﬁ« ··„‘—Ê⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   22
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   5220
                     Width           =   4380
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "  ﬁ—Ì— »÷„«‰ «·«⁄„«· ÿ»ﬁ« ··⁄„Ì·"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   21
                     Left            =   5430
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   4920
                     Width           =   4380
                  End
                  Begin VB.CheckBox chkIsPand 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ì »⁄ ··»‰œ"
                     Height          =   375
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   2340
                     Width           =   1575
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„⁄—›… «·„Ê«œ «·„ÿ·Ê»… ··„‘—Ê⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   20
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   4620
                     Width           =   4380
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Ê«œ Ê«·»‰Êœ «·„‰›–… ··„‘«—Ì⁄ . (ÿ»ﬁ« ··„” Œ·’« )"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   19
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   4320
                     Width           =   4380
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„«·… «·Õ«·Ì… ··„‘—Ê⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   18
                     Left            =   6300
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   2760
                     Width           =   3480
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „‘—Ê⁄"
                     CausesValidation=   0   'False
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   17
                     Left            =   6180
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   360
                     Width           =   3600
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«—’œ… «·„œ›Ê⁄«  «·„ﬁœ„…"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   16
                     Left            =   6300
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   2400
                     Width           =   3480
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "»Ì«‰«  «·„‘«—Ì⁄"
                     CausesValidation=   0   'False
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   15
                     Left            =   6180
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   2040
                     Width           =   3600
                  End
                  Begin VB.ComboBox billto 
                     DataSource      =   "Adodc1"
                     Height          =   315
                     ItemData        =   "frmProjectsReports.frx":040F
                     Left            =   120
                     List            =   "frmProjectsReports.frx":0419
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   3150
                     Width           =   3975
                  End
                  Begin MSComCtl2.DTPicker Todate1 
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
                     Left            =   7350
                     TabIndex        =   60
                     ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                     Top             =   6540
                     Visible         =   0   'False
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   231145475
                     CurrentDate     =   37357
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ—ﬂ… «·⁄„«·…"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   14
                     Left            =   6060
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   1680
                     Width           =   3720
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„ «»⁄… „Êﬁ› „” Œ·’«  «·„‘«—Ì⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   435
                     Index           =   13
                     Left            =   5340
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   3480
                     Width           =   4440
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ›«’Ì· —Ê« »  «·„‘«—Ì⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   12
                     Left            =   6060
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   1320
                     Width           =   3720
                  End
                  Begin VB.TextBox TxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   2850
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   3870
                     Width           =   1245
                  End
                  Begin VB.Frame Fra 
                     BackColor       =   &H00E2E9E9&
                     BorderStyle     =   0  'None
                     Height          =   375
                     Index           =   26
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   3840
                     Width           =   1545
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   ">"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   360
                        Index           =   3
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   49
                        ToolTipText     =   "«ﬂ»— „‰"
                        Top             =   0
                        Width           =   465
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "="
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   360
                        Index           =   4
                        Left            =   480
                        RightToLeft     =   -1  'True
                        TabIndex        =   48
                        ToolTipText     =   "Ì”«ÊÏ"
                        Top             =   0
                        Value           =   -1  'True
                        Width           =   495
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "<"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   5
                        Left            =   960
                        RightToLeft     =   -1  'True
                        TabIndex        =   47
                        ToolTipText     =   "«’€— „‰"
                        Top             =   0
                        Width           =   555
                     End
                  End
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«ŸÂ«— «·œ›⁄«  «·„ﬁœ„…"
                     Height          =   255
                     Index           =   11
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   4920
                     Width           =   2175
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„ﬁ»Ê÷«  «·„‘«—Ì⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   11
                     Left            =   6420
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   4020
                     Width           =   3360
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "›Ê« Ì— «·„‘«—Ì⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   10
                     Left            =   6300
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   3240
                     Width           =   3480
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  «·„” Œ·’« "
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   9
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   6570
                     Value           =   -1  'True
                     Width           =   5820
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   " Õ·Ì·Ì „‘—Ê⁄ „⁄Ì‰"
                     CausesValidation=   0   'False
                     Height          =   255
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   5700
                     Visible         =   0   'False
                     Width           =   1695
                  End
                  Begin VB.Frame Frame1 
                     Caption         =   "‰Ê⁄ «·„‘—Ê⁄"
                     Height          =   495
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   480
                     Width           =   3975
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "«·ﬂ·"
                        Height          =   195
                        Index           =   2
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   40
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   1215
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "≈›  «ÕÌ"
                        Height          =   195
                        Index           =   1
                        Left            =   1560
                        RightToLeft     =   -1  'True
                        TabIndex        =   39
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.OptionButton Opt 
                        Alignment       =   1  'Right Justify
                        Caption         =   "ÃœÌœ"
                        Height          =   195
                        Index           =   0
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   38
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin VB.CheckBox CHkEmployee 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„ÊŸ›/·„‰œÊ» „⁄Ì‰"
                     Height          =   375
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   2670
                     Width           =   1695
                  End
                  Begin VB.CheckBox chkbranch 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·›—⁄ „⁄Ì‰"
                     Height          =   375
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   960
                     Width           =   1575
                  End
                  Begin VB.CheckBox chkCustomers1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„ﬁ«Ê· »«ÿ‰ „⁄Ì‰"
                     Height          =   375
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   2010
                     Width           =   1575
                  End
                  Begin VB.CheckBox chkCustomers 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·⁄„Ì· „⁄Ì‰"
                     Height          =   375
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   1680
                     Width           =   1575
                  End
                  Begin VB.CheckBox ChkMain 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ «· ﬁ—Ì— ÿ»ﬁ« ··ÊÕœ… «·ﬂ»—Ï"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     ToolTipText     =   " ” Œœ„ ··„ƒ””«  «· Ì  »Ì⁄ »ÊÕœ… Ê«Õœ… ›ﬁÿ ··’„› «·Ê«Õœ"
                     Top             =   -840
                     Width           =   2535
                  End
                  Begin VB.CheckBox chkProjects 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„‘—Ê⁄ „⁄Ì‰"
                     Height          =   375
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   1320
                     Width           =   1575
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  «·„‘«—Ì⁄"
                     CausesValidation=   0   'False
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   8
                     Left            =   6180
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   600
                     Width           =   3600
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
                     Left            =   8640
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   3000
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·„‘«—Ì⁄ «·„ ⁄«ﬁœ ⁄·ÌÂ«"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   5
                     Left            =   5700
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   960
                     Width           =   4080
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   6060
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.TextBox Txt_order_no 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   5760
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÿ·»Ì«  «· Ì ·„  ”·„ Õ Ï «·«‰"
                     Height          =   195
                     Index           =   4
                     Left            =   8400
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   3600
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  ”·Ì„ «·ÿ·»Ì« "
                     Height          =   195
                     Index           =   3
                     Left            =   8880
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   3585
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«’‰«› «·„‰ Ã…  „‰ ÿ·»Ì… „⁄Ì‰"
                     Height          =   195
                     Index           =   2
                     Left            =   7200
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
                     Top             =   5640
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
                     Left            =   120
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Top             =   4680
                     Width           =   2475
                     _cx             =   4366
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
                        Format          =   231145475
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
                        Format          =   231145475
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
                     Left            =   240
                     TabIndex        =   17
                     Top             =   5760
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
                     ButtonImage     =   "frmProjectsReports.frx":0435
                     ColorButton     =   14871017
                     ColorHoverText  =   16777215
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16777215
                  End
                  Begin MSDataListLib.DataCombo DCProjects 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   26
                     Top             =   1320
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcCustomers 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   30
                     Top             =   1680
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcCustomers1 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   31
                     Top             =   2010
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo Dcbranch 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   33
                     Top             =   960
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboEmp 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   36
                     Top             =   2730
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboBox 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   52
                     Top             =   4320
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboBankN 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   55
                     Top             =   3480
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker FrmDate1 
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
                     Left            =   7350
                     TabIndex        =   59
                     ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                     Top             =   6180
                     Visible         =   0   'False
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   231145475
                     CurrentDate     =   37357
                  End
                  Begin MSDataListLib.DataCombo cmbPands 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   70
                     Top             =   2340
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label18 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„” Œ·’ «·Ï"
                     ForeColor       =   &H00000000&
                     Height          =   375
                     Left            =   4200
                     TabIndex        =   62
                     Top             =   3120
                     Width           =   1155
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·»‰ﬂ"
                     Height          =   375
                     Index           =   1
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   3510
                     Width           =   915
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·Œ“‰…"
                     Height          =   375
                     Index           =   62
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   4350
                     Width           =   915
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬁÌ„… «·„»·€"
                     Height          =   375
                     Index           =   61
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   3900
                     Width           =   915
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
                     Top             =   6060
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
                     Top             =   5760
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label LblAccountName 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0C8C0&
                     Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹—  «·„‘«—Ì⁄"
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
                     Top             =   75
                     Width           =   6150
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmProjectsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String

Private Sub ChangeLang()
    'Label1.Caption = "Des"
OptAccount(12).Caption = "Project Details"
Frame1.Caption = "Project Status"
opt(0).Caption = "All"
opt(1).Caption = "New"
opt(2).Caption = "Opening"
OptAccount(9).Visible = False

 Command1.Caption = "Project Details"

    Me.Caption = "Projects Reports"
    OptAccount(14).Caption = "Employee"
  LblAccountName.Caption = Me.Caption
  OptAccount(13).Caption = "Projects Status   ..."
  OptAccount(10).Caption = "Projects Bills   ..."
    OptAccount(8).Caption = "Projects Data   ..."
    OptAccount(9).Caption = "Projects Bills   ..."
    OptAccount(5).Caption = "Projects Contract Data   ..."
    lbl(61).Caption = "Value"
    lbl(62).Caption = "Box"
    lbl(1).Caption = "Bank"
 OptAccount(11).Caption = "Receipts Projects"
 MainTab.TabCaption(0) = Me.Caption
 chkbranch.Caption = "To Branch"
  chkProjects.Caption = "To Project"
  chkCustomers.Caption = "To Customer"
  chkCustomers1.Caption = "To Sub-Con."
CHkEmployee.Caption = "To Employee"
chk(11).RightToLeft = False
chk(11).Caption = "Advance Payments"

    OptAccount(6).Caption = "Print Chart of Accounts"
    Ele(1).Caption = "In"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
 
    Cmd.Caption = "Exit"

End Sub

Private Sub chkbranch_Click()
    If chkbranch.value = vbUnchecked Then
        DcBranch.text = ""
        DcBranch.Enabled = False
    Else
        DcBranch.Enabled = True
    End If
End Sub

Private Sub chkCustomers_Click()

    If chkCustomers.value = vbUnchecked Then
        dcCustomers.text = ""
        dcCustomers.Enabled = False
    Else
        dcCustomers.Enabled = True
    End If

End Sub

 

Private Sub chkCustomers1_Click()
    If chkCustomers1.value = vbUnchecked Then
        dcCustomers1.text = ""
        dcCustomers1.Enabled = False
    Else
        dcCustomers1.Enabled = True
    End If
End Sub

Private Sub CHkEmployee_Click()
    If CHkEmployee.value = vbUnchecked Then
        DcboEmp.text = ""
        DcboEmp.Enabled = False
    Else
        DcboEmp.Enabled = True
    End If
End Sub

Private Sub chkIsPand_Click()
  If chkIsPand.value = vbUnchecked Then
        cmbPands.text = ""
        cmbPands.Enabled = False
    Else
        cmbPands.Enabled = True
    End If
End Sub

Private Sub chkProjects_Click()

    If chkProjects.value = vbUnchecked Then
        dcprojects.text = ""
        dcprojects.Enabled = False
    Else
        dcprojects.Enabled = True
    End If

End Sub

Private Sub Cmd_Click()
    Unload Me
End Sub

 Sub ShowReportsPaymentsProj()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkbranch.value = vbChecked Then
 If val(DcBranch.BoundText) = 0 And DcBranch.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
  Else
    MsgBox "Please Select Branch"
   End If
   DcBranch.SetFocus
  Exit Sub
 End If
End If
If chkProjects.value = vbChecked Then
 If val(dcprojects.BoundText) = 0 And dcprojects.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
  Else
    MsgBox "Please Select Project"
   End If
   dcprojects.SetFocus
  Exit Sub
 End If
End If
If chkCustomers.value = vbChecked Then
 If val(dcCustomers.BoundText) = 0 And dcCustomers.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
  Else
    MsgBox "Please Select Customer"
   End If
   dcCustomers.SetFocus
  Exit Sub
 End If
End If

If CHkEmployee.value = vbChecked Then
 If val(DcboEmp.BoundText) = 0 And DcboEmp.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
  Else
    MsgBox "Please Select Employee"
   End If
   DcboEmp.SetFocus
  Exit Sub
 End If
End If

    
    sql = "SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.Notes.UserID, dbo.TblUsers.UserName, "
    sql = sql & "                  dbo.Notes.CashingType, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.TransactionTypes.TransactionTypeName,"
    sql = sql & "                  dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.Transaction_Type, dbo.Notes.RevenuesID, dbo.TblRevenuesTypes.RevenuesName,"
    sql = sql & "                  dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.AccountsCode, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmployeeFullcode,"
    sql = sql & "                  dbo.TblEmployee.Emp_Namee, dbo.Notes.EmpId, dbo.Notes.project_id, dbo.projects.Project_name, dbo.projects.Project_nameE,"
    sql = sql & "                 dbo.projects.Fullcode AS ProjectFullcode, dbo.Notes.BankID, dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.Notes.branch_no,"
    sql = sql & "                   dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.End_user_id, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    sql = sql & "                   dbo.TblCustemers.fullcode  , dbo.projects.EmpId1 , dbo.TblBoxesData.BoxNameE , dbo.projects.Pstate"
    sql = sql & "      FROM         dbo.TblBoxesData RIGHT OUTER JOIN"
    sql = sql & "                    dbo.TblRevenuesTypes RIGHT OUTER JOIN"
    sql = sql & "                   dbo.TblCustemers RIGHT OUTER JOIN"
    sql = sql & "                   dbo.projects ON dbo.TblCustemers.CusID = dbo.projects.End_user_id RIGHT OUTER JOIN"
    sql = sql & "                   dbo.BanksData RIGHT OUTER JOIN"
    sql = sql & "                   dbo.Transactions RIGHT OUTER JOIN"
    sql = sql & "                   dbo.TblUsers INNER JOIN"
    sql = sql & "                   dbo.Notes ON dbo.TblUsers.UserID = dbo.Notes.UserID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id ON dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID ON"
    sql = sql & "                   dbo.BanksData.BankID = dbo.Notes.BankID ON dbo.projects.id = dbo.Notes.project_id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID ON dbo.TblRevenuesTypes.RevenuesID = dbo.Notes.RevenuesID ON"
    sql = sql & "                   dbo.TblBoxesData.BoxID = dbo.Notes.BoxID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & " Where (dbo.Notes.NoteType = 4)"
    sql = sql + " AND  (Notes.CashingType = 5)"

    If chk(11).value = vbChecked Then
sql = sql + " and Notes.NCashingType=3"
End If
If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.Notes.branch_no =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
sql = sql & " and dbo.Notes.project_id =" & val(Me.dcprojects.BoundText) & ""
End If
If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If

If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
sql = sql & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
If Not (IsNull(DTPickerAccFrom.value)) Then
sql = sql & " and dbo.Notes.NoteDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If
If Not (IsNull(DTPickerAccTo.value)) Then
sql = sql & " and dbo.Notes.NoteDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
End If
If val(DcboBankN.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.Notes.BankID =" & val(Me.DcboBankN.BoundText) & ""
End If
If val(DcboBox.BoundText) <> 0 And DcboBox.text <> "" Then
sql = sql & " and dbo.Notes.BoxID =" & val(Me.DcboBox.BoundText) & ""
End If
  If val(Me.TxtValue.text) > 0 Then
        If Me.opt(5).value = True Then
                sql = sql + " AND Note_Value >" & val(Me.TxtValue.text) & ""
            End If
         If Me.opt(4).value = True Then
                sql = sql + " AND Note_Value =" & val(Me.TxtValue.text) & ""
            End If
          If Me.opt(3).value = True Then
                sql = sql + " AND Note_Value <" & val(Me.TxtValue.text) & ""
            End If
            
   End If
                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportPaymentProjects.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportPaymentProjects.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
     Else
     Msg = "NO Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Sub
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
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Sub
Public Function ShowReportsProject()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkbranch.value = vbChecked Then
 If val(DcBranch.BoundText) = 0 And DcBranch.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
  Else
    MsgBox "Please Select Branch"
   End If
   DcBranch.SetFocus
  Exit Function
 End If
End If
If chkProjects.value = vbChecked Then
 If val(dcprojects.BoundText) = 0 And dcprojects.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
  Else
    MsgBox "Please Select Project"
   End If
   dcprojects.SetFocus
  Exit Function
 End If
End If
If chkCustomers.value = vbChecked Then
 If val(dcCustomers.BoundText) = 0 And dcCustomers.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
  Else
    MsgBox "Please Select Customer"
   End If
   dcCustomers.SetFocus
  Exit Function
 End If
End If
If chkCustomers1.value = vbChecked Then
 If val(dcCustomers1.BoundText) = 0 And dcCustomers1.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— „ﬁ«Ê· «·»«ÿ‰"
  Else
    MsgBox "Please Select Subcontract"
   End If
   dcCustomers1.SetFocus
  Exit Function
 End If
End If
If CHkEmployee.value = vbChecked Then
 If val(DcboEmp.BoundText) = 0 And DcboEmp.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
  Else
    MsgBox "Please Select Employee"
   End If
   DcboEmp.SetFocus
  Exit Function
 End If
End If
If IsNull(DTPickerAccFrom.value) Or IsNull(DTPickerAccTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
Else
MsgBox "Please select Period"
End If
Exit Function
End If
    
    sql = " SELECT     dbo.projects.id, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.projects.Fullcode, dbo.projects.End_user_id, TblCustemers_1.CusName, "
    sql = sql & "                   TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS CusFullcode, dbo.projects.sub_contractor_id, TblCustemers_1.CusName AS SubconCusName,"
    sql = sql & "                  TblCustemers_1.CusNamee AS SubconCusNameE, TblCustemers_1.Fullcode AS Supcon, dbo.projects.branch_no, dbo.TblBranchesData.branch_name,"
    sql = sql & "                  dbo.TblBranchesData.branch_namee, dbo.projects.StartDate, dbo.projects.EmpId1, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode,"
    sql = sql & "                  dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Name1,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.projects.cost_after_discount, dbo.projects.project_cost,"
    sql = sql & "                  dbo.projects.DiscountPercentage,"
    sql = sql & "             dbo.projects.general_discount ,"
    sql = sql & " dbo.GetProjectBillValue(dbo.projects.id," & SQLDate(DTPickerAccFrom.value, True) & ", " & SQLDate(DTPickerAccTo.value, True) & ") AS BillValue,"
    sql = sql & "                  dbo.GetProjectPaymentValue(dbo.projects.id," & SQLDate(DTPickerAccFrom.value, True) & " , " & SQLDate(DTPickerAccTo.value, True) & ") AS PaymentValue"
    sql = sql & "     FROM         dbo.projects LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.projects.EmpId1 = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.projects.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID"
    sql = sql & "  WHERE     (NOT (dbo.projects.Fullcode = N'isnull'))"

If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.projects.branch_no =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
sql = sql & " and dbo.projects.id =" & val(Me.dcprojects.BoundText) & ""
End If
If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
sql = sql & " and dbo.projects.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
End If
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
sql = sql & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
If Not (IsNull(DTPickerAccFrom.value)) Then
sql = sql & " and dbo.projects.StartDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If
If Not (IsNull(DTPickerAccTo.value)) Then
sql = sql & " and dbo.projects.StartDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
End If

                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReStusProjectsReport.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReStusProjectsReportE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim Inpu As Integer
   Dim k As Integer
   If SystemOptions.UserInterface = ArabicInterface Then
   Msg = "Â·  —€» ›Ì «ŸÂ«— «·„‘«—Ì⁄ «· Ì ·Â« „” Œ·’«  ›ﬁÿ"
   Else
   Msg = "Show projects that have bills "
   End If
   Inpu = MsgBox(Msg, vbYesNo)
   If Inpu = vbYes Then
   k = 1
   Else
   k = 0
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
        '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue k
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function




Function print_reportProjectBill(Optional NoteSerial As Integer)
    
     On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
   'new
    
        MySQL = "SELECT         project_billl.PerforValue,discount1ID, discount1value,discount2value,PerforValue,advancedPayment,Discount,Results,advancedPayment,"
MySQL = MySQL & "                               dbo.project_billl.ID , dbo.project_billl.bill_date, dbo.project_billl.ManualNO,"
MySQL = MySQL & "                               dbo.project_billl.dueDate1 , dbo.project_billl.discount, dbo.project_billl.dueDate, dbo.project_billl.NoteSerial, dbo.project_billl.total, "

MySQL = MySQL & "                             CashCusValue =(  SELECT SUM(VALUE) FROM ProjectBillBuy WHERE  Bill_id  IN (SELECT Id FROM project_billl BB WHERE BB.project_no = project_billl.project_no )),"
 

MySQL = MySQL & "                             ProjectValue =(  SELECT SUM(TotalValue) FROM project_billl BB  WHERE BB.project_no = project_billl.project_no ),"
MySQL = MySQL & "                             advancedPaymen33t =(  SELECT SUM(advancedPayment) FROM project_billl BB  WHERE BB.project_no = project_billl.project_no ),"



MySQL = MySQL & "                           dbo.project_billl.Remarks, dbo.project_billl.Results, dbo.project_billl.advancedPayment, dbo.project_billl.discount2value, dbo.project_billl.discount1value, dbo.project_billl.bill_type, dbo.project_billl.project_no,"
MySQL = MySQL & "                                 dbo.projects.Fullcode, dbo.project_billl.project_name, dbo.project_billl.End_user_name, dbo.project_billl.Sub_user_name, dbo.project_billl.End_user_account, dbo.project_billl.bill_to, dbo.project_billl.Sub_user_account,"
MySQL = MySQL & "                                 dbo.project_billl.revenue_account, dbo.project_billl.subContractorId, dbo.TblCustemers.Address, dbo.TblCustemers.VATNO, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.project_billl.Branch_NO,"
MySQL = MySQL & "                                 dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.project_billl.discount1ID, dbo.project_billl.discount2ID, dbo.project_billl.note_id,"
MySQL = MySQL & "                                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.projects.REVENUE_account_balance, dbo.projects.Project_nameE,TblCustemers.CusID "


MySQL = MySQL & "        From TblBranchesData "
MySQL = MySQL & "               RIGHT OUTER JOIN project_billl"
MySQL = MySQL & "               LEFT OUTER JOIN TblCustemers"
MySQL = MySQL & "               RIGHT OUTER JOIN projects"
MySQL = MySQL & "                    ON  TblCustemers.CusID = projects.End_user_id"
MySQL = MySQL & "                    ON  project_billl.project_no = projects.id"
MySQL = MySQL & "                    ON  TblBranchesData.branch_id = project_billl.Branch_NO"



'26082015
'txtDiscount.Text = Round(IIf(IsNull(rs("Discount").value), 0, (rs("Discount").value)), Decimal_Places)
'Results.Text = Round(IIf(IsNull(rs("Results").value), 0, (rs("Results").value)), Decimal_Places)
'advancedPayment.Text = Round(IIf(IsNull(rs("advancedPayment").value), 0, (rs("advancedPayment").value)), Decimal_Places)

MySQL = MySQL & "  Where (iSnULL(advancedPayment,0) <> 0 oR iSnULL(PerforValue,0) <> 0)"


If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
    sql = sql & " and dbo.projects.branch_no =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
    sql = sql & " and dbo.projects.id =" & val(Me.dcprojects.BoundText) & ""
End If
If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
    sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
    sql = sql & " and dbo.projects.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
End If
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
    sql = sql & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
If Not (IsNull(DTPickerAccFrom.value)) Then
    sql = sql & " and dbo.projects.StartDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If
If Not (IsNull(DTPickerAccTo.value)) Then
    sql = sql & " and dbo.projects.StartDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
End If
MySQL = MySQL & sql

    'MySQL = MySQL + " order by project_bill_details.id"

  
    If OptAccount(21).value Then
        StrFileName = App.path & "\Reports\REPORTS NEW\rpt_ProjectsBills2.rpt"
    ElseIf OptAccount(22).value Then
        StrFileName = App.path & "\Reports\REPORTS NEW\rpt_ProjectsBills22.rpt"
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

    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng

    End If
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(3).AddCurrentValue GetRegVATNo(val(DcBranch.BoundText))
    End If

'xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.VATRegNo
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


Public Function ShowReportsEmpTransProject()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkbranch.value = vbChecked Then
 If val(DcBranch.BoundText) = 0 And DcBranch.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
  Else
    MsgBox "Please Select Branch"
   End If
   DcBranch.SetFocus
  Exit Function
 End If
End If
If chkProjects.value = vbChecked Then
 If val(dcprojects.BoundText) = 0 And dcprojects.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
  Else
    MsgBox "Please Select Project"
   End If
   dcprojects.SetFocus
  Exit Function
 End If
End If
If chkCustomers.value = vbChecked Then
 If val(dcCustomers.BoundText) = 0 And dcCustomers.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
  Else
    MsgBox "Please Select Customer"
   End If
   dcCustomers.SetFocus
  Exit Function
 End If
End If
If chkCustomers1.value = vbChecked Then
 If val(dcCustomers1.BoundText) = 0 And dcCustomers1.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— „ﬁ«Ê· «·»«ÿ‰"
  Else
    MsgBox "Please Select Subcontract"
   End If
   dcCustomers1.SetFocus
  Exit Function
 End If
End If
If CHkEmployee.value = vbChecked Then
 If val(DcboEmp.BoundText) = 0 And DcboEmp.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
  Else
    MsgBox "Please Select Employee"
   End If
   DcboEmp.SetFocus
  Exit Function
 End If
End If

    
    sql = " SELECT     dbo.projects.Project_name, dbo.projects_des.des, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID, "
    sql = sql & "                   dbo.projects_des.oprid, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.projects.Project_nameE, dbo.projects.Fullcode,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmpJobsTypes.JobTypeName AS HJobTypeName,"
    sql = sql & "                  dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
    sql = sql & "                  dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    sql = sql & "                  dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.opr_employee_details.*, dbo.projects.End_user_id AS Expr1,"
    sql = sql & "                  dbo.projects.sub_contractor_id AS Expr2"
    sql = sql & "  FROM         dbo.opr_employee_details LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmpDepartments ON dbo.opr_employee_details.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.opr_employee_details.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmpJobsTypes ON dbo.opr_employee_details.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.projects_des ON dbo.opr_employee_details.PandID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblProcessDEF ON dbo.opr_employee_details.OperID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
    sql = sql & "                  dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id"
    sql = sql & "  WHERE     (NOT (dbo.projects.Fullcode = N'isnull'))"
If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.opr_employee_details.BranchId =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
sql = sql & " and dbo.opr_employee_details.ProjectID =" & val(Me.dcprojects.BoundText) & ""
End If
If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
sql = sql & " and dbo.projects.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
End If
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
sql = sql & " and dbo.opr_employee_details.Emp_id =" & val(Me.DcboEmp.BoundText) & ""
End If
If Not (IsNull(DTPickerAccFrom.value)) And Not (IsNull(DTPickerAccTo.value)) Then
FrmDate1.value = DateAdd("d", -1, Me.DTPickerAccFrom.value)
ToDate1.value = DTPickerAccTo.value
sql = sql & " and dbo.opr_employee_details.FromDate between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
sql = sql & " and dbo.opr_employee_details.ToDate between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
ElseIf Not (IsNull(DTPickerAccTo.value)) Then
sql = sql & " and dbo.opr_employee_details.ToDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
ElseIf Not (IsNull(DTPickerAccFrom.value)) Then
sql = sql & " and dbo.opr_employee_details.FromDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If

                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReEmpTransProjectsReport.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReEmpTransProjectsReportE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim Inpu As Integer

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue k
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function

Public Function ShowReportsProjectsbillImplemented()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    
    sql = " SELECT T.Qty cost_after_discount,"
    sql = sql & "     T.ID ,"
    sql = sql & "          T.Project_name,T.Name Project_account,"
    sql = sql & "            T.Item Project_nameE,"
    sql = sql & "            SUM(ISNULL(T.quntExc, 0))     net,"
    sql = sql & "            total = ISNULL(T.Qty, 0) -"
    sql = sql & "            SUM (IsNull(t.quntExc, 0))"
    sql = sql & "     FROM   ("
    sql = sql & "                SELECT ProjectMainDes.Qty,projects.ID,"
    sql = sql & "                       ProjectMainDes.Name,"
    'sql = sql & "                       ProjectMainDes.ID,"
    sql = sql & "                       projects.Project_name,"
    sql = sql & "                       project_bill_details.item ,"
    sql = sql & "                       project_bill_details.quntExc"
    
    sql = sql & "   From projects"
    sql = sql & "               LEFT OUTER JOIN project_billl"
    sql = sql & "                    ON  projects.id = project_billl.project_no"
    sql = sql & "               LEFT OUTER JOIN project_bill_details"
    sql = sql & "                    ON  project_billl.id = project_bill_details.bill_id"
     sql = sql & "              LEFT OUTER JOIN ProjectMainDes"
   ' sql = sql & "                    --ON  projects.id = ProjectMainDes.ProjectID"
    sql = sql & "                    On project_bill_details.PrMainDesID = ProjectMainDes.ID"
'
'    sql = sql & "           From opr_employee_details"
'    sql = sql & "                  RIGHT OUTER JOIN ProjectMainDes"
'    sql = sql & "                       ON  opr_employee_details.Project_id = ProjectMainDes.ProjectID"
'    sql = sql & "                  LEFT OUTER JOIN project_billl"
'    sql = sql & "                  INNER JOIN project_bill_details"
'    sql = sql & "                       ON  project_billl.id = project_bill_details.bill_id"
'    sql = sql & "                       ON  ProjectMainDes.ProjectID = project_billl.project_no"
'    sql = sql & "                       AND ProjectMainDes.ID = project_bill_details.PrMainDesID"
'    sql = sql & "                  LEFT OUTER JOIN projects"
'    sql = sql & "                       ON  ProjectMainDes.ProjectID = projects.id"
    sql = sql & "                Where  1 = 1 "
    'sql = sql & "(ProjectMainDes.ProjectID = 30)"
    If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
    sql = sql & " and dbo.projects.branch_no =" & val(Me.DcBranch.BoundText) & ""
    End If
    If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
    sql = sql & " and projects.ID =" & val(Me.dcprojects.BoundText) & ""
    End If
    If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
    sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
    End If
    If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
    'sql = sql & " and dbo.projects.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
    
    
sql = sql & "  and projects_des.oprid in ("
sql = sql & "  SELECT     dbo.project_bill_details.oprid"
sql = sql & "  FROM         dbo.project_billl INNER JOIN"
sql = sql & "                       dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id"
sql = sql & " Where (dbo.project_billl.subContractorId = " & val(dcCustomers1.BoundText) & ")"
sql = sql & " )"

    End If
'    If val(DcboEmp.BoundText) <> 0 And DcboEmp.Text <> "" Then
'    sql = sql & " and dbo.opr_employee_details.Emp_id =" & val(Me.DcboEmp.BoundText) & ""
'    End If
    If Not (IsNull(DTPickerAccFrom.value)) And Not (IsNull(DTPickerAccTo.value)) Then
    FrmDate1.value = DateAdd("d", -1, Me.DTPickerAccFrom.value)
    ToDate1.value = DTPickerAccTo.value
    sql = sql & " and dbo.project_billl.bill_date between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
 '   sql = sql & " and dbo.project_billl.bill_date between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(Todate1.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccTo.value)) Then
    sql = sql & " and dbo.project_billl.bill_date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccFrom.value)) Then
    sql = sql & " and dbo.project_billl.bill_date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
   
    sql = sql & "            )                             T"
    sql = sql & "         Where IsNull(ID, 0) <> 0"
    sql = sql & "     Group By"
    sql = sql & "            T.Qty,"
    sql = sql & "            T.ID,Name,"
    sql = sql & "            T.Project_name,"
    sql = sql & "            t.Item"
    
    
    
sql = " SELECT p.Project_name ,p.Id,"
sql = sql & "               projects_des.Des Project_nameE,"
sql = sql & "               Qty cost_after_discount,"
sql = sql & "               net = ("
sql = sql & "                   SELECT SUM(quntExc)"
sql = sql & "                   From project_bill_details"
sql = sql & "                          INNER JOIN project_billl"
sql = sql & "                               ON  project_billl.id = project_bill_details.bill_id"
sql = sql & "                   Where project_billl.project_no = p.ID"
sql = sql & "                          AND project_bill_details.oprid = projects_des.oprid"

  
'    If val(DcboEmp.BoundText) <> 0 And DcboEmp.Text <> "" Then
'    sql = sql & " and dbo.opr_employee_details.Emp_id =" & val(Me.DcboEmp.BoundText) & ""
'    End If
    If Not (IsNull(DTPickerAccFrom.value)) And Not (IsNull(DTPickerAccTo.value)) Then
    FrmDate1.value = DateAdd("d", -1, Me.DTPickerAccFrom.value)
    ToDate1.value = DTPickerAccTo.value
    sql = sql & " and dbo.project_billl.bill_date between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
 '   sql = sql & " and dbo.project_billl.bill_date between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(Todate1.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccTo.value)) Then
    sql = sql & " and dbo.project_billl.bill_date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccFrom.value)) Then
    sql = sql & " and dbo.project_billl.bill_date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    

sql = sql & "               )"
sql = sql & "               ,total = ISNULL(Qty, 0) -"
sql = sql & "                ("
sql = sql & "                   SELECT SUM(quntExc)"
sql = sql & "                   From project_bill_details"
sql = sql & "                          INNER JOIN project_billl"
sql = sql & "                               ON  project_billl.id = project_bill_details.bill_id"
sql = sql & "                   Where project_billl.project_no = p.ID"
sql = sql & "                          AND project_bill_details.oprid = projects_des.oprid"
sql = sql & "               )"
sql = sql & "               ,"
sql = sql & "               Project_account  = ("
sql = sql & "                   SELECT NAME"
sql = sql & "                   From ProjectMainDes"
sql = sql & "                   Where ProjectMainDes.ID = projects_des.PrMainDesID"
sql = sql & "               )"
sql = sql & "        From projects_des"
sql = sql & "               INNER JOIN projects AS p"
sql = sql & "                    ON  p.id = projects_des.project_id"
    
If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
    sql = sql & " and p.branch_no =" & val(Me.DcBranch.BoundText) & ""
    End If
    If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
    sql = sql & " and p.ID =" & val(Me.dcprojects.BoundText) & ""
    End If
    If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
    sql = sql & " and p.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
    End If
    If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
'    sql = sql & " and p.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
    
        If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
    'sql = sql & " and dbo.projects.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
    
    
sql = sql & "  and projects_des.oprid in ("
sql = sql & "  SELECT     dbo.project_bill_details.oprid"
sql = sql & "  FROM         dbo.project_billl INNER JOIN"
sql = sql & "                       dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id"
sql = sql & " Where (dbo.project_billl.subContractorId = " & val(dcCustomers1.BoundText) & ")"
sql = sql & " )"

    End If
    
    End If
                       
    If val(cmbPands.BoundText) <> 0 And cmbPands.text <> "" Then
        sql = sql & " and projects_des.PanID =" & val(Me.cmbPands.BoundText) & ""
    End If
                                      
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "projectsbillImplemented.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "projectsbillImplemented.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim Inpu As Integer

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(5).AddCurrentValue dcCustomers1.text
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function


Public Function ShowReportsProjectsbillImplemented2()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    

    
    
sql = " SELECT"
sql = sql + "       T.Qty cost_after_discount,T.oprid,"
sql = sql + "                T.ID ,"
sql = sql + "           T.Project_name,T.des Project_account,Name  sub_contractor_name,"
sql = sql + "                      t.Remark Project_nameE From ("
sql = sql + " SELECT dbo.projects_des.fullcode,p.Project_name,p.Id,"
sql = sql + "       dbo.projects_des.[index],"
sql = sql + "       dbo.projects_des.des,"
sql = sql + "       dbo.projects_des.qty,"
sql = sql + "       dbo.projects_des.cost,"
sql = sql + "       dbo.projects_des.total,"
sql = sql + "       dbo.projects_des.discount,"
sql = sql + "       dbo.projects_des.net,"
sql = sql + "       dbo.projects_des.project_id,"
sql = sql + "       dbo.projects_des.sub_contractor_id,"
sql = sql + "       dbo.TblCustemers.CusName,"
sql = sql + "       dbo.projects_des.oprid,"
sql = sql + "       dbo.projects_des.Remark,"
sql = sql + "       dbo.projects_des.esQty,"
sql = sql + "       dbo.projects_des.PandUnitID,"
sql = sql + "       dbo.TblProcessUnites.UnitName,"
sql = sql + "       dbo.TblProcessUnites.UnitNamee,"
sql = sql + "       dbo.projects_des.QtyNo,"
sql = sql + "       dbo.projects_des.CodeBand,"
sql = sql + "       dbo.projects_des.PanID,"
sql = sql + "       dbo.TblPands.Name,"
sql = sql + "       dbo.TblPands.NameE,"
sql = sql + "       dbo.projects_des.TotalExe,"
sql = sql + "       dbo.projects_des.PriceExe,"
sql = sql + "       dbo.projects_des.QtyExe,"
sql = sql + "       dbo.projects_des.PrMainDesID"
sql = sql + "  From dbo.projects_des"

sql = sql + "        LEFT OUTER JOIN dbo.TblPands"
sql = sql + "             ON  dbo.projects_des.PanID = dbo.TblPands.ID"
sql = sql + "             INNER JOIN projects AS p ON p.id = dbo.projects_des.project_id"
sql = sql + "        LEFT OUTER JOIN dbo.TblProcessUnites"
sql = sql + "             ON  dbo.projects_des.PandUnitID = dbo.TblProcessUnites.UnitID"
sql = sql + "        LEFT OUTER JOIN dbo.TblCustemers"
sql = sql + "             ON  dbo.projects_des.sub_contractor_id = dbo.TblCustemers.CusID"
sql = sql + "  Where 1 = 1 "

    If val(cmbPands.BoundText) <> 0 And cmbPands.text <> "" Then
        sql = sql & " and projects_des.PanID =" & val(Me.cmbPands.BoundText) & ""
    End If
    'sql = sql & "(ProjectMainDes.ProjectID = 30)"
    If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
    sql = sql & " and p.branch_no =" & val(Me.DcBranch.BoundText) & ""
    End If
    If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
    sql = sql & " and p.ID =" & val(Me.dcprojects.BoundText) & ""
    End If
    If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
    sql = sql & " and p.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
    End If
    If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
    'sql = sql & " and p.sub_contractor_id =" & val(Me.dcCustomers1.BoundText) & ""
    sql = sql & "  and projects_des.oprid in ("
sql = sql & "  SELECT     dbo.project_bill_details.oprid"
sql = sql & "  FROM         dbo.project_billl INNER JOIN"
sql = sql & "                       dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id"
sql = sql & " Where (dbo.project_billl.subContractorId = " & val(dcCustomers1.BoundText) & ")"
sql = sql & " )"


    End If
'    If val(DcboEmp.BoundText) <> 0 And DcboEmp.Text <> "" Then
'    sql = sql & " and dbo.opr_employee_details.Emp_id =" & val(Me.DcboEmp.BoundText) & ""
'    End If
    If Not (IsNull(DTPickerAccFrom.value)) And Not (IsNull(DTPickerAccTo.value)) Then
    FrmDate1.value = DateAdd("d", -1, Me.DTPickerAccFrom.value)
    ToDate1.value = DTPickerAccTo.value
   ' sql = sql & " and projects.StartDate  between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(Todate1.value, True) & ""
   ' sql = sql & " and projects.EndDate  between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(Todate1.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccTo.value)) Then
   ' sql = sql & " and projects.EndDate  <=" & SQLDate(DTPickerAccTo.value, True) & ""
    ElseIf Not (IsNull(DTPickerAccFrom.value)) Then
   ' sql = sql & " and projects.StartDate  >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    
    sql = sql & "            )                             T"
    sql = sql & "             Order By T.oprid"
'    sql = sql & "     Group By"
'    sql = sql & "            T.Qty,"
'    sql = sql & "            T.ID,Name,"
'    sql = sql & "            T.Project_name,"
'    sql = sql & "            t.Item"
    
    

                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "projectsbillImplemented2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "projectsbillImplemented2.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim Inpu As Integer

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue k
     xReport.ParameterFields(5).AddCurrentValue dcCustomers1.text
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function





Public Function ShowReportsEmpInProject()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkbranch.value = vbChecked Then
 If val(DcBranch.BoundText) = 0 And DcBranch.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
  Else
    MsgBox "Please Select Branch"
   End If
   DcBranch.SetFocus
  Exit Function
 End If
End If
If chkProjects.value = vbChecked Then
 If val(dcprojects.BoundText) = 0 And dcprojects.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
  Else
    MsgBox "Please Select Project"
   End If
   dcprojects.SetFocus
  Exit Function
 End If
End If
    sql = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, "
    sql = sql & "                   dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
    sql = sql & "                   dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.project_id, dbo.projects.Fullcode AS ProjectCode, dbo.projects.Project_name,"
    sql = sql & "                   dbo.TblEmployee.BranchID , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    sql = sql & "         FROM         dbo.TblEmployee INNER JOIN"
    sql = sql & "                   dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
    sql = sql & "   where not (dbo.TblEmployee.project_id  is null) "
If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.TblEmployee.BranchId =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
sql = sql & " and dbo.TblEmployee.project_id =" & val(Me.dcprojects.BoundText) & ""
End If

                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReEmpInProject.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReEmpInProjectE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim Inpu As Integer

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue k
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function


Private Sub CmdAccount_Click()
  Dim Fromdate As Date
  Dim todate As Date

Dim StrSQL As String
    Dim i As Integer
    Dim cAccountReport As ClsReportViewer
    Dim StrDes As String
        StrDes = ""

    For i = 0 To Me.OptAccount.count - 1

        If Me.OptAccount(i).value = True Then Exit For
    Next i
 
    
            Screen.MousePointer = 11
            Set cAccountReport = New ClsReportViewer
 
    Select Case i
    Case 23
        printProjectByItems
    Case 16
    Dim StrAccountCode As String
   If IsNull(DTPickerAccFrom.value) Then
   Fromdate = "01/01/2000"
Else
 Fromdate = DTPickerAccFrom.value
End If


If IsNull(DTPickerAccTo.value) Then
 todate = Date
Else
 todate = DTPickerAccTo.value
End If


  If CHkEmployee.value = vbChecked And val(DcboEmp.BoundText) <> 0 Then
  
  StrAccountCode = "  select Account_Code2 from TblCustemers where      not ( Account_Code2 ) is null and Type=1  and    EmpId in(" & val(DcboEmp.BoundText) & ")"
  Else
  StrAccountCode = "select Account_Code2 from TblCustemers where   not ( Account_Code2 ) is null  and Type=1 "
  End If
      If chkCustomers.value = vbChecked And val(dcCustomers.BoundText) <> 0 Then
  
  StrAccountCode = StrAccountCode + "  and  CusID in(" & val(dcCustomers.BoundText) & ")"
  
  End If
  
  
           
   updateopeningbalanceNewFromsqlTrialBalance2 DTPickerAccFrom.value, DTPickerAccTo.value, True, , , StrAccountCode, 3
            Dim x1 As Integer
           Dim ShowGrouping As Integer
    Dim Showzeros As String
                    If SystemOptions.UserInterface = ArabicInterface Then
                    x1 = MsgBox("Â«  —Ìœ ⁄—÷ «·Õ”«»«  «·’›—Ì…   ", vbCritical + vbYesNo)
                Else
                    x1 = MsgBox("Show zero accounts Y :N  ", vbCritical + vbYesNo)
                End If

                    If x1 = vbYes Then
                        Showzeros = "1"
                    Else
                        Showzeros = "0"
                    End If
                        If SystemOptions.UserInterface = ArabicInterface Then
                    ShowGrouping = MsgBox("Â·  —Ìœ  ⁄„· ›—“ »«·„‰œÊ» ", vbCritical + vbYesNo)
                Else
                    ShowGrouping = MsgBox("Allow Grouping ", vbCritical + vbYesNo)
                End If
                            If ShowGrouping = vbNo Then
                            ShowGrouping = 0
                            Else
                            ShowGrouping = 1
                            End If
    '
    StrSQL = "SELECT     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.EmpId, "
 StrSQL = StrSQL & "                        ISNULL(dbo.ACCOUNTS.DepitBalance, 0) AS depit, ISNULL(dbo.ACCOUNTS.CreditBalance, 0) AS Credit, ISNULL(dbo.ACCOUNTS.opening_balance, 0) AS balance,    dbo.TblEmployee.Emp_Name,"
 StrSQL = StrSQL & "  dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.Type, dbo.TblCustemers.CustomerandVendor"
 StrSQL = StrSQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
 StrSQL = StrSQL & "   dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code2 = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
 StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TblCustemers.EmpId = dbo.TblEmployee.Emp_ID"
 StrSQL = StrSQL & " WHERE       (dbo.TblCustemers.Type = 1 OR"
 StrSQL = StrSQL & "   dbo.TblCustemers.CustomerandVendor = 1)"
 StrSQL = StrSQL & "  and ( dbo.TblCustemers.BranchId in(" & Current_branchSql & ") or Isnull(dbo.TblCustemers.BranchId,0) = 0)"
StrSQL = StrSQL + "  and  not ( TblCustemers.Account_Code2 ) is null"
  If CHkEmployee.value = vbChecked And val(DcboEmp.BoundText) <> 0 Then
  
  StrSQL = StrSQL + "  and  EmpId in(" & val(DcboEmp.BoundText) & ")"
  
  End If
  
    If chkCustomers.value = vbChecked And val(dcCustomers.BoundText) <> 0 Then
  
  StrSQL = StrSQL + "  and  CusID in(" & val(dcCustomers.BoundText) & ")"
  
  End If
  
  
 
     ' Dim Reports As ClsRepoerts
      Dim Reports As New ClsRepoerts
    Reports.ShowSallingTime StrSQL, DTPickerAccFrom.value, DTPickerAccTo.value, , 600, , CHR(13) & " «·œ›⁄«  «·„ﬁœ„…", "", , , , , "", Showzeros, , , ShowGrouping
    
         
        
        



    Case 15
                 If chkProjects.value = vbChecked Then
                  
                If Me.dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                 Else
                    Msg = "Please Select Project...!!" & CHR(13)
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            End If
    
            If chkCustomers.value = vbChecked Then
        
                If Me.dcCustomers.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· ...!!" & CHR(13)
                  Else
                   Msg = "Please Select Customer...!!" & CHR(13)
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            End If
    
    
 
            
            
    If chkbranch.value = vbChecked Then
        
                        If Me.DcBranch.BoundText = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ» «Œ Ì«—   «·›—⁄   ...!!" & CHR(13)
                            Else
                            Msg = "Please Select Brnach   ...!!" & CHR(13)
                            End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            DcBranch.SetFocus
                            Sendkeys "{F4}"
                            Exit Sub
                        End If
    StrDes = StrDes & "      " & DcBranch.text
            End If
   
     printProjectByCustomers


    Case 12
             If chkProjects.value = vbChecked Then
                  
                If Me.dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                 Else
                    Msg = "Please Select Project...!!" & CHR(13)
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            End If
    
            If chkCustomers.value = vbChecked Then
        
                If Me.dcCustomers.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· ...!!" & CHR(13)
                  Else
                   Msg = "Please Select Customer...!!" & CHR(13)
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            End If
    
    
            If chkCustomers1.value = vbChecked Then
        
                If Me.dcCustomers1.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— „ﬁ«Ê· «·»«ÿ‰   ...!!" & CHR(13)
                  Else
                  Msg = "Please Select Subcontract   ...!!" & CHR(13)
                  
                End If
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers1.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
        StrDes = StrDes & "      " & " «·„ﬁ«Ê· " & dcCustomers1.text

            End If
            
            
    If chkbranch.value = vbChecked Then
        
                        If Me.DcBranch.BoundText = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ» «Œ Ì«—   «·›—⁄   ...!!" & CHR(13)
                            Else
                            Msg = "Please Select Brnach   ...!!" & CHR(13)
                            End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            DcBranch.SetFocus
                            Sendkeys "{F4}"
                            Exit Sub
                        End If
    StrDes = StrDes & "      " & DcBranch.text
            End If
   
    printProjectDet


Case 9

  If Me.dcprojects.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                   dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
                
                
ShowReports
     Case 10
 ShowReportsBills
     Case 11
ShowReportsPaymentsProj
    Case 13
ShowReportsProject
    Case 14
ShowReportsEmpTransProject
    Case 18
ShowReportsEmpInProject
Case 19
ShowReportsProjectsbillImplemented
Case 20
ShowReportsProjectsbillImplemented2
Case 21, 22
print_reportProjectBill

        Case 8
                 
            If chkProjects.value = vbChecked Then
                  
                If Me.dcprojects.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            StrDes = StrDes & "      " & " «·„‘—Ê⁄ " & dcprojects.text
            End If
    
            If chkCustomers.value = vbChecked Then
        
                If Me.dcCustomers.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    StrDes = StrDes & "     " & " «·⁄„Ì· " & dcCustomers.text
            End If
    
    
            If chkCustomers1.value = vbChecked Then
        
                If Me.dcCustomers1.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— „ﬁ«Ê· «·»«ÿ‰   ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers1.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
        StrDes = StrDes & "      " & " «·„ﬁ«Ê· " & dcCustomers1.text

            End If
            
            
                      If chkbranch.value = vbChecked Then
        
                        If Me.DcBranch.BoundText = "" Then
                            Msg = "ÌÃ» «Œ Ì«—   «·›—⁄   ...!!" & CHR(13)
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            DcBranch.SetFocus
                            Sendkeys "{F4}"
                            Exit Sub
                        End If
    StrDes = StrDes & "      " & DcBranch.text
            End If
            
 
    Dim My_SQL2 As String
    Dim NetExpensen As Double

    Dim Balance As String
 
 

  
    
  
  If IsNull(DTPickerAccFrom.value) Then
   Fromdate = "01/01/2000"
Else
 Fromdate = DTPickerAccFrom.value
End If


If IsNull(DTPickerAccTo.value) Then
 todate = Date
Else
 todate = DTPickerAccTo.value
End If
    
    Dim openingbalacedate As Date
    getOpeningBalancedate , , , , year(todate), openingbalacedate, True

 
        StrSQL = " update projects"
     
     
          StrSQL = StrSQL & " set  expansesE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', expanses_account, 1 , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & " UnderAccountBalance=  isnull(dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', AccountUnderImp, 1 , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & " ),0)+ isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', AccountUnderImp,1),0),"
          StrSQL = StrSQL & "    expansesM_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Material_account, 1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & "    expansesS_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Salary_account, 1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & "    REVENUE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', REVENUE_account,  1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & " Legal_account_balance=  isnull(dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', legal, 1 , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0) + isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', legal,1),0),"
          StrSQL = StrSQL & "    CashingValue= dbo.GetProjectCashing('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',  id) "
          
          
    
          'CashingValue
          
          ' StrSQL = StrSQL & "    Legal_account_balance= dbo.GetBalance('" & SQLDate(fromdate) & "','" & SQLDate(todate) & "', legal, 1) + dbo.GetOpeningBalance(" &  SQLDate( openingbalacedate )" & ", legal,1)"
   StrSQL = StrSQL & "    where 1=1"
    
    
        If chkbranch.value = vbChecked Then
    StrSQL = StrSQL & " and   branch_no=" & val(DcBranch.BoundText)
    End If
    
    
    If chkProjects.value = vbChecked Then
    StrSQL = StrSQL & " and   id=" & val(dcprojects.BoundText)
    End If
    
    If chkCustomers.value = vbChecked Then
    StrSQL = StrSQL & " and   End_user_id='" & val(dcCustomers.BoundText) & "'"
    End If
    
    If chkCustomers1.value = vbChecked Then
    StrSQL = StrSQL & " and   sub_contractor_id='" & val(dcCustomers1.BoundText) & "'"
    End If
    
   If opt(0).value = True Then '  ÃœÌœ
       StrSQL = StrSQL & " and   (Pstate=0   or Pstate is null) "
    ElseIf opt(1).value = True Then '  «›  «ÕÌ
        StrSQL = StrSQL & " and   Pstate=1"
    End If
    
     Cn.CommandTimeout = 10000
   Cn.Execute StrSQL
        
  
 
 
    Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report

    Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
 '   sql = "SELECT * from projects where 1=1"
    sql = "SELECT     id, End_user_Account, End_user_name, sub_contractor_Account, sub_contractor_name, Fullcode, prifix, Code, Project_name, Contract_type, Project_status, "
sql = sql & "  expanses_account, REVENUE_account, Project_account, ISNULL(project_cost, 0) AS project_cost, branch_no, departement, project_code, branche_ID,"
sql = sql & " ISNULL(expanses_account_balance, 0) AS expanses_account_balance, ISNULL(expansese_account_balance, 0) AS expansese_account_balance,"
sql = sql & "  ISNULL(expansesm_account_balance, 0) AS expansesm_account_balance, ISNULL(expansess_account_balance, 0) AS expansess_account_balance,"
sql = sql & " ISNULL(REVENUE_account_balance, 0) AS REVENUE_account_balance, ISNULL(Project_account_balance, 0) AS Project_account_balance, Contract_type_name,"
sql = sql & " End_user_id, sub_contractor_id, ISNULL(general_discount, 0) AS general_discount, ISNULL(cost_after_discount, 0) AS cost_after_discount, ISNULL(total, 0) AS total,"
sql = sql & " ISNULL(sub_discount_total, 0) AS sub_discount_total, ISNULL(net, 0) AS net, Material_account, Salary_account, legal, ISNULL(items_total, 0) AS items_total,"
sql = sql & " ISNULL(Legal_account_balance, 0) AS Legal_account_balance, CurrencyID, StartDate, Project_nameE, opening_balance_voucher_id, OpenBalanceDate,"
sql = sql & "  OpenBalanceType, OpenBalance, OpenBalanceType1, OpenBalance1, OpenBalanceType2, OpenBalance2, OpenBalanceType3, OpenBalance3, OpenBalanceType4,"
sql = sql & "  OpenBalance4 , EmpID, EmpId1 ,ISNULL(UnderAccountBalance, 0) AS UnderAccountBalance"
sql = sql & " ,  ISNULL(CashingValue, 0) AS CashingValue   FROM         dbo.projects where 1=1 and not ( Fullcode is null) "
    
    
        If chkbranch.value = vbChecked Then
    sql = sql & " and   branch_no=" & val(DcBranch.BoundText)
    End If
    
    
    If chkProjects.value = vbChecked Then
    sql = sql & " and   id=" & val(dcprojects.BoundText)
    End If
    
    If chkCustomers.value = vbChecked Then
    sql = sql & " and   End_user_id='" & val(dcCustomers.BoundText) & "'"
    End If
    
    If chkCustomers1.value = vbChecked Then
    sql = sql & " and   sub_contractor_id='" & val(dcCustomers1.BoundText) & "'"
    End If
    
   If opt(0).value = True Then '  ÃœÌœ
       sql = sql & " and   (Pstate=0   or Pstate is null) "
    ElseIf opt(1).value = True Then '  «›  «ÕÌ
        sql = sql & " and   Pstate=1"
    End If
    
    
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\REPORT1A.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\REPORT1.rpt")
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = (App.path & "\reports\construction\REPORT1A.rpt")
    FrmReport.CRViewer.viewReport
  cAccountReport.CreateLogo xReport
   
If SystemOptions.UserInterface = ArabicInterface Then
xReport.reporttitle = cCompanyInfo.ArabCompanyName

       xReport.ParameterFields(1).AddCurrentValue Format(Fromdate, "DD/MM/YYYY")
     xReport.ParameterFields(2).AddCurrentValue Format(todate, "DD/MM/YYYY")
       xReport.ParameterFields(3).AddCurrentValue StrDes

Else

xReport.reporttitle = cCompanyInfo.EngCompanyName

'      xReport.ParameterFields(1).AddCurrentValue Format(FromDate, "DD/MM/YYYY")
'       xReport.ParameterFields(2).AddCurrentValue Format(Todate, "DD/MM/YYYY")
'         xReport.ParameterFields(3).AddCurrentValue StrDes
End If

    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
'    SendKeys "{RIGHT}"
                     
                     
   
Case 5
     Set cCompanyInfo = New ClsCompanyInfo

StrDes = " ﬁ—Ì— «·„‘«—Ì⁄ «·„ ⁄«ﬁœ ⁄·ÌÂ« "
            If CHkEmployee.value = vbChecked Then
        
                If Me.DcboEmp.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«—   «·„‰œÊ»   ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboEmp.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
        StrDes = StrDes & "      " & "«·„‰œÊ» " & DcboEmp.text

            End If


  If IsNull(DTPickerAccFrom.value) Then
   Fromdate = "01/01/2000"
Else
 Fromdate = DTPickerAccFrom.value
End If


If IsNull(DTPickerAccTo.value) Then
 todate = Date
Else
 todate = DTPickerAccTo.value
End If
    
 
   sql = "SELECT     dbo.projects.Pstate, dbo.projects.id, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.Notes.note_count, dbo.projects.StartDate, dbo.Notes.project_id, "
   sql = sql & "                    dbo.Notes.Emp_ID, dbo.Notes.EmpId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.Notes.NoteDate,"
   sql = sql & "                    dbo.Notes.Note_Value, dbo.projects.project_cost, dbo.RemainingValueProject(dbo.Notes.project_id) AS sumPayment"
   sql = sql & "  FROM         dbo.projects INNER JOIN"
   sql = sql & "                     dbo.Notes ON dbo.projects.id = dbo.Notes.project_id LEFT OUTER JOIN"
   sql = sql & "                     dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID"
   sql = sql & "  Where (dbo.Notes.NoteType = 4)"
       
           If chkProjects.value = vbChecked Then
                  
                If Me.dcprojects.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            StrDes = StrDes & "      " & " «·„‘—Ê⁄ " & dcprojects.text
            End If
    
            If chkCustomers.value = vbChecked Then
        
                If Me.dcCustomers.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    StrDes = StrDes & "     " & " «·⁄„Ì· " & dcCustomers.text
            End If
    
 
    
    If chkProjects.value = vbChecked Then
    sql = sql & " and   id=" & val(dcprojects.BoundText)
    End If
    
    If chkCustomers.value = vbChecked Then
    sql = sql & " and   End_user_id='" & val(dcCustomers.BoundText) & "'"
    End If
    
    
        If CHkEmployee.value = vbChecked Then
    sql = sql & " and   dbo.Notes.EmpId=" & val(DcboEmp.BoundText)
    End If
    
 
         If chkbranch.value = vbChecked Then
    sql = sql & " and   projects.branch_no=" & val(DcBranch.BoundText)
    End If
    
    
         If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            sql = sql + " and     StartDate >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If

        If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
            sql = sql + " and StartDate <=" & SQLDate(Me.DTPickerAccTo, True) & ""
        End If
        
        
        
   If opt(0).value = True Then '  ÃœÌœ
       sql = sql & " and   (Pstate=0   or Pstate is null) "
    ElseIf opt(1).value = True Then '  «›  «ÕÌ
        sql = sql & " and   Pstate=1"
    End If
        
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\projectsrevenue.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\projectsrevenue.rpt")
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
          xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
            xReport.ParameterFields(2).AddCurrentValue ""
            
           xReport.ParameterFields(4).AddCurrentValue Format(Fromdate, "DD/MM/YYYY")
        xReport.ParameterFields(5).AddCurrentValue Format(todate, "DD/MM/YYYY")
     xReport.reporttitle = StrDes
     
    FrmReport.txtPath = (App.path & "\reports\construction\projectsrevenue.rpt")
    FrmReport.CRViewer.viewReport
  cAccountReport.CreateLogo xReport
 

    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
'    SendKeys "{RIGHT}"
 
Case 17  'ﬂ‘› Õ”«» „‘—Ê⁄
 
     Set cCompanyInfo = New ClsCompanyInfo

StrDes = "ﬂ‘› Õ”«» „‘—Ê⁄"
            If CHkEmployee.value = vbChecked Then
        
                If Me.DcboEmp.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«—   «·„‰œÊ»   ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboEmp.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
        StrDes = StrDes & "      " & "«·„‰œÊ» " & DcboEmp.text

            End If


  If IsNull(DTPickerAccFrom.value) Then
   Fromdate = "01/01/2000"
Else
 Fromdate = DTPickerAccFrom.value
End If


If IsNull(DTPickerAccTo.value) Then
 todate = Date
Else
 todate = DTPickerAccTo.value
End If
    
 
   sql = "SELECT     dbo.projects.Pstate, dbo.projects.id, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.Notes.note_count, dbo.projects.StartDate, dbo.Notes.project_id, "
   sql = sql & "                    dbo.Notes.Emp_ID, dbo.Notes.EmpId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.Notes.NoteDate,"
   sql = sql & "                    dbo.Notes.Note_Value, dbo.projects.project_cost, dbo.RemainingValueProject(dbo.Notes.project_id) AS sumPayment"
   sql = sql & "  FROM         dbo.projects INNER JOIN"
   sql = sql & "                     dbo.Notes ON dbo.projects.id = dbo.Notes.project_id LEFT OUTER JOIN"
   sql = sql & "                     dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID"
      '03062018
       sql = "SELECT     dbo.projects.Pstate, dbo.projects.id, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.Notes.note_count, dbo.projects.StartDate, dbo.Notes.project_id, "
   sql = sql & "                       dbo.Notes.Emp_ID, dbo.Notes.EmpId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.Notes.NoteDate,"
   sql = sql & "                       dbo.Notes.Note_Value, dbo.projects.project_cost, dbo.RemainingValueProject(dbo.Notes.project_id) AS sumPayment, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1,"
   sql = sql & "                       dbo.Notes.Remark"
   sql = sql & "  FROM         dbo.projects INNER JOIN"
   sql = sql & "                       dbo.Notes ON dbo.projects.id = dbo.Notes.project_id LEFT OUTER JOIN"
   sql = sql & "                       dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID"
         sql = sql & "  Where (dbo.Notes.NoteType = 4)"
                 
           If chkProjects.value = vbChecked Then
                  
                If Me.dcprojects.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„‘—Ê⁄...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcprojects.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            StrDes = StrDes & "      " & " «·„‘—Ê⁄ " & dcprojects.text
            End If
    
            If chkCustomers.value = vbChecked Then
        
                If Me.dcCustomers.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· ...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    dcCustomers.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
    StrDes = StrDes & "     " & " «·⁄„Ì· " & dcCustomers.text
            End If
    
 
    
    If chkProjects.value = vbChecked Then
    sql = sql & " and   id=" & val(dcprojects.BoundText)
    End If
    
    If chkCustomers.value = vbChecked Then
    sql = sql & " and   End_user_id='" & val(dcCustomers.BoundText) & "'"
    End If
    
    
        If CHkEmployee.value = vbChecked Then
    sql = sql & " and   dbo.Notes.EmpId=" & val(DcboEmp.BoundText)
    End If
    
 
         If chkbranch.value = vbChecked Then
    sql = sql & " and   projects.branch_no=" & val(DcBranch.BoundText)
    End If
    
    
         If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
            sql = sql + " and     StartDate >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
        End If

        If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
            sql = sql + " and StartDate <=" & SQLDate(Me.DTPickerAccTo, True) & ""
        End If
        
        
        
   If opt(0).value = True Then '  ÃœÌœ
       sql = sql & " and   (Pstate=0   or Pstate is null) "
    ElseIf opt(1).value = True Then '  «›  «ÕÌ
        sql = sql & " and   Pstate=1"
    End If
        
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\projectsrevenue1.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\projectsrevenue1.rpt")
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
          xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
            xReport.ParameterFields(2).AddCurrentValue ""
            
           xReport.ParameterFields(4).AddCurrentValue Format(Fromdate, "DD/MM/YYYY")
        xReport.ParameterFields(5).AddCurrentValue Format(todate, "DD/MM/YYYY")
     xReport.reporttitle = StrDes
     
    FrmReport.txtPath = (App.path & "\reports\construction\projectsrevenue1.rpt")
    FrmReport.CRViewer.viewReport
  cAccountReport.CreateLogo xReport
 

    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
'    SendKeys "{RIGHT}"

   
     End Select

    CuurentLogdata

End Sub
Public Function ShowReportsBills()
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkbranch.value = vbChecked Then
 If val(DcBranch.BoundText) = 0 And DcBranch.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
  Else
    MsgBox "Please Select Branch"
   End If
   DcBranch.SetFocus
  Exit Function
 End If
End If
If chkProjects.value = vbChecked Then
 If val(dcprojects.BoundText) = 0 And dcprojects.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
  Else
    MsgBox "Please Select Project"
   End If
   dcprojects.SetFocus
  Exit Function
 End If
End If
If chkCustomers.value = vbChecked Then
 If val(dcCustomers.BoundText) = 0 And dcCustomers.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
  Else
    MsgBox "Please Select Customer"
   End If
   dcCustomers.SetFocus
  Exit Function
 End If
End If
If chkCustomers1.value = vbChecked Then
 If val(dcCustomers1.BoundText) = 0 And dcCustomers1.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— „ﬁ«Ê· «·»«ÿ‰"
  Else
    MsgBox "Please Select Subcontract"
   End If
   dcCustomers1.SetFocus
  Exit Function
 End If
End If
If CHkEmployee.value = vbChecked Then
 If val(DcboEmp.BoundText) = 0 And DcboEmp.text = "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
  Else
    MsgBox "Please Select Employee"
   End If
   DcboEmp.SetFocus
  Exit Function
 End If
End If

    
    sql = "SELECT     dbo.project_billl.id, dbo.project_billl.bill_date, dbo.project_billl.ManualNO, dbo.project_billl.project_no, dbo.projects.Fullcode, dbo.projects.Project_name, "
    sql = sql & "                   dbo.projects.Project_nameE, dbo.project_billl.discount1ID, dbo.project_billl.discount2ID, dbo.project_billl.discount1value, dbo.project_billl.discount2value,"
    sql = sql & "                  dbo.project_billl.subContractorId, dbo.TblCustemers.CusName AS SubContCusName, dbo.TblCustemers.CusNamee AS SubContCusNameE,"
    sql = sql & "                  dbo.TblCustemers.Fullcode AS SubContFullcode, dbo.projects.id AS ProjectID, dbo.project_billl.Branch_NO, dbo.TblBranchesData.branch_name,"
    sql = sql & "                  dbo.TblBranchesData.branch_namee, dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.projects.EmpId1,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.project_billl.total, dbo.project_billl.Results,"
    sql = sql & "                  dbo.project_billl.discount , dbo.project_billl.advancedPayment , dbo.project_billl.NoteSerial1"
    sql = sql & ", dbo.projects.cost_after_discount,dbo.project_billl.bill_to    FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    sql = sql & "                  dbo.projects ON dbo.TblEmployee.Emp_ID = dbo.projects.EmpId1 LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers TblCustemers_1 ON dbo.projects.End_user_id = TblCustemers_1.CusID RIGHT OUTER JOIN"
    sql = sql & "                  dbo.project_billl LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.project_billl.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.project_billl.subContractorId = dbo.TblCustemers.CusID ON dbo.projects.id = dbo.project_billl.project_no"
sql = sql & " where 1=1"
If val(billto.ListIndex) <> -1 And billto.text <> "" Then
sql = sql & " and dbo.project_billl.bill_to =" & val(Me.billto.ListIndex) & ""
End If

If val(DcBranch.BoundText) <> 0 And DcBranch.text <> "" Then
sql = sql & " and dbo.project_billl.Branch_NO =" & val(Me.DcBranch.BoundText) & ""
End If
If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
sql = sql & " and dbo.projects.id =" & val(Me.dcprojects.BoundText) & ""
End If
If val(dcCustomers.BoundText) <> 0 And dcCustomers.text <> "" Then
sql = sql & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
sql = sql & " and dbo.project_billl.subContractorId =" & val(Me.dcCustomers1.BoundText) & ""
End If
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
sql = sql & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
If Not (IsNull(DTPickerAccFrom.value)) Then
sql = sql & " and dbo.project_billl.bill_date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If
If Not (IsNull(DTPickerAccTo.value)) Then
sql = sql & " and dbo.project_billl.bill_date <=" & SQLDate(DTPickerAccTo.value, True) & ""
End If

                       
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReBillProjectsReport.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReBillProjectsReportE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:

End Function
Private Sub Command1_Click()
 
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\construction\" & "DetailedProject.rpt"
    Else
        StrFileName = App.path & "\Reports\construction\" & "DetailedProject.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then

    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    'If SystemOptions.UserInterface = ArabicInterface Then
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    '
    '    StrReportTitle = "" '& StrAccountName
    ' Else
    '
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    '
    '     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    '    StrReportTitle = ""
    ' End If
    'xReport.ParameterFields(3).AddCurrentValue user_name
    'xReport.ReportTitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

 

End Sub

Function printProjectByItems()
     Dim rs As ADODB.Recordset
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    Dim s As String

s = " SELECT"
s = s & "     dbo.Transaction_Details.project_id1,TblItems.ItemName,TblItems.ItemNamee"
s = s & "    ,(Transaction_Details.Quantity * Transaction_Details.Price) AS TOTAL"
s = s & "    ,Transactions.Transaction_Date"
s = s & "    ,Transactions.Transaction_ID"
s = s & "    ,Transactions.Transaction_Type"
s = s & "    ,Transactions.NoteSerial1"
s = s & "    ,projects.Project_name"
s = s & "    ,Transaction_Details.Pand_ID"
s = s & "    ,Transaction_Details.Oper_ID"
s = s & "    ,pd.des"
s = s & "    ,Transaction_Details.Quantity"
s = s & "    ,Transaction_Details.Price"
s = s & " From dbo.Transaction_Details"
s = s & " INNER JOIN dbo.Transactions"
s = s & "     ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
s = s & " INNER JOIN dbo.TblItems"
s = s & "     ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
s = s & " INNER JOIN dbo.projects"
s = s & "     ON dbo.Transaction_Details.project_ID1 = dbo.projects.id"
s = s & " INNER JOIN dbo.projects_des pd"
s = s & "     ON Transaction_Details.Pand_ID = pd.oprid"
s = s & " WHERE Transactions.Transaction_Type IN ( 990,8,17,18, 991,66)"


If DcBranch.text <> "" Or val(DcBranch.BoundText) <> 0 Then
s = s & "  and  dbo.projects.branch_no =" & val(DcBranch.BoundText)
End If
 
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
s = s & "   And dbo.Projects.EmpID1 = " & val(Me.DcboEmp.BoundText)
End If
 
 'If chkProjects.value = vbChecked Then
 
 If val(dcprojects.BoundText) <> 0 And dcprojects.text <> "" Then
    s = s & " and   projects.id=" & val(dcprojects.BoundText)
End If


 
Dim flg As Integer
 
      If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Reports\REPORTS NEW" & "\projectsItemsInOut.rpt"
        Else
         StrFileName = App.path & "\Reports\REPORTS NEW" & "\projectsItemsInOut.rpt"
    End If
           

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 
    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
       If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
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
   
        StrReportTitle = ""
    End If
 
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , s

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function


Function printProjectByCustomers()
     Dim rs As ADODB.Recordset
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT     TOP 100 PERCENT dbo.TblCustemers.Fullcode AS cuztomerFullcode, dbo.TblCustemers.CusName AS CusName, dbo.TblCustemers.CusNamee AS CusNamee, "
MySQL = MySQL & "                       dbo.projects.*"
MySQL = MySQL & " FROM         dbo.projects INNER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.projects.End_user_id = dbo.TblCustemers.CusID"
                      
 
If dcCustomers.text <> "" Or val(dcCustomers.BoundText) <> 0 Then
MySQL = MySQL & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If DcBranch.text <> "" Or val(DcBranch.BoundText) <> 0 Then
MySQL = MySQL & " and  dbo.projects.branch_no =" & val(DcBranch.BoundText) & ""
End If
 
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
MySQL = MySQL & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
 
 
 
Dim flg As Integer
 
      If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Reports\REPORTS NEW" & "\ProjectsByCustomers.rpt"
        Else
         StrFileName = App.path & "\Reports\REPORTS NEW" & "\ProjectsByCustomers.rpt"
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
        Msg = "No Data"
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
   
        StrReportTitle = ""
    End If
 
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



Function printProjectDet(Optional NoteSerial As String)
     Dim rs As ADODB.Recordset
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT     TOP 100 PERCENT dbo.ProJectMofrdSalar.ID, dbo.ProJectMofrdSalar.EmpID, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee, "
MySQL = MySQL & "                       dbo.ProJectMofrdSalar.ProjID, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.ProJectMofrdSalar.MofrdID, dbo.mofrdat.mofrad_name,"
MySQL = MySQL & "                      dbo.mofrdat.mofrad_namee, dbo.ProJectMofrdSalar.Valuee, dbo.ProJectMofrdSalar.Total, dbo.ProJectMofrdSalar.NoDay, dbo.ProJectMofrdSalar.YearID,"
MySQL = MySQL & "                      dbo.ProJectMofrdSalar.MonthID, TblEmployee_2.SalaryType, TblEmployee_2.DepartmentID, TblEmployee_2.BranchId, TblEmployee_2.ContractID,"
MySQL = MySQL & "                      TblEmployee_2.GroupID, dbo.projects.Salary_account, dbo.ProJectMofrdSalar.pk_id, dbo.ProJectMofrdSalar.TypeSalary, dbo.projects.Fullcode AS ProjectFullcode,"
MySQL = MySQL & "                      dbo.projects.EmpId1, TblEmployee_1.Emp_Name AS SuperEmp_Name, TblEmployee_1.Fullcode AS SuperFullcode,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee AS SuperEmp_NameE, dbo.projects.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                      dbo.projects.End_user_id, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.projects.sub_contractor_id,"
MySQL = MySQL & "                      TblCustemers_1.CusName AS SubCusName, TblCustemers_1.CusNamee AS SubCusNameE, TblCustemers_1.Fullcode AS SubFullcode, dbo.projects.StartDate,dbo.ProJectMofrdSalar.FromDate,dbo.ProJectMofrdSalar.ToDate"
MySQL = MySQL & " FROM         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects ON TblCustemers_1.CusID = dbo.projects.sub_contractor_id RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.ProJectMofrdSalar ON dbo.projects.id = dbo.ProJectMofrdSalar.ProjID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.mofrdat ON dbo.ProJectMofrdSalar.MofrdID = dbo.mofrdat.mofrad_code ON TblEmployee_2.Emp_ID = dbo.ProJectMofrdSalar.EmpID ON"
MySQL = MySQL & "                      dbo.TblCustemers.CusID = dbo.projects.End_user_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.projects.branch_no = dbo.TblBranchesData.branch_id ON TblEmployee_1.Emp_ID = dbo.projects.EmpId1"
MySQL = MySQL & " Where (dbo.ProJectMofrdSalar.ProjID <> 0) And (Not (dbo.Projects.fullcode Is Null))"
If val(dcCustomers1.BoundText) <> 0 And dcCustomers1.text <> "" Then
MySQL = MySQL & " and dbo.projects. sub_contractor_id='" & val(Me.dcCustomers1.BoundText) & "'"
End If
If dcCustomers.text <> "" Or val(dcCustomers.BoundText) <> 0 Then
MySQL = MySQL & " and dbo.projects.End_user_id =" & val(Me.dcCustomers.BoundText) & ""
End If
If DcBranch.text <> "" Or val(DcBranch.BoundText) <> 0 Then
MySQL = MySQL & " and  dbo.projects.branch_no =" & val(DcBranch.BoundText) & ""
End If
If dcprojects.text <> "" Or val(dcprojects.BoundText) <> 0 Then
MySQL = MySQL & "  and dbo.ProJectMofrdSalar.ProjID =" & val(dcprojects.BoundText) & ""
End If
If val(DcboEmp.BoundText) <> 0 And DcboEmp.text <> "" Then
MySQL = MySQL & " and dbo.projects.EmpId1 =" & val(Me.DcboEmp.BoundText) & ""
End If
Dim YearID As Integer
Dim YearID2 As Integer
Dim MonthID As Integer
Dim MonthID2 As Integer
'If Not (IsNull(DTPickerAccFrom.value)) Then
'YearID = year(DTPickerAccFrom.value) - 2006
'MonthID = Month(DTPickerAccFrom.value) - 1
'MySQL = MySQL & " and dbo.ProJectMofrdSalar.YearID >=" & YearID & ""
'MySQL = MySQL & " and dbo.ProJectMofrdSalar.MonthID >=" & MonthID & ""
'End If
'If Not (IsNull(DTPickerAccTo.value)) Then
'YearID2 = year(DTPickerAccTo.value) - 2006
'MonthID2 = Month(DTPickerAccTo.value) - 1
'MySQL = MySQL & " and dbo.ProJectMofrdSalar.YearID <=" & YearID2 & ""
'MySQL = MySQL & " and dbo.ProJectMofrdSalar.MonthID <=" & MonthID2 & ""
'End If
If Not (IsNull(DTPickerAccFrom.value)) And Not (IsNull(DTPickerAccTo.value)) Then
FrmDate1.value = DateAdd("d", -1, Me.DTPickerAccFrom.value)
ToDate1.value = DTPickerAccTo.value
MySQL = MySQL & " and dbo.ProJectMofrdSalar.FromDate between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
MySQL = MySQL & " and dbo.ProJectMofrdSalar.ToDate between " & SQLDate(FrmDate1.value, True) & " and " & SQLDate(ToDate1.value, True) & ""
ElseIf Not (IsNull(DTPickerAccTo.value)) Then
MySQL = MySQL & " and dbo.ProJectMofrdSalar.ToDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
ElseIf Not (IsNull(DTPickerAccFrom.value)) Then
MySQL = MySQL & " and dbo.ProJectMofrdSalar.FromDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
End If
Dim flg As Integer
flg = 0
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â·  —Ìœ  ﬁ—Ì— «Ã„«·Ì"
Else
Msg = "Do you want total reports"
End If
 If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
 flg = 1
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjectsReport.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjectsReportE.rpt"
        End If
   Else
      If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjectsReport2.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepEmpSalaryProjectsReport2E.rpt"
        End If
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
        'GetMsgs 138, vbExclamation
       If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    If flg = 1 Then
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
       xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    End If
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub dcCustomers_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF3 Then
         FrmCustemerSearch.SearchType = 10
            FrmCustemerSearch.show vbModal
           
        End If
End Sub

Private Sub dcCustomers1_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF3 Then
         FrmCustemerSearch.SearchType = 11
            FrmCustemerSearch.show vbModal
           
        End If
End Sub

Private Sub DCProjects_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 7
             FrmProjectSearch.show vbModal
           
        End If
End Sub

Private Sub Form_Load()
    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
    If SystemOptions.UserInterface = EnglishInterface Then
 With billto
.Clear
.AddItem "End User"
.AddItem "Sub-Contractor"
End With
Else
 With billto
.Clear
.AddItem "⁄„Ì· ‰Â«∆Ì"
.AddItem "„ﬁ«Ê· »«ÿ‰"
End With
 End If
    ScreenNameArabic = "  ‹‹ﬁ‹‹«—Ì‹‹‹— «·„‘«—Ì⁄  "
    ScreenNameEnglish = "  Projects  Report "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Dcombos.GetProjects dcprojects
   Dcombos.GetCustomersSuppliers 1, Me.dcCustomers, True
 Dcombos.GetPersons Me.dcCustomers1
   Dcombos.GetBranches DcBranch
  Dcombos.GetPand cmbPands
   Dcombos.GetBoxes DcboBox
   Dcombos.GetBanks Me.DcboBankN
    Dcombos.GetSalesRepData Me.DcboEmp
    OptAccount(10).value = True
    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo
   Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    Me.DTPickerAccFrom = FirstPeriodDateInthisYear
   ' DTPickerAccFrom.value = Date
    DTPickerAccTo.value = Date
    
  '  StrSQL = "  select Emp_ID,Emp_name  from TblEmployee order by Emp_name   "
  '  fill_combo dcEmployee, StrSQL

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

Public Function ShowReports()
    On Error Resume Next

    Dim sql As String
 

    Dim Balance As String
    Dim rs  As ADODB.Recordset
 '*********************
     Dim openingbalacedate As Date
    getOpeningBalancedate , , , , year(todate), openingbalacedate, True
Dim Fromdate As Date
' Dim todate1 As Date
'fromdate = DTPickerAccFrom.value
'todate1 = DTPickerAccTo.value

         If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
        Fromdate = DTPickerAccFrom.value
    End If

    If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
        ToDate1 = DTPickerAccTo.value
    End If
     
        StrSQL = " update projects"
     
     
  ' Dim Fromdate As Date
   StrSQL = StrSQL & " set  expansesE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "', expanses_account, 1 , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
      StrSQL = StrSQL & "    expansesM_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "', Material_account, 1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
       StrSQL = StrSQL & "    expansesS_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "', Salary_account, 1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & "    REVENUE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "', REVENUE_account,  1, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
          StrSQL = StrSQL & " Legal_account_balance=  isnull(dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "', legal, 1 , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0)+ isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', legal,1),0),"
         StrSQL = StrSQL & "    CashingValue= dbo.GetProjectCashing('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate1) & "',  id) "
          
          
          
   
   StrSQL = StrSQL & "    where 1=1"
    
    
        If chkbranch.value = vbChecked Then
    StrSQL = StrSQL & " and   branch_no=" & val(DcBranch.BoundText)
    End If
    
    
    If chkProjects.value = vbChecked Then
    StrSQL = StrSQL & " and   id=" & val(dcprojects.BoundText)
    End If
    
 
    
     Cn.CommandTimeout = 10000
   Cn.Execute StrSQL
   
 '*************************

         
    Dim i As Integer

 
    Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report
Dim FileName As String

    'Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
    'sql = "SELECT * from projects WHERE     (dbo.projects.id = " & val(txt_project_id.text) & ")"
    
    sql = " SELECT     200 AS collected, dbo.project_billl.id, dbo.project_billl.ManualNO, dbo.projects.Project_name, dbo.project_billl.project_no, dbo.project_billl.total, "
sql = sql & " dbo.ProjectBillBuy.[Value], dbo.ProjectBillBuy.TxtNoteSerial, dbo.projects.Project_nameE, dbo.projects.Fullcode, dbo.projects.StartDate, dbo.projects.EndDate,"
sql = sql & " dbo.projects.net, dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.projects.WarrantyNO, dbo.projects.WarrantyValue,"
sql = sql & " dbo.projects.WarrDateStart, dbo.projects.WarrDateEnd, dbo.projects.WarrExtension, dbo.projects.WarrBank, dbo.TblMunicipality.name AS amanah,"
sql = sql & " dbo.TblMunicipality.namee AS amanhe, dbo.TblMunicipalityDet.name AS bldya, dbo.TblMunicipalityDet.namee AS bldyae, dbo.BanksData.BankName,"
sql = sql & " dbo.BanksData.BankNamee, dbo.project_status.name AS project_statusA, dbo.project_status.namee AS project_statusE, dbo.contract_type.name AS contract_typeA,"
sql = sql & " dbo.contract_type.namee AS contract_typeE, dbo.project_billl.bill_date, dbo.ProjectBillBuy.RecordDate, TblCustemers_1.CusName AS subcontractname,"
sql = sql & " TblCustemers_1.CusNamee AS subcontractnamee , ISNULL(dbo.projects.expansese_account_balance, 0) AS expansese_account_balance, "
sql = sql & " ISNULL(dbo.projects.expansesm_account_balance, 0) AS expansesm_account_balance, ISNULL(dbo.projects.expansess_account_balance, 0)"
sql = sql & " AS expansess_account_balance, ISNULL(dbo.projects.REVENUE_account_balance, 0) AS REVENUE_account_balance, ISNULL(dbo.projects.Legal_account_balance,"
sql = sql & " 0) AS [.Legal_account_balance], ISNULL(dbo.projects.CashingValue, 0) AS CashingValue"
sql = sql & "  FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
sql = sql & " dbo.projects ON TblCustemers_1.CusID = dbo.projects.JobeContractorID LEFT OUTER JOIN"
sql = sql & " dbo.contract_type ON dbo.projects.Contract_type = dbo.contract_type.id LEFT OUTER JOIN"
sql = sql & " dbo.project_status ON dbo.projects.Project_status = dbo.project_status.id LEFT OUTER JOIN"
sql = sql & " dbo.BanksData ON dbo.projects.WarrBank = dbo.BanksData.BankID LEFT OUTER JOIN"
sql = sql & " dbo.TblMunicipalityDet ON dbo.projects.Municipalityid = dbo.TblMunicipalityDet.ID RIGHT OUTER JOIN"
sql = sql & " dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
sql = sql & " dbo.project_billl LEFT OUTER JOIN"
sql = sql & " dbo.ProjectBillBuy ON dbo.project_billl.id = dbo.ProjectBillBuy.Bill_id ON dbo.projects.id = dbo.project_billl.project_no LEFT OUTER JOIN"
sql = sql & "  dbo.TblMunicipality ON dbo.projects.Amanhid = dbo.TblMunicipality.ID"
 'cancelled
 
 
 sql = " SELECT     200 AS collected, dbo.projects.Project_name, dbo.ProjectBillBuy.[Value], dbo.ProjectBillBuy.TxtNoteSerial, dbo.projects.Project_nameE, dbo.projects.Fullcode, "
 sql = sql & "                       dbo.projects.StartDate, dbo.projects.EndDate, dbo.projects.net, dbo.projects.End_user_id, TblCustemers_2.CusName, TblCustemers_2.CusNamee,"
  sql = sql & "                      dbo.projects.WarrantyNO, dbo.projects.WarrantyValue, dbo.projects.WarrDateStart, dbo.projects.WarrDateEnd, dbo.projects.WarrExtension, dbo.projects.WarrBank,"
 sql = sql & "                       dbo.TblMunicipality.name AS amanah, dbo.TblMunicipality.namee AS amanhe, dbo.TblMunicipalityDet.name AS bldya, dbo.TblMunicipalityDet.namee AS bldyae,"
 sql = sql & "                       dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.project_status.name AS project_statusA, dbo.project_status.namee AS project_statusE,"
 sql = sql & "                       dbo.contract_type.name AS contract_typeA, dbo.contract_type.namee AS contract_typeE, dbo.ProjectBillBuy.RecordDate,"
 sql = sql & "                       TblCustemers_1.CusName AS subcontractname, TblCustemers_1.CusNamee AS subcontractnamee, ISNULL(dbo.projects.expansese_account_balance, 0)"
 sql = sql & "                       AS expansese_account_balance, ISNULL(dbo.projects.expansesm_account_balance, 0) AS expansesm_account_balance,"
 sql = sql & "                       ISNULL(dbo.projects.expansess_account_balance, 0) AS expansess_account_balance, ISNULL(dbo.projects.REVENUE_account_balance, 0)"
 sql = sql & "                       AS REVENUE_account_balance, ISNULL(dbo.projects.Legal_account_balance, 0) AS [.Legal_account_balance], ISNULL(dbo.projects.CashingValue, 0)"
sql = sql & "                        AS CashingValue, ISNULL(dbo.project_billl.discount, 0) AS discount, ISNULL(dbo.project_billl.Results, 0) AS Results, dbo.project_billl.project_no,"
sql = sql & "                        dbo.project_billl.ManualNO , dbo.project_billl.bill_date, dbo.project_billl.total"
sql = sql & ", isnull(dbo.project_billl.advancedPayment,0) , dbo.project_billl.id  FROM         dbo.TblMunicipalityDet RIGHT OUTER JOIN"
sql = sql & "                        dbo.BanksData RIGHT OUTER JOIN"
sql = sql & "                        dbo.ProjectBillBuy RIGHT OUTER JOIN"
sql = sql & "                        dbo.project_billl ON dbo.ProjectBillBuy.Bill_id = dbo.project_billl.id LEFT OUTER JOIN"
sql = sql & "                        dbo.projects ON dbo.project_billl.project_no = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                        dbo.TblCustemers TblCustemers_1 ON dbo.projects.JobeContractorID = TblCustemers_1.CusID LEFT OUTER JOIN"
sql = sql & "                        dbo.contract_type ON dbo.projects.Contract_type = dbo.contract_type.id LEFT OUTER JOIN"
sql = sql & "                        dbo.project_status ON dbo.projects.Project_status = dbo.project_status.id ON dbo.BanksData.BankID = dbo.projects.WarrBank ON"
sql = sql & "                        dbo.TblMunicipalityDet.ID = dbo.projects.Municipalityid RIGHT OUTER JOIN"
sql = sql & "                        dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
sql = sql & "                        dbo.TblMunicipality ON dbo.projects.Amanhid = dbo.TblMunicipality.ID"
 sql = sql & "  WHERE      dbo.projects.id = " & val(Me.dcprojects.BoundText) & " "
    
        If Me.DTPickerAccFrom <> Empty Or Me.DTPickerAccFrom <> Null Then
        sql = sql + " and   bill_date >=" & SQLDate(Me.DTPickerAccFrom, True) & ""
    End If

    If Me.DTPickerAccTo <> Empty Or Me.DTPickerAccTo <> Null Then
        sql = sql + " and bill_date <=" & SQLDate(DTPickerAccTo, True) & " "
    End If
    
    
   
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       If rs.RecordCount = 0 Then
       
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "·« ÌÊÃœ »Ì«‰« "
      Else
      MsgBox "No Data to print"
      End If
       Screen.MousePointer = vbDefault
       Exit Function
       End If
    If SystemOptions.UserInterface = ArabicInterface Then
    FileName = App.path & "\reports\REPORTS NEW\ProjectRPT.rpt "
    
        Set xReport = xApp.OpenReport(FileName)
    Else
    FileName = App.path & "\reports\REPORTS NEW\ProjectRPTe.rpt"
        Set xReport = xApp.OpenReport(FileName)
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = FileName
    
         If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(3).AddCurrentValue user_name
                If IsNull(rs("StartDate").value) Or IsNull(rs("EndDate").value) Then
                xReport.ParameterFields(12).AddCurrentValue ""
                xReport.ParameterFields(13).AddCurrentValue ""
                
                Else
                xReport.ParameterFields(12).AddCurrentValue CStr(DateDiff("M", rs("StartDate").value, rs("EndDate").value))
                xReport.ParameterFields(13).AddCurrentValue CStr(DateDiff("M", Date, rs("EndDate").value))
                
                End If
                
       
     
        StrReportTitle = "" '& StrAccountName
 
    Else
 
     '
 
    End If
    
    FrmReport.CRViewer.viewReport
 
 '   xReport.reporttitle = cCompanyInfo.ArabCompanyName

    
    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
    Sendkeys "{RIGHT}"

End Function
Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub OptAccount_Click(Index As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
If OptAccount(14).value = True Then
Dcombos.GetEmployees Me.DcboEmp
Else
Dcombos.GetSalesRepData Me.DcboEmp
End If
End Sub
