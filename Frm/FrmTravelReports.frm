VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form frmTravelRports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·—Õ·« "
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12210
   HelpContextID   =   470
   Icon            =   "FrmTravelReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   12210
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
      Height          =   6930
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12210
      _cx             =   21537
      _cy             =   12224
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
      _GridInfo       =   $"FrmTravelReports.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EleMain 
         Height          =   6375
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   12150
         _cx             =   21431
         _cy             =   11245
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
            Height          =   6195
            Left            =   90
            TabIndex        =   2
            Top             =   90
            Width           =   11970
            _cx             =   21114
            _cy             =   10927
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
            Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·—Õ·« | ﬁ«—Ì— 2| ﬁ«—Ì— 3"
            Align           =   0
            CurrTab         =   2
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
               Height          =   5820
               Index           =   0
               Left            =   -12825
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   45
               Width           =   11880
               _cx             =   20955
               _cy             =   10266
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
                  Height          =   5175
                  Index           =   2
                  Left            =   90
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   11430
                  _cx             =   20161
                  _cy             =   9128
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
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰… «Ã„«·Ì« "
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   0
                     Left            =   8880
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   3540
                     Width           =   5820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰… ·”«∆ﬁ „⁄Ì‰"
                     Height          =   195
                     Index           =   1
                     Left            =   11160
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   1680
                     Width           =   3420
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰… ·”Ì«—… „⁄Ì‰…"
                     Height          =   195
                     Index           =   2
                     Left            =   11520
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   1920
                     Width           =   3420
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰… ·⁄„Ì· „⁄Ì‰"
                     Height          =   195
                     Index           =   3
                     Left            =   11400
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   2145
                     Width           =   3300
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÿ·»Ì«  «· Ì ·„  ”·„ Õ Ï «·«‰"
                     Height          =   195
                     Index           =   4
                     Left            =   14040
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.TextBox Txt_order_no 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   14400
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   2280
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   13560
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   3240
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   5
                     Left            =   11160
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   960
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   6
                     Left            =   11280
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   1200
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰…  Õ·Ì·Ì"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   7
                     Left            =   8160
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   720
                     Value           =   -1  'True
                     Width           =   2940
                  End
                  Begin VB.Frame Frame2 
                     Caption         =   "»Ì«‰«  «·—Õ·…"
                     Height          =   2055
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   7
                     Top             =   1080
                     Width           =   11175
                     Begin VB.TextBox TxtLocation 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   8
                        Top             =   960
                        Width           =   8745
                     End
                     Begin MSDataListLib.DataCombo DcCityFromId 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   9
                        Top             =   240
                        Width           =   8715
                        _ExtentX        =   15372
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCEmp 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   10
                        Top             =   1680
                        Width           =   8715
                        _ExtentX        =   15372
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcCityToId 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   11
                        Top             =   600
                        Width           =   8715
                        _ExtentX        =   15372
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCCar 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   12
                        Top             =   1320
                        Width           =   8715
                        _ExtentX        =   15372
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·Ï"
                        Height          =   285
                        Index           =   42
                        Left            =   9600
                        RightToLeft     =   -1  'True
                        TabIndex        =   17
                        Top             =   600
                        Width           =   1215
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "ÃÂ… «·Ê’Ê·"
                        Height          =   285
                        Index           =   41
                        Left            =   9600
                        RightToLeft     =   -1  'True
                        TabIndex        =   16
                        Top             =   960
                        Width           =   1215
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·”«∆ﬁ"
                        Height          =   285
                        Index           =   28
                        Left            =   9600
                        RightToLeft     =   -1  'True
                        TabIndex        =   15
                        Top             =   1680
                        Width           =   1215
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·„⁄œÂ/«·”Ì«—…"
                        Height          =   285
                        Index           =   26
                        Left            =   9600
                        RightToLeft     =   -1  'True
                        TabIndex        =   14
                        Top             =   1320
                        Width           =   1215
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·—Õ·… „‰ "
                        Height          =   285
                        Index           =   25
                        Left            =   9600
                        RightToLeft     =   -1  'True
                        TabIndex        =   13
                        Top             =   240
                        Width           =   1215
                     End
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «·—Õ·«  Œ·«· › —… „⁄Ì‰… «Ã„«·Ì« "
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   8
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   720
                     Width           =   3660
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ——Ì— «—»«Õ «·‘««Õ‰« "
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   9
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   5
                     Top             =   720
                     Width           =   3660
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1065
                     Index           =   1
                     Left            =   210
                     TabIndex        =   28
                     TabStop         =   0   'False
                     Top             =   3480
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
                        TabIndex        =   29
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
                        Format          =   112525315
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
                        TabIndex        =   30
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
                        Format          =   112525315
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
                        TabIndex        =   32
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
                        TabIndex        =   31
                        Top             =   285
                        Width           =   555
                     End
                  End
                  Begin ImpulseButton.ISButton CmdAccount 
                     Height          =   405
                     Left            =   120
                     TabIndex        =   33
                     Top             =   4680
                     Width           =   1305
                     _ExtentX        =   2302
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
                     ButtonImage     =   "FrmTravelReports.frx":040F
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
                     Height          =   405
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   75
                     Width           =   11070
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·—Õ·…"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   17
                     Left            =   11520
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   2400
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   0
                     Left            =   13440
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   3000
                     Visible         =   0   'False
                     Width           =   1335
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   5820
               Left            =   -12525
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   45
               Width           =   11880
               _cx             =   20955
               _cy             =   10266
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
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   120
                  TabIndex        =   76
                  Top             =   1920
                  Width           =   4815
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   0
                     Left            =   1680
                     TabIndex        =   77
                     Top             =   0
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   78
                     Top             =   0
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "„„·Êﬂ… ··€Ì—"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   2
                     Left            =   3240
                     TabIndex        =   79
                     Top             =   0
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·ﬂ·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.TextBox Text2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3975
                  TabIndex        =   75
                  Top             =   2640
                  Width           =   930
               End
               Begin VB.TextBox TxtItemCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3945
                  TabIndex        =   74
                  Top             =   1215
                  Width           =   930
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3945
                  TabIndex        =   73
                  Top             =   480
                  Width           =   930
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  Height          =   495
                  Left            =   120
                  TabIndex        =   69
                  Top             =   3000
                  Width           =   4695
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   3
                     Left            =   1680
                     TabIndex        =   70
                     Top             =   120
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   " „ «’œ«— ›« Ê—… "
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   71
                     Top             =   120
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·„ Ì „ «·«’œ«—"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChCarType 
                     Height          =   255
                     Index           =   5
                     Left            =   3240
                     TabIndex        =   72
                     Top             =   120
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·ﬂ·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   495
                  Left            =   0
                  TabIndex        =   59
                  Top             =   4320
                  Width           =   11415
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   0
                     Left            =   9840
                     TabIndex        =   60
                     Top             =   120
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·„—ﬂ»… «·„„·Êﬂ…"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   1
                     Left            =   8520
                     TabIndex        =   61
                     Top             =   120
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "„„·Êﬂ… ··€Ì—"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   2
                     Left            =   7800
                     TabIndex        =   62
                     Top             =   120
                     Width           =   735
                     _Version        =   786432
                     _ExtentX        =   1296
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "··„«·ﬂ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   3
                     Left            =   7080
                     TabIndex        =   63
                     Top             =   120
                     Width           =   735
                     _Version        =   786432
                     _ExtentX        =   1296
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "··”«∆ﬁ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   4
                     Left            =   6360
                     TabIndex        =   64
                     Top             =   120
                     Width           =   735
                     _Version        =   786432
                     _ExtentX        =   1296
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "··⁄„Ì·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   5
                     Left            =   5280
                     TabIndex        =   65
                     Top             =   120
                     Width           =   975
                     _Version        =   786432
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·‰Ê⁄ «·‰ﬁ·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   6
                     Left            =   3360
                     TabIndex        =   66
                     Top             =   120
                     Width           =   975
                     _Version        =   786432
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "··’‰›"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   7
                     Left            =   2640
                     TabIndex        =   67
                     Top             =   120
                     Width           =   735
                     _Version        =   786432
                     _ExtentX        =   1296
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·· «—ÌŒ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   8
                     Left            =   4320
                     TabIndex        =   68
                     Top             =   120
                     Width           =   855
                     _Version        =   786432
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "·”›Ì‰…"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   9
                     Left            =   1440
                     TabIndex        =   130
                     Top             =   120
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ﬂ„Ì…  Õ„Ì·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdSort 
                     Height          =   255
                     Index           =   10
                     Left            =   120
                     TabIndex        =   131
                     Top             =   120
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ﬂ„Ì…  ›—Ì€"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.TextBox TxtCard1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   58
                  Top             =   3120
                  Width           =   1290
               End
               Begin VB.TextBox TxtCard2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   57
                  Top             =   3120
                  Width           =   1290
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Height          =   375
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   3870
                  Width           =   3495
                  Begin XtremeSuiteControls.RadioButton RdNet 
                     Height          =   255
                     Index           =   0
                     Left            =   3000
                     TabIndex        =   52
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   ">"
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdNet 
                     Height          =   255
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   53
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "<"
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdNet 
                     Height          =   255
                     Index           =   2
                     Left            =   1800
                     TabIndex        =   54
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdNet 
                     Height          =   255
                     Index           =   3
                     Left            =   1080
                     TabIndex        =   55
                     Top             =   0
                     Width           =   615
                     _Version        =   786432
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   ">="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdNet 
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   56
                     Top             =   0
                     Width           =   855
                     _Version        =   786432
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "<="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Height          =   375
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   3510
                  Width           =   3495
                  Begin XtremeSuiteControls.RadioButton RdTotal 
                     Height          =   255
                     Index           =   0
                     Left            =   3000
                     TabIndex        =   46
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   ">"
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdTotal 
                     Height          =   255
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   47
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "<"
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdTotal 
                     Height          =   255
                     Index           =   2
                     Left            =   1800
                     TabIndex        =   48
                     Top             =   0
                     Width           =   495
                     _Version        =   786432
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdTotal 
                     Height          =   255
                     Index           =   3
                     Left            =   1080
                     TabIndex        =   49
                     Top             =   0
                     Width           =   615
                     _Version        =   786432
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   ">="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdTotal 
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   50
                     Top             =   0
                     Width           =   855
                     _Version        =   786432
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "<="
                     BackColor       =   14871017
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                  End
               End
               Begin VB.TextBox TxtQtyUpload 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   44
                  Top             =   3480
                  Width           =   1290
               End
               Begin VB.TextBox TxtQtyDownLoad 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   43
                  Top             =   3840
                  Width           =   1290
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   495
                  Left            =   1440
                  TabIndex        =   38
                  Top             =   4800
                  Width           =   9375
                  Begin XtremeSuiteControls.RadioButton RdPrint 
                     Height          =   255
                     Index           =   0
                     Left            =   7920
                     TabIndex        =   39
                     Top             =   120
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "›Ê« Ì— «·⁄„·«¡"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdPrint 
                     Height          =   255
                     Index           =   1
                     Left            =   6720
                     TabIndex        =   40
                     Top             =   120
                     Width           =   975
                     _Version        =   786432
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   " Õ·Ì·Ì"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdPrint 
                     Height          =   255
                     Index           =   2
                     Left            =   5400
                     TabIndex        =   41
                     Top             =   120
                     Width           =   975
                     _Version        =   786432
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·»«Œ—« "
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RdPrint 
                     Height          =   255
                     Index           =   3
                     Left            =   4440
                     TabIndex        =   42
                     Top             =   120
                     Width           =   975
                     _Version        =   786432
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·—Õ·« "
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Index           =   0
                  Left            =   6000
                  TabIndex        =   80
                  Top             =   480
                  Width           =   4230
                  _ExtentX        =   7461
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
               Begin MSDataListLib.DataCombo DcCityFromId2 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   81
                  Top             =   840
                  Width           =   4230
                  _ExtentX        =   7461
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
               Begin MSDataListLib.DataCombo DcCityToId2 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   82
                  Top             =   1200
                  Width           =   4230
                  _ExtentX        =   7461
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
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   83
                  Top             =   480
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  BoundColumn     =   ""
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbTypeTransport 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   84
                  Top             =   840
                  Width           =   4755
                  _ExtentX        =   8387
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboItems 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1215
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbShip 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   86
                  Top             =   1560
                  Width           =   4230
                  _ExtentX        =   7461
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo VehicleType 
                  Height          =   315
                  Index           =   0
                  Left            =   6000
                  TabIndex        =   87
                  Top             =   1920
                  Width           =   4230
                  _ExtentX        =   7461
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbHarbor 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1560
                  Width           =   4755
                  _ExtentX        =   8387
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCCar2 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   89
                  Top             =   2280
                  Width           =   4230
                  _ExtentX        =   7461
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCCar3 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   90
                  Top             =   2640
                  Width           =   4230
                  _ExtentX        =   7461
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   705
                  Index           =   3
                  Left            =   120
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   3480
                  Width           =   4635
                  _cx             =   8176
                  _cy             =   1244
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
                  Begin MSComCtl2.DTPicker ToDate 
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
                     TabIndex        =   92
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
                     Format          =   112525315
                     CurrentDate     =   37357
                  End
                  Begin MSComCtl2.DTPicker FrmDate 
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
                     Left            =   2490
                     TabIndex        =   93
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
                     Format          =   112525315
                     CurrentDate     =   37357
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰"
                     Height          =   285
                     Index           =   9
                     Left            =   3990
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   285
                     Width           =   555
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈·Ï"
                     Height          =   285
                     Index           =   10
                     Left            =   1590
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   240
                     Width           =   555
                  End
               End
               Begin ImpulseButton.ISButton ISButton1 
                  Height          =   405
                  Left            =   0
                  TabIndex        =   96
                  Top             =   4920
                  Width           =   1305
                  _ExtentX        =   2302
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
                  ButtonImage     =   "FrmTravelReports.frx":07A9
                  ColorButton     =   14871017
                  ColorHoverText  =   16777215
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16777215
               End
               Begin MSDataListLib.DataCombo DcbEmployee 
                  Height          =   315
                  Index           =   0
                  Left            =   150
                  TabIndex        =   132
                  Top             =   2280
                  Width           =   4755
                  _ExtentX        =   8387
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DBCboClientName2 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   133
                  Top             =   2640
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  BoundColumn     =   ""
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   134
                  Top             =   4920
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  BackStyle       =   0
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„«·ﬂ"
                  Height          =   285
                  Index           =   3
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   2640
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„„·ÊﬂÂ ··€Ì—"
                  Height          =   285
                  Index           =   57
                  Left            =   10095
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   2640
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”«∆ﬁ"
                  Height          =   285
                  Index           =   1
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   2280
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
                  Height          =   285
                  Index           =   5
                  Left            =   10095
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   2280
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Ì‰«¡"
                  Height          =   285
                  Index           =   55
                  Left            =   4635
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   1560
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·„—ﬂ»…"
                  Height          =   285
                  Index           =   49
                  Left            =   10095
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   1920
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·„—ﬂ»…"
                  Height          =   285
                  Index           =   81
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   1920
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”›Ì‰…"
                  Height          =   210
                  Index           =   6
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1560
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   210
                  Index           =   7
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   1200
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   210
                  Index           =   8
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   840
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·’‰›"
                  Height          =   315
                  Index           =   78
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1200
                  Width           =   1530
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   510
                  Left            =   5445
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   1200
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«‰Ê«⁄ «·‰ﬁ·"
                  Height          =   315
                  Index           =   72
                  Left            =   4275
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   840
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„Ì·"
                  Height          =   315
                  Index           =   64
                  Left            =   4275
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   480
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   210
                  Index           =   13
                  Left            =   10350
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   480
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
                  Height          =   210
                  Index           =   11
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   3120
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " — Ì» ÿ»ﬁ«"
                  ForeColor       =   &H00800000&
                  Height          =   210
                  Index           =   12
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   4080
                  Width           =   990
               End
               Begin VB.Label Label1 
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
                  Height          =   405
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   0
                  Width           =   11910
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ ﬂ—  «· Õ„Ì·"
                  Height          =   285
                  Index           =   14
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   3120
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ ﬂ—  «· ›—Ì€"
                  Height          =   285
                  Index           =   15
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   3120
                  Width           =   1215
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ„Ì… «· Õ„Ì·"
                  Height          =   315
                  Index           =   7
                  Left            =   10215
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   3480
                  Width           =   1065
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ„Ì… «· ›—Ì€"
                  Height          =   315
                  Index           =   0
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   3840
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬁ—Ì—"
                  ForeColor       =   &H00800000&
                  Height          =   210
                  Index           =   16
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   4920
                  Width           =   990
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5820
               Index           =   4
               Left            =   45
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   45
               Width           =   11880
               _cx             =   20955
               _cy             =   10266
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
               Begin VB.Frame Frame10 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H80000008&
                  Height          =   765
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   3930
                  Width           =   5655
                  Begin VB.ComboBox cmbToYear 
                     Height          =   315
                     Left            =   60
                     RightToLeft     =   -1  'True
                     TabIndex        =   180
                     Top             =   210
                     Width           =   1005
                  End
                  Begin VB.ComboBox cmbToMonthName 
                     Height          =   315
                     Left            =   1110
                     RightToLeft     =   -1  'True
                     TabIndex        =   179
                     Top             =   210
                     Width           =   1005
                  End
                  Begin VB.ComboBox cmbFromYear 
                     Height          =   315
                     Left            =   2910
                     RightToLeft     =   -1  'True
                     TabIndex        =   178
                     Top             =   240
                     Width           =   1005
                  End
                  Begin VB.ComboBox cmbFromMonthName 
                     Height          =   315
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   177
                     Top             =   240
                     Width           =   1005
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï ‘Â—"
                     Height          =   435
                     Index           =   36
                     Left            =   2250
                     RightToLeft     =   -1  'True
                     TabIndex        =   176
                     Top             =   240
                     Width           =   585
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰ ‘Â—"
                     Height          =   375
                     Index           =   35
                     Left            =   5130
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   240
                     Width           =   405
                  End
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   2280
                  Width           =   1215
               End
               Begin VB.ComboBox cmbTypeRep 
                  Height          =   315
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   3495
                  Width           =   2325
               End
               Begin VB.ComboBox carStatus 
                  Height          =   315
                  ItemData        =   "FrmTravelReports.frx":0B43
                  Left            =   120
                  List            =   "FrmTravelReports.frx":0B45
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   3150
                  Width           =   4665
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.ComboBox orderStatus 
                  Height          =   315
                  ItemData        =   "FrmTravelReports.frx":0B47
                  Left            =   120
                  List            =   "FrmTravelReports.frx":0B49
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   2745
                  Width           =   4665
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
                  Height          =   1995
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   3720
                  Width           =   5535
                  Begin XtremeSuiteControls.RadioButton RepUploadOrderChk 
                     Height          =   255
                     Left            =   2760
                     TabIndex        =   155
                     Top             =   240
                     Width           =   2220
                     _Version        =   786432
                     _ExtentX        =   3916
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   " ﬁ—Ì— √Ê«„— «· Õ„Ì·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton RepCarsDelayChk 
                     Height          =   225
                     Left            =   2850
                     TabIndex        =   156
                     Top             =   960
                     Width           =   2130
                     _Version        =   786432
                     _ExtentX        =   3757
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   " ﬁ—Ì—  √Œ— «·„⁄œ« /«·”Ì«—« "
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   525
                     Left            =   120
                     TabIndex        =   157
                     Top             =   720
                     Width           =   1425
                     _ExtentX        =   2514
                     _ExtentY        =   926
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
                     ButtonImage     =   "FrmTravelReports.frx":0B4B
                     ColorButton     =   14871017
                     ColorHoverText  =   16777215
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16777215
                  End
                  Begin XtremeSuiteControls.RadioButton notInvoicing 
                     Height          =   255
                     Left            =   1800
                     TabIndex        =   167
                     Top             =   480
                     Width           =   3180
                     _Version        =   786432
                     _ExtentX        =   5609
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   " ﬁ—Ì— √Ê«„— «· Õ„Ì· «· Ì ·„  ÕÊ· ·—Õ·Â"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Invoicing 
                     Height          =   255
                     Left            =   1800
                     TabIndex        =   168
                     Top             =   720
                     Width           =   3180
                     _Version        =   786432
                     _ExtentX        =   5609
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   " ﬁ—Ì— √Ê«„— «· Õ„Ì· «· Ì    „  ÕÊÌ·Â« ·—Õ·Â"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                     Value           =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton optRepMonth 
                     Height          =   225
                     Left            =   2850
                     TabIndex        =   172
                     Top             =   1200
                     Width           =   2130
                     _Version        =   786432
                     _ExtentX        =   3757
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "Õ—ﬂ… «·—œÊœ «·‘Â—Ì…"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton optOrderUploadRec 
                     Height          =   225
                     Left            =   2850
                     TabIndex        =   173
                     Top             =   1470
                     Width           =   2130
                     _Version        =   786432
                     _ExtentX        =   3757
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   " ﬁ—Ì—  ”·Ì„ «·›Ê« Ì—"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton KPI 
                     Height          =   225
                     Left            =   2880
                     TabIndex        =   181
                     Top             =   1680
                     Width           =   2130
                     _Version        =   786432
                     _ExtentX        =   3757
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "„ Ê”ÿ «·—œÊœ ··⁄„·«¡"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.TextBox TxtSearchCode1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   9255
                  TabIndex        =   135
                  Top             =   2520
                  Width           =   1065
               End
               Begin VB.TextBox txtId 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   5805
                  TabIndex        =   126
                  Top             =   3390
                  Width           =   4530
               End
               Begin VB.Frame Frame8 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H80000008&
                  Height          =   945
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   4590
                  Width           =   5655
                  Begin MSComCtl2.DTPicker ToDate2 
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
                     Left            =   600
                     TabIndex        =   122
                     ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                     Top             =   360
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CheckBox        =   -1  'True
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   112525315
                     CurrentDate     =   37357
                  End
                  Begin MSComCtl2.DTPicker FromDate2 
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
                     Left            =   3030
                     TabIndex        =   123
                     ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                     Top             =   360
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   609
                     _Version        =   393216
                     CalendarBackColor=   -2147483624
                     CalendarTitleBackColor=   10383715
                     CheckBox        =   -1  'True
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   112525315
                     CurrentDate     =   37357
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰"
                     Height          =   285
                     Index           =   22
                     Left            =   4530
                     RightToLeft     =   -1  'True
                     TabIndex        =   125
                     Top             =   390
                     Width           =   345
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï"
                     Height          =   285
                     Index           =   23
                     Left            =   2100
                     RightToLeft     =   -1  'True
                     TabIndex        =   124
                     Top             =   390
                     Width           =   345
                  End
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Index           =   1
                  Left            =   5880
                  TabIndex        =   127
                  Top             =   270
                  Width           =   4470
                  _ExtentX        =   7885
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
               Begin XtremeSuiteControls.RadioButton ChCarType 
                  Height          =   225
                  Index           =   6
                  Left            =   10305
                  TabIndex        =   136
                  Top             =   2070
                  Width           =   750
                  _Version        =   786432
                  _ExtentX        =   1323
                  _ExtentY        =   397
                  _StockProps     =   79
                  Caption         =   "„„·Êﬂ…"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
                  Value           =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton ChCarType 
                  Height          =   210
                  Index           =   7
                  Left            =   10290
                  TabIndex        =   137
                  Top             =   2565
                  Width           =   765
                  _Version        =   786432
                  _ExtentX        =   1349
                  _ExtentY        =   370
                  _StockProps     =   79
                  Caption         =   "«Œ—Ï"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCar 
                  Height          =   315
                  Left            =   5805
                  TabIndex        =   138
                  Top             =   2070
                  Width           =   4515
                  _ExtentX        =   7964
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbSupplem 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   139
                  Top             =   1860
                  Width           =   4665
                  _ExtentX        =   8229
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DBCboClientName1 
                  Height          =   315
                  Left            =   5805
                  TabIndex        =   140
                  Top             =   2520
                  Width           =   3465
                  _ExtentX        =   6112
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  BoundColumn     =   ""
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic12 
                  Height          =   1200
                  Left            =   90
                  TabIndex        =   145
                  TabStop         =   0   'False
                  Top             =   645
                  Width           =   11745
                  _cx             =   20717
                  _cy             =   2117
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
                  Begin VB.TextBox TxtLeaderName 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   5775
                     RightToLeft     =   -1  'True
                     TabIndex        =   147
                     Top             =   720
                     Width           =   4515
                  End
                  Begin VB.TextBox Text6 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   9210
                     RightToLeft     =   -1  'True
                     TabIndex        =   146
                     Top             =   360
                     Width           =   1080
                  End
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   255
                     Index           =   0
                     Left            =   10320
                     TabIndex        =   148
                     Top             =   405
                     Width           =   1020
                     _Version        =   786432
                     _ExtentX        =   1799
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "œ«Œ·Ì"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                     Value           =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   225
                     Index           =   1
                     Left            =   10395
                     TabIndex        =   149
                     Top             =   750
                     Width           =   930
                     _Version        =   786432
                     _ExtentX        =   1640
                     _ExtentY        =   397
                     _StockProps     =   79
                     Caption         =   "Œ«—ÃÌ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcEmployee2 
                     Height          =   315
                     Left            =   5775
                     TabIndex        =   150
                     Top             =   360
                     Width           =   3435
                     _ExtentX        =   6059
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·”«∆ﬁ"
                     Height          =   240
                     Index           =   31
                     Left            =   10110
                     RightToLeft     =   -1  'True
                     TabIndex        =   152
                     Top             =   120
                     Width           =   1440
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   480
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   151
                     Top             =   240
                     Width           =   1020
                  End
               End
               Begin MSDataListLib.DataCombo DcEmpSuper 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   158
                  Top             =   240
                  Width           =   3630
                  _ExtentX        =   6403
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
               Begin MSDataListLib.DataCombo DcbCar2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   162
                  Top             =   2310
                  Width           =   3465
                  _ExtentX        =   6112
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbSupplem2 
                  Height          =   315
                  Left            =   5805
                  TabIndex        =   163
                  Top             =   2955
                  Width           =   4515
                  _ExtentX        =   7964
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·—œ"
                  Height          =   240
                  Index           =   34
                  Left            =   4020
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Top             =   3495
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·„⁄œÂ/«·”Ì«—…"
                  Height          =   240
                  Index           =   33
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   3150
                  Width           =   930
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·√„—"
                  Height          =   240
                  Index           =   30
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   2775
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„—«ﬁ»"
                  Height          =   210
                  Index           =   27
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   285
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  Height          =   270
                  Index           =   32
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   720
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„—ﬂ»…"
                  Height          =   225
                  Index           =   29
                  Left            =   4785
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   2355
                  Width           =   750
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„·Õﬁ"
                  Height          =   240
                  Index           =   24
                  Left            =   4785
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   1890
                  Width           =   750
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„·Õﬁ"
                  Height          =   240
                  Index           =   20
                  Left            =   10410
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   2955
                  Width           =   570
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„—ﬂ»…"
                  Height          =   240
                  Index           =   19
                  Left            =   11190
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   2100
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   210
                  Index           =   18
                  Left            =   10590
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   270
                  Width           =   870
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ Õ—ﬂ… «· Õ„Ì·"
                  Height          =   285
                  Index           =   21
                  Left            =   10425
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   3390
                  Width           =   1215
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmTravelRports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String


Private Sub DcEmployee2_Change()
DcEmployee2_Click (0)
End Sub

Private Sub DcEmployee2_Click(Area As Integer)
Dim Nationality As String
Dim NumEkama As String
    If val(DcEmployee2.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcEmployee2.BoundText, EmpCode
      Text6.Text = EmpCode
       
        get_employee_information val(Me.DcEmployee2.BoundText), , , , , , , , , Nationality, , , , , NumEkama
'        TxtNationality.Text = Nationality
'        TxtIDNo.Text = NumEkama
       ' RetriveCarsInfo val(DcEmployee2.BoundText), 1
End Sub
Private Sub ChDrievType_Click(Index As Integer)
If ChDrievType(0).value = True Then
Text6.Enabled = True
DcEmployee2.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DcEmployee2.Enabled = False
TxtLeaderName.Enabled = True
DcEmployee2.BoundText = 0
Text6.Text = ""
End If
End Sub

Private Sub Invoicing_Click()
optRepMonth_Click
End Sub

Private Sub notInvoicing_Click()
optRepMonth_Click
End Sub

Private Sub optOrderUploadRec_Click()
    optRepMonth_Click
End Sub

Private Sub optRepMonth_Click()
    If optRepMonth Then
        Frame10.Enabled = True
        Frame8.Enabled = False
    Else
        Frame10.Enabled = False
        Frame8.Enabled = True
    End If
End Sub

Private Sub RepCarsDelayChk_Click()
optRepMonth_Click
End Sub

Private Sub RepUploadOrderChk_Click()
optRepMonth_Click
End Sub

Private Sub Text4_Change()
    Dim Dcombos As New ClsDataCombos
    
    Dcombos.GetQuicSearch DcbCar2, Text4, "FixedAssets", "Id", , , , " And ISEQUP = 1 "
    
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcEmployee2.BoundText = EmpID
    End If
End Sub


Private Sub ChCarType_Click(Index As Integer)
If ChCarType(6).value = True Then
DcbSupplem.Enabled = True
DcbCar.Enabled = True
TxtSearchCode1.Enabled = False
TxtSearchCode1.Text = ""
DBCboClientName.Enabled = False
DcbCar2.Enabled = False
DcbSupplem2.Enabled = False

ElseIf ChCarType(7).value = True Then
DcbSupplem2.Enabled = True
DcbSupplem2.BoundText = 0
DcbCar2.BoundText = 0
DcbCar2.Enabled = True
DBCboClientName.BoundText = 0
DBCboClientName.Enabled = True
TxtSearchCode1.Enabled = True
DcbCar.BoundText = 0
DcbCar.Enabled = False
DcbSupplem.BoundText = 0
DcbSupplem.Enabled = False
End If

End Sub
Private Sub DBCboClientName1_Click(Area As Integer)
'If Me.TxtModFlg.Text <> "R" Then
'If val(DBCboClientName1.BoundText) <> 0 Then
'If ChCarType(0).value = True Then
'RetriveClinCounr val(val(DcbCar.BoundText)), 0
'Else
'RetriveClinCounr val(val(DcbCar2.BoundText)), 1
'End If
'End If
'End If
End Sub

Private Sub DBCboClientName1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 101
        FrmCustemerSearch.show vbModal

    End If
End Sub

Private Sub DcbCar_Change()
DcbCar_Click (0)
End Sub

Private Sub DcbCar_Click(Area As Integer)
 Dim Dcombos As New ClsDataCombos
Dcombos.GetPartCar DcbSupplem, val(DcbCar.BoundText)
'RetriveClinCounr val(val(DcbCar.BoundText)), 0
RetriveCarsInfo val(DcbCar.BoundText), 0
End Sub

Private Sub DcbCar2_Change()
DcbCar2_Click (0)
End Sub
Private Sub DcbCar2_Click(Area As Integer)
 Dim Dcombos As New ClsDataCombos
Dcombos.GetBartCarByVonder DcbSupplem2, val(DcbCar2.BoundText)
'If Me.TxtModFlg.Text <> "R" Then
'RetriveClinCounr val(val(DcbCar2.BoundText)), 1
'Calc

'End If
End Sub

Private Sub DcbSupplem2_Change()
DcbSupplem2_Click (0)
End Sub

Private Sub DcbSupplem2_Click(Area As Integer)
'If Me.TxtModFlg.Text <> "R" Then
'Calc
'End If
End Sub

Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode1.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub










Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional Emp_id As String, Optional Typ As Integer = 0)

Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where id = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where Emp_id=" & Emp_id & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
If Typ <> 0 Then
DcbCar.BoundText = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
End If
'DcEmployee2.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
If Typ <> 0 Then
DcbCar.BoundText = 0
End If
End If

End Sub

Private Sub DcbCar_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "TravelRports"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub ChangeLang()
    'Label1.Caption = "Des"

    Me.Caption = "Transportation Reports"
    Me.MainTab.TabCaption(0) = Me.Caption
    LblAccountName.Caption = Me.Caption

    OptAccount(7).Caption = "Detail Trips Report"
    OptAccount(8).Caption = "Total Trips Report"
    OptAccount(9).Caption = "Trips Profit Report"
    Frame2.Caption = "Trip Data"
    lbl(25).Caption = "From"
    lbl(42).Caption = "To"
    lbl(41).Caption = "Location"
    lbl(26).Caption = "Vehicle"
    lbl(28).Caption = "Driver"
 
    Ele(1).Caption = "Period"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
 
    Cmd.Caption = "Exit"

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

        Case 7
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            cAccountReport.ShowTripsData val(Dcemp.BoundText), val(DcCityFromId.BoundText), val(DcCityToId.BoundText), val(Dccar.BoundText)
            Set cAccountReport = Nothing
    
        Case 8
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            cAccountReport.ShowTripsData val(Dcemp.BoundText), val(DcCityFromId.BoundText), val(DcCityToId.BoundText), val(Dccar.BoundText), True
            Set cAccountReport = Nothing
    
        Case 9
            Screen.MousePointer = 11
           ' Set cAccountReport = New ClsAccReports
           ' cAccountReport.BegineDate = Me.DTPickerAccFrom.value
           ' cAccountReport.EndDate = Me.DTPickerAccTo.value
           ' cAccountReport.ShowTripsData val(DCEmp.BoundText), val(DcCityFromId.BoundText), val(DcCityToId.BoundText), val(DCCar.BoundText), True, 1
           ' Set cAccountReport = Nothing
             print_reportCarsProfit
        Case 1

            If Not IsNumeric(Text1.Text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·«„—": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowProductOrderExpenses val(Text1.Text)
            Set cAccountReport = Nothing

        Case 2

            If Not IsNumeric(Txt_order_no.Text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·ÿ·»Ì…": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersiTEMS Txt_order_no.Text
            Set cAccountReport = Nothing

        Case 3
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersStatus Txt_order_no.Text
            Set cAccountReport = Nothing

        Case 7
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            cAccountReport.ShowProductionSummury
            Set cAccountReport = Nothing
                
    End Select

    CuurentLogdata

End Sub
  Function print_reportCarsProfit()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = " SELECT     *"
    MySQL = MySQL & " FROM         (SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS SumValue, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteType,"
    MySQL = MySQL & "                                          dbo.TblCarsData.BoardNO , dbo.TblCarsData.Name, dbo.DOUBLE_ENTREY_VOUCHERS.CarID"
    MySQL = MySQL & "                    FROM         dbo.Notes LEFT OUTER JOIN"
    MySQL = MySQL & "                                          dbo.DOUBLE_ENTREY_VOUCHERS LEFT OUTER JOIN"
    MySQL = MySQL & "                                          dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id ON"
    MySQL = MySQL & "                                          dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_id"
    MySQL = MySQL & "                    WHERE     (dbo.Notes.NoteType = 3) AND (NOT (dbo.DOUBLE_ENTREY_VOUCHERS.Carid IS NULL)) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Carid <> 0) AND"
    MySQL = MySQL & "                                          (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    MySQL = MySQL & "                    GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteType, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name,"
    MySQL = MySQL & "                                          dbo.DOUBLE_ENTREY_VOUCHERS.CarID"
    MySQL = MySQL & "                    Union"
    MySQL = MySQL & "                    SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS SumValue, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteType,"
    MySQL = MySQL & "                                          dbo.TblCarsData.BoardNO , dbo.TblCarsData.Name, dbo.DOUBLE_ENTREY_VOUCHERS.CarID"
    MySQL = MySQL & "                    FROM         dbo.Notes LEFT OUTER JOIN"
    MySQL = MySQL & "                                          dbo.DOUBLE_ENTREY_VOUCHERS LEFT OUTER JOIN"
    MySQL = MySQL & "                                          dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id ON"
    MySQL = MySQL & "                                          dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_id"
    MySQL = MySQL & "                    WHERE     (dbo.Notes.NoteType = 9080) AND (NOT (dbo.DOUBLE_ENTREY_VOUCHERS.Carid IS NULL)) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Carid <> 0) AND"
    MySQL = MySQL & "                                          (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"
        If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    MySQL = MySQL & "                    GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteType, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name,"
    MySQL = MySQL & "                                          dbo.DOUBLE_ENTREY_VOUCHERS.Carid) XB"
    MySQL = MySQL & "    Where (1 = 1)"

       
    If val(Me.Dccar.BoundText) <> 0 Then
        MySQL = MySQL & " and  XB.Carid =" & val(Dccar.BoundText) & ""
    End If

        MySQL = MySQL & " order by XB.Carid"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptCarProfit.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptCarProfitE.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
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
Function print_report2(Optional NoteSerial As String)
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   ' StrSQL = "SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteType, dbo.notes_all.CityFromId, TblCountriesGovernments_2.GovernmentName, dbo.notes_all.CityToId, "
   ' StrSQL = StrSQL & "                  TblCountriesGovernments_1.GovernmentName AS ToGovernmentName, dbo.notes_all.general_des, dbo.notes_all.CusID, TblCustemers_2.CusName,"
   ' StrSQL = StrSQL & "                  TblCustemers_2.CusNamee, TblCustemers_2.Fullcode, dbo.notes_all.NoteDate, dbo.notes_all.LeaderName, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name,"
   ' StrSQL = StrSQL & "                  dbo.TblEmployee.Fullcode AS DriverCode, dbo.TblEmployee.Emp_Namee, dbo.notes_all.VendorID, TblCustemers_1.CusName AS VendorName,"
   ' StrSQL = StrSQL & "                  TblCustemers_1.CusNamee AS VendorNameE, TblCustemers_1.Fullcode AS VendorCode, dbo.notes_all.CarId, dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2,"
   ' StrSQL = StrSQL & "                  dbo.TblVendorCars.BoardNo AS BoardNo2, ISNULL(dbo.GetTrivOwnVale(dbo.notes_all.CityFromId, dbo.notes_all.CityToId), 0) AS OwnValu,"
   ' StrSQL = StrSQL & "                  ISNULL(dbo.GetTrivNotOwnValeComplete(dbo.notes_all.CityFromId, dbo.notes_all.CityToId, dbo.notes_all.VendorID), 0) AS NotOwnValeComplete,"
   ' StrSQL = StrSQL & "                  ISNULL(dbo.GetTrivNotOwnVale(dbo.notes_all.CityFromId, dbo.notes_all.CityToId, dbo.notes_all.VendorID), 0) AS NotOwnVale,"
   ' StrSQL = StrSQL & "                  dbo.SumQtyUpload(dbo.notes_all.NoteID) AS QtyUpload, dbo.SumQtyDownload(dbo.notes_all.NoteID) AS QtyDownload, dbo.notes_all.branch_no,"
   ' StrSQL = StrSQL & "                  dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee, dbo.notes_all.VehicleType, dbo.notes_all.CarType, ISNULL(dbo.notes_all.SupplemID2, 0)"
   ' StrSQL = StrSQL & "                  AS part"
   ' StrSQL = StrSQL & "   FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
   ' StrSQL = StrSQL & "                   dbo.notes_all ON dbo.TblBranchesData.branch_id = dbo.notes_all.branch_no LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_1 ON dbo.notes_all.VendorID = TblCustemers_1.CusID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_2 ON dbo.notes_all.CusID = TblCustemers_2.CusID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "                  dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID"
   ' StrSQL = StrSQL & "    Where (dbo.notes_all.NoteType = 370)"
  StrSQL = "  SELECT        dbo.notes_all.NoteID, dbo.notes_all.NoteType, dbo.notes_all.CityFromId, TblCountriesGovernments_2.GovernmentName, dbo.notes_all.CityToId,"
  StrSQL = StrSQL & "                       TblCountriesGovernments_1.GovernmentName AS ToGovernmentName, dbo.notes_all.general_des, dbo.notes_all.CusID, TblCustemers_2.CusName, TblCustemers_2.CusNamee, TblCustemers_2.Fullcode,"
  StrSQL = StrSQL & "                       dbo.notes_all.NoteDate, dbo.notes_all.LeaderName, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS DriverCode, dbo.TblEmployee.Emp_Namee, dbo.notes_all.VendorID,"
  StrSQL = StrSQL & "                       TblCustemers_1.CusName AS VendorName, TblCustemers_1.CusNamee AS VendorNameE, TblCustemers_1.Fullcode AS VendorCode, dbo.notes_all.CarId, dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2,"
  StrSQL = StrSQL & "                       dbo.TblVendorCars.BoardNo AS BoardNo2, ISNULL(dbo.GetTrivOwnVale(dbo.notes_all.CityFromId, dbo.notes_all.CityToId), 0) AS OwnValu, ISNULL(dbo.GetTrivNotOwnValeComplete(dbo.notes_all.CityFromId,"
  StrSQL = StrSQL & "                       dbo.notes_all.CityToId, dbo.notes_all.VendorID), 0) AS NotOwnValeComplete, ISNULL(dbo.GetTrivNotOwnVale(dbo.notes_all.CityFromId, dbo.notes_all.CityToId, dbo.notes_all.VendorID), 0) AS NotOwnVale,"
  StrSQL = StrSQL & "                       dbo.SumQtyUpload(dbo.notes_all.NoteID) AS QtyUpload, dbo.SumQtyDownload(dbo.notes_all.NoteID) AS QtyDownload, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name,"
  StrSQL = StrSQL & "                       dbo.TblBranchesData.branch_namee, dbo.notes_all.VehicleType, dbo.notes_all.CarType, ISNULL(dbo.notes_all.SupplemID2, 0) AS part, dbo.TblTripTypesTransport.CardNO,"
  StrSQL = StrSQL & "                       dbo.TblTripTypesTransport.QtyDownload AS QtyDownloadDet, dbo.TblTripTypesTransport.CardNO2, dbo.TblTripTypesTransport.QtyDischarge, dbo.TblTripTypesTransport.BillDate,"
  StrSQL = StrSQL & "                       dbo.TblTripTypesTransport.ItemID , dbo.TblItems.ItemName, dbo.TblItems.itemcode, dbo.TblItems.ItemNamee, dbo.notes_all.ShipID, dbo.TblShipsData.Name, dbo.TblShipsData.NameE"
  StrSQL = StrSQL & "     FROM            dbo.TblShipsData RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.notes_all ON dbo.TblShipsData.id = dbo.notes_all.ShipID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblTripTypesTransport LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblItems ON dbo.TblTripTypesTransport.ItemID = dbo.TblItems.ItemID ON dbo.notes_all.NoteID = dbo.TblTripTypesTransport.NotesallID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblCustemers AS TblCustemers_1 ON dbo.notes_all.VendorID = TblCustemers_1.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblCustemers AS TblCustemers_2 ON dbo.notes_all.CusID = TblCustemers_2.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                       dbo.TblCountriesGovernments AS TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblCountriesGovernments AS TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID"
  StrSQL = StrSQL & "  Where (dbo.notes_all.NoteType = 370)"
    If val(dcBranch(0).BoundText) <> 0 And dcBranch(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.branch_no = " & val(dcBranch(0).BoundText) & ""
    End If
    If val(DBCboClientName.BoundText) <> 0 And DBCboClientName.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CusID = " & val(DBCboClientName.BoundText) & ""
    End If
    If ChCarType(0).value = True Then
        StrSQL = StrSQL & " AND  notes_all.CarType = 0"
    ElseIf ChCarType(1).value = True Then
        StrSQL = StrSQL & " AND  notes_all.CarType = 1"
    End If
    If val(DcCityFromId2.BoundText) <> 0 And DcCityFromId2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CityFromId = " & val(DcCityFromId2.BoundText) & ""
    End If
    If val(DcCityToId2.BoundText) <> 0 And DcCityToId2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CityToId = " & val(DcCityToId2.BoundText) & ""
    End If
    If val(VehicleType(0).BoundText) <> 0 And VehicleType(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.VehicleType(0) = " & val(VehicleType(0).BoundText) & ""
    End If
    If val(DCCar2.BoundText) <> 0 And DCCar2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CarId = " & val(DCCar2.BoundText) & ""
    End If
    If val(DCCar3.BoundText) <> 0 And DCCar3.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CarID2 = " & val(DCCar3.BoundText) & ""
    End If
    If val(DcbEmployee(0).BoundText) <> 0 And DcbEmployee(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.DriverId = " & val(DcbEmployee(0).BoundText) & ""
    End If
    If val(DBCboClientName2.BoundText) <> 0 And DBCboClientName2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.VendorID = " & val(DBCboClientName2.BoundText) & ""
    End If
    If val(DcbShip.BoundText) <> 0 And DcbShip.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.ShipID = " & val(DcbShip.BoundText) & ""
    End If
   If Not IsNull(FrmDate.value) Then
            StrSQL = StrSQL & " AND dbo.TblTripTypesTransport.BillDate >=" & SQLDate(FrmDate.value, True) & ""
    End If
    If Not IsNull(todate.value) Then
            StrSQL = StrSQL & " AND dbo.TblTripTypesTransport.BillDate <=" & SQLDate(todate.value, True) & ""
    End If

    

    If RdSort(0).value = True Then
    sql = sql & " ORDER BY dbo.notes_all.CarId"
    ElseIf RdSort(1).value = True Then
    sql = sql & " ORDER BY dbo.notes_all.CarID2"
    ElseIf RdSort(2).value = True Then
    sql = sql & " ORDER BY dbo.notes_all.VendorID"
    ElseIf RdSort(3).value = True Then
    sql = sql & " ORDER BY dbo.notes_all.DriverId"
    ElseIf RdSort(4).value = True Then
    sql = sql & " ORDER BY dbo.notes_all.CusID"
    ElseIf RdSort(7).value = True Then
    sql = sql & " ORDER BY dbo.TblTripTypesTransport.BillDate"
    ElseIf RdSort(9).value = True Then
    sql = sql & " ORDER BY dbo.TblTripTypesTransport.QtyDownload"
     ElseIf RdSort(10).value = True Then
    sql = sql & " ORDER BY dbo.TblTripTypesTransport.QtyDischarge "
    Else
    sql = sql & " ORDER BY dbo.notes_all.NoteSerial1"
    End If
  
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans4.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans4E.rpt"
       End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
    
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
 
    End If

      xReport.ParameterFields(3).AddCurrentValue user_name
If Not IsNull(FrmDate.value) Then
    xReport.ParameterFields(4).AddCurrentValue FrmDate.value
End If
If Not IsNull(todate.value) Then
    xReport.ParameterFields(5).AddCurrentValue todate.value
End If
Msg = "«·Ê“‰ «·’«›Ì Â· ÂÊ ﬂ„Ì… «· ›—Ì€"
If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
xReport.ParameterFields(6).AddCurrentValue 1
Else
xReport.ParameterFields(6).AddCurrentValue 0
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
Function print_report(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    StrSQL = " SELECT     dbo.TblTripTypesTransport.ID, dbo.TblTripTypesTransport.NotesallID, dbo.TblTripTypesTransport.CardNO, "
    StrSQL = StrSQL & "                  dbo.TblTripTypesTransport.QtyDownload, dbo.TblTripTypesTransport.CardNO2, dbo.TblTripTypesTransport.QtyDischarge, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_namee, dbo.notes_all.NoteDate, dbo.notes_all.NoteSerial1, dbo.notes_all.general_des, dbo.notes_all.CityFromId,"
    StrSQL = StrSQL & "                   TblCountriesGovernments_2.GovernmentName, dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS GovernmentNameTO,"
    StrSQL = StrSQL & "                   dbo.notes_all.VehicleType, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.notes_all.CarId, dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2,"
    StrSQL = StrSQL & "                   dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.notes_all.NoteID, dbo.notes_all.NoteType, dbo.notes_all.branch_no, dbo.notes_all.CarType, dbo.notes_all.ShipID,"
    StrSQL = StrSQL & "                   dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee, dbo.notes_all.LeaderName, dbo.TblTripTypesTransport.allocations, dbo.TblTripTypesTransport.BillDate,"
    StrSQL = StrSQL & "                   dbo.TblTripTypesTransport.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.notes_all.CusID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                   dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.notes_all.VendorID, TblCustemers_1.CusName AS VenCusName,"
    StrSQL = StrSQL & "                   TblCustemers_1.CusNamee AS VenCusNameE, TblCustemers_1.Fullcode AS VenFullcode, dbo.notes_all.HarborID, dbo.TblHarborsData.Name AS HarabName,"
    StrSQL = StrSQL & "                   dbo.TblHarborsData.NameE AS HarabNameE, dbo.notes_all.TypeTransportID, dbo.TblTypesTransport.Name AS TypeTransName,"
    StrSQL = StrSQL & "                   dbo.TblTypesTransport.NameE AS TypeTransNameE"
    StrSQL = StrSQL & "       FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.notes_all LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblTypesTransport ON dbo.notes_all.TypeTransportID = dbo.TblTypesTransport.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblHarborsData ON dbo.notes_all.HarborID = dbo.TblHarborsData.id ON TblCustemers_1.CusID = dbo.notes_all.VendorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblShipsData ON dbo.notes_all.ShipID = dbo.TblShipsData.id RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblItems RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblTripTypesTransport ON dbo.TblItems.ItemID = dbo.TblTripTypesTransport.ItemID ON"
    StrSQL = StrSQL & "                   dbo.notes_all.NoteID = dbo.TblTripTypesTransport.NotesallID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & "           Where (dbo.notes_all.NoteType = 370)"
    
    If val(dcBranch(0).BoundText) <> 0 And dcBranch(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.branch_no = " & val(dcBranch(0).BoundText) & ""
    End If
    
    If val(DBCboClientName.BoundText) <> 0 And DBCboClientName.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CusID = " & val(DBCboClientName.BoundText) & ""
    End If
    
    If val(DcCityFromId2.BoundText) <> 0 And DcCityFromId2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CityFromId = " & val(DcCityFromId2.BoundText) & ""
    End If
    
    If val(DcCityToId2.BoundText) <> 0 And DcCityToId2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CityToId = " & val(DcCityToId2.BoundText) & ""
    End If
    
    If ChCarType(0).value = True Then
        StrSQL = StrSQL & " AND  notes_all.CarType = 0"
    ElseIf ChCarType(1).value = True Then
        StrSQL = StrSQL & " AND  notes_all.CarType = 1"
    End If
    
    If val(DcbTypeTransport.BoundText) <> 0 And DcbTypeTransport.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.TypeTransportID = " & val(DcbTypeTransport.BoundText) & ""
    End If
    
    If val(DCboItemS.BoundText) <> 0 And DCboItemS.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.TblTripTypesTransport.ItemID = " & val(DCboItemS.BoundText) & ""
    End If
    
    If val(DcbHarbor.BoundText) <> 0 And DcbHarbor.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.HarborID = " & val(DcbHarbor.BoundText) & ""
    End If
    
    If val(DcbShip.BoundText) <> 0 And DcbShip.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.ShipID = " & val(DcbShip.BoundText) & ""
    End If
    
    If val(VehicleType(0).BoundText) <> 0 And VehicleType(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.VehicleType = " & val(VehicleType(0).BoundText) & ""
    End If
    
    If val(DCCar2.BoundText) <> 0 And DCCar2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CarId = " & val(DCCar2.BoundText) & ""
    End If
    
    If val(DCCar3.BoundText) <> 0 And DCCar3.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.CarID2 = " & val(DCCar3.BoundText) & ""
    End If
    
    If val(DcbEmployee(0).BoundText) <> 0 And DcbEmployee(0).Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.DriverId = " & val(DcbEmployee(0).BoundText) & ""
    End If
    
    If val(DBCboClientName2.BoundText) <> 0 And DBCboClientName2.Text <> "" Then
    StrSQL = StrSQL & " and  dbo.notes_all.VendorID = " & val(DBCboClientName2.BoundText) & ""
    End If
    
    If Me.TxtCard1.Text <> "" Then
           StrSQL = StrSQL & " AND  dbo.TblTripTypesTransport.CardNO Like N'%" & (Me.TxtCard1.Text) & "%'"
    End If
    
    If Me.TxtCard2.Text <> "" Then
           StrSQL = StrSQL & " AND  dbo.TblTripTypesTransport.CardNO2 Like N'%" & (Me.TxtCard2.Text) & "%'"
    End If
    
    If ChCarType(3).value = True Then
     StrSQL = StrSQL & " AND  TblTripTypesTransport.allocations =1"
    End If
    
    If ChCarType(4).value = True Then
     StrSQL = StrSQL & " AND ( TblTripTypesTransport.allocations =0 or TblTripTypesTransport.allocations is null) "
    End If
    
   If Not IsNull(FrmDate.value) Then
            StrSQL = StrSQL & " AND dbo.TblTripTypesTransport.BillDate >=" & SQLDate(FrmDate.value, True) & ""
    End If
    
    If Not IsNull(todate.value) Then
            StrSQL = StrSQL & " AND dbo.TblTripTypesTransport.BillDate <=" & SQLDate(todate.value, True) & ""
    End If
    
    If val(TxtQtyUpload.Text) <> 0 Then
    If RdTotal(0).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDownload >" & val(TxtQtyUpload.Text) & ""
    ElseIf RdTotal(1).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDownload <" & val(TxtQtyUpload.Text) & ""
    ElseIf RdTotal(2).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDownload =" & val(TxtQtyUpload.Text) & ""
    ElseIf RdTotal(3).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDownload >=" & val(TxtQtyUpload.Text) & ""
    ElseIf RdTotal(4).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDownload <=" & val(TxtQtyUpload.Text) & ""
    End If
    End If
    
    If val(TxtQtyDownload.Text) <> 0 Then
    If RdTotal(0).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDischarge >" & val(TxtQtyDownload.Text) & ""
    ElseIf RdTotal(1).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDischarge <" & val(TxtQtyDownload.Text) & ""
    ElseIf RdTotal(2).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDischarge =" & val(TxtQtyDownload.Text) & ""
    ElseIf RdTotal(3).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDischarge >=" & val(TxtQtyDownload.Text) & ""
    ElseIf RdTotal(4).value = True Then
    StrSQL = StrSQL & " AND  TblTripTypesTransport.QtyDischarge <=" & val(TxtQtyDownload.Text) & ""
    End If
    End If
    If RdSort(0).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.CarId"
    ElseIf RdSort(1).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.CarID2"
    ElseIf RdSort(2).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.VendorID"
    ElseIf RdSort(3).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.DriverId"
    ElseIf RdSort(4).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.CusID"
    ElseIf RdSort(5).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.TypeTransportID"
    ElseIf RdSort(6).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.ItemID"
    ElseIf RdSort(7).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.NoteDate"
    ElseIf RdSort(8).value = True Then
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.ShipID"
    ElseIf RdSort(9).value = True Then
    StrSQL = StrSQL & " ORDER BY TblTripTypesTransport.QtyDownload"
    ElseIf RdSort(10).value = True Then
    StrSQL = StrSQL & " ORDER BY TblTripTypesTransport.QtyDischarge"
    Else
    StrSQL = StrSQL & " ORDER BY dbo.notes_all.NoteSerial1"
    End If
    If RdPrint(0).value = True Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans1E.rpt"
        End If
    ElseIf RdPrint(1).value = True Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans2E.rpt"
        End If
   ElseIf RdPrint(2).value = True Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans3E.rpt"
        End If
      ElseIf RdPrint(3).value = True Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans4.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTrans3E.rpt"
        End If
    End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
        StrReportTitle = ""
 
    End If

      xReport.ParameterFields(3).AddCurrentValue user_name
If Not IsNull(FrmDate.value) Then
    xReport.ParameterFields(4).AddCurrentValue FrmDate.value
End If
If Not IsNull(todate.value) Then
    xReport.ParameterFields(5).AddCurrentValue todate.value
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
Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode
End Sub

Private Sub DBCboClientName2_Change()
DBCboClientName2_Click (0)
End Sub

Private Sub DBCboClientName2_Click(Area As Integer)
    Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName2.BoundText), , Fullcode, 2
    Text2.Text = Fullcode
End Sub

Private Sub DcboItems_Change()
DcboItems_Click (0)
End Sub

Private Sub DcboItems_Click(Area As Integer)
Me.TxtItemCode.Text = GetItemCode(val(Me.DCboItemS.BoundText))
End Sub

Private Sub ISButton1_Click()
If RdPrint(3).value = False And RdPrint(2).value = False And RdPrint(1).value = False And RdPrint(0).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «· ﬁ—Ì—"
Else
MsgBox "Please Select Report"
End If
Exit Sub
End If
If RdPrint(3).value = False Then
print_report
Else
print_report2
End If
End Sub
Private Sub ISButton2_Click()

    On Error Resume Next
    
    Dim MySQL As String
    Dim sqlStr2 As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim mFromDate As String
    Dim mToDate As String
    mFromDate = FromDate2.value
    mToDate = ToDate2.value
    StrSQL = " "
        If val(dcBranch(1).BoundText) <> 0 And dcBranch(1).Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.BranchID = " & val(dcBranch(1).BoundText) & ""
    End If
    
    If val(DcbCar.BoundText) <> 0 And DcbCar.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.CarID = " & val(DcbCar.BoundText) & ""
    End If
    
    If val(DcbCar2.BoundText) <> 0 And DcbCar2.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.CarID2 = " & val(DcbCar2.BoundText) & ""
    End If

    If val(DcEmployee2.BoundText) <> 0 And DcEmployee2.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.EmpID= " & val(DcEmployee2.BoundText) & ""
    End If
    
    If val(DBCboClientName1.BoundText) <> 0 And DBCboClientName1.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.CUSTID1= " & val(DBCboClientName1.BoundText) & ""
    End If
    
    If val(DcbSupplem2.BoundText) <> 0 And DcbSupplem2.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.SupplemID2= " & val(DcbSupplem2.BoundText) & ""
    End If
    
    If val(DcbSupplem.BoundText) <> 0 And DcbSupplem.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.SupplemID= " & val(DcbSupplem.BoundText) & ""
    End If
     
    If Trim(TxtLeaderName) <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.LeaderName = N'" & Trim(TxtLeaderName) & "'"
    End If
            
    If Me.txtid.Text <> "" Then
           StrSQL = StrSQL & " AND  TblOrderUpload.ID = " & val(Me.txtid.Text)
    End If
    If optRepMonth Then
        StrSQL = StrSQL & " AND Month(TblOrderUpload.RecordDate) >=" & cmbFromMonthName.ListIndex + 1
        StrSQL = StrSQL & " AND Month(TblOrderUpload.RecordDate) <=" & cmbToMonthName.ListIndex + 1
        
        StrSQL = StrSQL & " AND Year(TblOrderUpload.RecordDate) >=" & val(cmbFromYear.Text)
        StrSQL = StrSQL & " AND Year(TblOrderUpload.RecordDate) <=" & val(cmbToYear.Text)
        
    Else
        If Not IsNull(FromDate2.value) Then
            StrSQL = StrSQL & " AND TblOrderUpload.RecordDate >=" & SQLDate(FromDate2.value, True) & ""
        End If
        
        If Not IsNull(ToDate2.value) Then
            StrSQL = StrSQL & " AND TblOrderUpload.RecordDate <=" & SQLDate(ToDate2.value, True) & ""
        End If
    End If
    If val(orderStatus.ListIndex) <> -1 And orderStatus.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.orderStatus = " & val(orderStatus.ListIndex) & ""
    End If
    
      If val(cmbTypeRep.ListIndex) <> -1 And cmbTypeRep.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.TypeRep = " & val(cmbTypeRep.ListIndex) & ""
    End If
    
    
    If val(carStatus.ListIndex) <> -1 And carStatus.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.carStatus = " & val(carStatus.ListIndex) & ""
    End If
    
    
    If val(DcEmpSuper.BoundText) <> 0 And DcEmpSuper.Text <> "" Then
        StrSQL = StrSQL & " and  dbo.TblOrderUpload.DcEmpSuper = " & val(DcEmpSuper.BoundText) & ""
    End If
    
   If KPI.value = True Then
   
   MySQL = "SELECT     COUNT(dbo.TblOrderUpload.ID) AS CountOFTrips, SUM(dbo.TblOrderUpload.Price) AS SUMOFTrips, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
     MySQL = MySQL & "                  SUM(dbo.TblOrderUpload.Price) / COUNT(dbo.TblOrderUpload.ID) AS AVGVALUE"
MySQL = MySQL & "  FROM         dbo.TblOrderUpload LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblOrderUpload.CustId1 = dbo.TblCustemers.CusID"
MySQL = MySQL & " Where 1 = 1  " & StrSQL
MySQL = MySQL & "  GROUP BY dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblOrderUpload.CustId1"


GoTo AVGTRip
   End If
    
   If optRepMonth Then
    MySQL = "SELECT CarID,Cars.BoardNO,"
    MySQL = MySQL & " MONTH(RecordDate)     Montht,"
    MySQL = MySQL & "    YEAR(RecordDate)      Yearr"
    MySQL = MySQL & " From TblOrderUpload LEFT OUTER JOIN TblCarsData Cars"
    MySQL = MySQL & "             ON  Cars.id = TblOrderUpload.CarID WHERE 1 = 1 "
    
    



    
        sqlStr2 = "        SELECT DISTINCT                         TripStatusID,"
        sqlStr2 = sqlStr2 & " MONTH(TblOrderUpload.RecordDate)             Montht,"
        sqlStr2 = sqlStr2 & " YEAR(TblOrderUpload.RecordDate)              Yearr,"
        sqlStr2 = sqlStr2 & " CarID,"
        sqlStr2 = sqlStr2 & " TblOrderUpload.EmpID DcEmpSuper,"
        sqlStr2 = sqlStr2 & " TblOrderUpload.CityID,"
        sqlStr2 = sqlStr2 & " TblOrderUpload.CityID2,"
        sqlStr2 = sqlStr2 & " CustId1,"
        sqlStr2 = sqlStr2 & " tc.CusName,"
        sqlStr2 = sqlStr2 & " te.Emp_Name,"
        sqlStr2 = sqlStr2 & " City.GovernmentName,"
        sqlStr2 = sqlStr2 & " City2.GovernmentName             GovernmentName2,"
        sqlStr2 = sqlStr2 & " TblOrderUpload.RecordDate,"
        sqlStr2 = sqlStr2 & " TblOrderUpload.Price + ISNULL(TblOrderUpload.PartPrice, 0) AS Total"
        sqlStr2 = sqlStr2 & " FROM   TblOrderUpload "
        sqlStr2 = sqlStr2 & " LEFT OUTER JOIN TblCarsData Cars"
        sqlStr2 = sqlStr2 & "     ON  Cars.id = TblOrderUpload.CarID"
        sqlStr2 = sqlStr2 & " LEFT OUTER JOIN TblCustemers  AS tc"
        sqlStr2 = sqlStr2 & " ON  TC.CusID = TblOrderUpload.CustId1"
        sqlStr2 = sqlStr2 & " LEFT OUTER JOIN TblEmployee   AS te"
        sqlStr2 = sqlStr2 & "            ON  te.Emp_ID = TblOrderUpload.EmpID"
        sqlStr2 = sqlStr2 & "       LEFT OUTER JOIN TblCountriesGovernments City"
        sqlStr2 = sqlStr2 & "            ON  City.GovernmentID = TblOrderUpload.CityID"
        sqlStr2 = sqlStr2 & "       LEFT OUTER JOIN TblCountriesGovernments City2"
        sqlStr2 = sqlStr2 & "            ON  City2.GovernmentID = TblOrderUpload.CityID2"
        sqlStr2 = sqlStr2 & " Where 1 = 1"
        
        
        
        
        
sqlStr2 = " SELECT * "
sqlStr2 = sqlStr2 & vbNewLine & " FROM   ("
sqlStr2 = sqlStr2 & vbNewLine & "            SELECT *"
sqlStr2 = sqlStr2 & vbNewLine & "            FROM   ("
sqlStr2 = sqlStr2 & vbNewLine & "                       SELECT DISTINCT TripStatusID,"
sqlStr2 = sqlStr2 & vbNewLine & "                              MONTH(TblOrderUpload.RecordDate) Montht,"
sqlStr2 = sqlStr2 & vbNewLine & "                              YEAR(TblOrderUpload.RecordDate) Yearr,"
sqlStr2 = sqlStr2 & vbNewLine & "                              CarID,"
sqlStr2 = sqlStr2 & vbNewLine & "                              TblOrderUpload.EmpID DcEmpSuper,"
sqlStr2 = sqlStr2 & vbNewLine & "                              TblOrderUpload.CityID,"
sqlStr2 = sqlStr2 & vbNewLine & "                              TblOrderUpload.CityID2,"
sqlStr2 = sqlStr2 & vbNewLine & "                              CustId1,"
sqlStr2 = sqlStr2 & vbNewLine & "                              tc.CusName,"
sqlStr2 = sqlStr2 & vbNewLine & "                              te.Emp_Name,"
sqlStr2 = sqlStr2 & vbNewLine & "                              City.GovernmentName,"
sqlStr2 = sqlStr2 & vbNewLine & "                              City2.GovernmentName GovernmentName2,"
sqlStr2 = sqlStr2 & vbNewLine & "                              TblOrderUpload.RecordDate,"
sqlStr2 = sqlStr2 & vbNewLine & "                              TblOrderUpload.Price + ISNULL(TblOrderUpload.PartPrice, 0) AS Total"
sqlStr2 = sqlStr2 & vbNewLine & "                       From TblTypesTripStatus"
sqlStr2 = sqlStr2 & vbNewLine & "                              INNER JOIN TblOrderUpload"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  TblOrderUpload.TripStatusID = TblTypesTripStatus.ID"
sqlStr2 = sqlStr2 & vbNewLine & "                              LEFT OUTER JOIN TblCarsData Cars"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  Cars.id = TblOrderUpload.CarID"
sqlStr2 = sqlStr2 & vbNewLine & "                              LEFT OUTER JOIN TblCustemers AS tc"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  TC.CusID = TblOrderUpload.CustId1"
sqlStr2 = sqlStr2 & vbNewLine & "                              LEFT OUTER JOIN TblEmployee AS te"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  te.Emp_ID = TblOrderUpload.EmpID"
sqlStr2 = sqlStr2 & vbNewLine & "                              LEFT OUTER JOIN TblCountriesGovernments City"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  City.GovernmentID = TblOrderUpload.CityID"
sqlStr2 = sqlStr2 & vbNewLine & "                              LEFT OUTER JOIN TblCountriesGovernments City2"
sqlStr2 = sqlStr2 & vbNewLine & "                                   ON  City2.GovernmentID = TblOrderUpload.CityID2"
sqlStr2 = sqlStr2 & vbNewLine & "                       Where TblTypesTripStatus.ID = 1"
sqlStr2 = sqlStr2 & vbNewLine & StrSQL
sqlStr2 = sqlStr2 & vbNewLine & "                   ) AS TS1"
Dim i As Integer
Dim sqlStr3 As String
sqlStr3 = ""
For i = 2 To 16
    sqlStr3 = sqlStr3 & vbNewLine & "                   FULL OUTER JOIN ("
    sqlStr3 = sqlStr3 & vbNewLine & "                            SELECT DISTINCT TripStatusID TripStatusID" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   MONTH(TblOrderUpload.RecordDate) Montht" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   YEAR(TblOrderUpload.RecordDate) Yearr" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   CarID CarID" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   TblOrderUpload.EmpID DcEmpSuper" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   CustId1 CustId1" & i & ","
    
    sqlStr3 = sqlStr3 & vbNewLine & "                                   tc.CusName CusName" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   te.Emp_Name Emp_Name" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   City.GovernmentName GovernmentName1" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   City2.GovernmentName GovernmentName2" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   TblOrderUpload.RecordDate RecordDate" & i & ","
    sqlStr3 = sqlStr3 & vbNewLine & "                                   TblOrderUpload.Price + ISNULL(TblOrderUpload.PartPrice, 0) AS Total" & i
    sqlStr3 = sqlStr3 & vbNewLine & "                            From TblTypesTripStatus"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   INNER JOIN TblOrderUpload"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  TblOrderUpload.TripStatusID = TblTypesTripStatus.ID"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   LEFT OUTER JOIN TblCarsData Cars"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  Cars.id = TblOrderUpload.CarID"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   LEFT OUTER JOIN TblCustemers AS tc"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  TC.CusID = TblOrderUpload.CustId1"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   LEFT OUTER JOIN TblEmployee AS te"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  te.Emp_ID = TblOrderUpload.EmpID"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   LEFT OUTER JOIN TblCountriesGovernments City"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  City.GovernmentID = TblOrderUpload.CityID"
    sqlStr3 = sqlStr3 & vbNewLine & "                                   LEFT OUTER JOIN TblCountriesGovernments City2"
    sqlStr3 = sqlStr3 & vbNewLine & "                                        ON  City2.GovernmentID = TblOrderUpload.CityID2"
    sqlStr3 = sqlStr3 & vbNewLine & "                            Where TblTypesTripStatus.ID = " & i
    sqlStr3 = sqlStr3 & vbNewLine & StrSQL
    sqlStr3 = sqlStr3 & vbNewLine & "                        ) TS" & i & ""
    sqlStr3 = sqlStr3 & vbNewLine & "                        ON  TS" & i & ".Montht" & i & " = TS1.Montht"
    sqlStr3 = sqlStr3 & vbNewLine & "                        AND TS" & i & ".Yearr" & i & " = TS1.Yearr"
    sqlStr3 = sqlStr3 & vbNewLine & "                        AND TS" & i & ".CarID" & i & " = TS1.CarID"
                           
Next
                                                                                                                   
sqlStr2 = sqlStr2 & sqlStr3 & vbNewLine & "        ) g"


   End If

    
    
    If RepUploadOrderChk.value = True Then
        MySQL = " SELECT  DISTINCT  TblOrderUpload.Price,TblOrderUpload.DrievType,TblOrderUpload.LeaderName,dbo.TblOrderUpload.CountOrders, dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, IsNull(dbo.TblOrderUpload.TypeRep,-1) TypeRep,"
        MySQL = MySQL & " CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) as TimeOrder,TblTypesTripStatus.Name TripStatus,"
        MySQL = MySQL & " dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
        MySQL = MySQL & " dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
        MySQL = MySQL & " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
        MySQL = MySQL & " dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice,  dbo.TblOrderUpload.Total,"
        MySQL = MySQL & " dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
        MySQL = MySQL & " TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
        MySQL = MySQL & " TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
        MySQL = MySQL & " dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
        MySQL = MySQL & " TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2"
        ',dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
        'MySQL = MySQL & "  dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"

        
MySQL = MySQL & "   , tripid="
MySQL = MySQL & "  ("
MySQL = MySQL & "   select  max( notes_all.NoteSerial1)  from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "  )"
MySQL = MySQL & "   , invoiceid="
MySQL = MySQL & "  ("
 MySQL = MySQL & "   SELECT   max(dbo.TblTravDueK.NoteSerial1)"
MySQL = MySQL & "  FROM         dbo.TblTravDueKDet RIGHT OUTER JOIN"
MySQL = MySQL & "                        dbo.TblTravDueK ON dbo.TblTravDueKDet.TravID = dbo.TblTravDueK.ID"
MySQL = MySQL & "  WHERE     dbo.TblTravDueKDet.TripNo in"
MySQL = MySQL & "      ("
MySQL = MySQL & "      select    max  ( notes_all.NoteSerial1) from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "       )"
MySQL = MySQL & "  )"

        MySQL = MySQL & " FROM dbo.TblUnites RIGHT OUTER JOIN "
        MySQL = MySQL & " dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
        MySQL = MySQL & " dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
        MySQL = MySQL & " LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblTypesTripStatus ON dbo.TblOrderUpload.TripStatusID = dbo.TblTypesTripStatus.Id"
        MySQL = MySQL & "    Where 1 = 1 "
    ElseIf optOrderUploadRec.value = True Then
            MySQL = " SELECT  DISTINCT  TblOrderUpload.Price,TblOrderUpload.DrievType,TblOrderUpload.LeaderName,dbo.TblOrderUpload.CountOrders, dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, IsNull(dbo.TblOrderUpload.TypeRep,-1) TypeRep,"
        MySQL = MySQL & " CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) as TimeOrder,TblTypesTripStatus.Name TripStatus,"
        MySQL = MySQL & " dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
        MySQL = MySQL & " dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
        MySQL = MySQL & " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
        MySQL = MySQL & " dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice,  dbo.TblOrderUpload.Total,"
        MySQL = MySQL & " dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
        MySQL = MySQL & " TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
        MySQL = MySQL & " TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
        MySQL = MySQL & " dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
        MySQL = MySQL & " TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2"
        ',dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
        'MySQL = MySQL & "  dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"

        
        MySQL = MySQL & "   , tripid="
        MySQL = MySQL & "  ("
        MySQL = MySQL & "   select  max( notes_all.NoteSerial1)  from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
        MySQL = MySQL & "  )"
        MySQL = MySQL & "   , invoiceid="
        MySQL = MySQL & "  ("
         MySQL = MySQL & "   SELECT   max(dbo.TblTravDueK.NoteSerial1)"
        MySQL = MySQL & "  FROM         dbo.TblTravDueKDet RIGHT OUTER JOIN"
        MySQL = MySQL & "                        dbo.TblTravDueK ON dbo.TblTravDueKDet.TravID = dbo.TblTravDueK.ID"
        MySQL = MySQL & "  WHERE     dbo.TblTravDueKDet.TripNo in"
        MySQL = MySQL & "      ("
        MySQL = MySQL & "      select    max  ( notes_all.NoteSerial1) from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
        MySQL = MySQL & "       )"
        MySQL = MySQL & "  )"

        MySQL = MySQL & " FROM dbo.TblUnites RIGHT OUTER JOIN "
        MySQL = MySQL & " dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
        MySQL = MySQL & " dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
        MySQL = MySQL & " LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblTypesTripStatus ON dbo.TblOrderUpload.TripStatusID = dbo.TblTypesTripStatus.Id"
        'StrSQL = " Where 1 = 1 "
        MySQL = MySQL & "    Where 1 = 1 "
    ElseIf notInvoicing.value = True Then
    
        MySQL = " SELECT  DISTINCT  TblOrderUpload.Price,TblOrderUpload.DrievType,TblOrderUpload.LeaderName,dbo.TblOrderUpload.CountOrders, dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, IsNull(dbo.TblOrderUpload.TypeRep,-1) TypeRep,"
        MySQL = MySQL & " CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) as TimeOrder,TblTypesTripStatus.Name TripStatus,"
        MySQL = MySQL & " dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
        MySQL = MySQL & " dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
        MySQL = MySQL & " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
        MySQL = MySQL & " dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice,  dbo.TblOrderUpload.Total,"
        MySQL = MySQL & " dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
        MySQL = MySQL & " TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
        MySQL = MySQL & " TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
        MySQL = MySQL & " dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
        MySQL = MySQL & " TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2"
        ',dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
        'MySQL = MySQL & "  dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"


MySQL = MySQL & "   , tripid="
MySQL = MySQL & "  ("
MySQL = MySQL & "   select  max( notes_all.NoteSerial1)  from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "  )"
MySQL = MySQL & "   , invoiceid="
MySQL = MySQL & "  ("
 MySQL = MySQL & "   SELECT   max(dbo.TblTravDueK.NoteSerial1)"
MySQL = MySQL & "  FROM         dbo.TblTravDueKDet RIGHT OUTER JOIN"
MySQL = MySQL & "                        dbo.TblTravDueK ON dbo.TblTravDueKDet.TravID = dbo.TblTravDueK.ID"
MySQL = MySQL & "  WHERE     dbo.TblTravDueKDet.TripNo in"
MySQL = MySQL & "      ("
MySQL = MySQL & "      select    max  ( notes_all.NoteSerial1) from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "       )"
MySQL = MySQL & "  )"
        
        
        MySQL = MySQL & " FROM dbo.TblUnites RIGHT OUTER JOIN "
        MySQL = MySQL & " dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
        MySQL = MySQL & " dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
        MySQL = MySQL & " LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblTypesTripStatus ON dbo.TblOrderUpload.TripStatusID = dbo.TblTypesTripStatus.Id"
              MySQL = MySQL & "   Where  TblOrderUpload.id not  in (select BasedNo from  notes_all where notetype=370 and BasedNo<>0) "
      'here
        

ElseIf Invoicing.value = True Then
    
        MySQL = " SELECT  DISTINCT  TblOrderUpload.Price,TblOrderUpload.DrievType,TblOrderUpload.LeaderName,dbo.TblOrderUpload.CountOrders, IsNull(dbo.TblOrderUpload.TypeRep,-1) TypeRep, dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
        MySQL = MySQL & " CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) as TimeOrder,TblTypesTripStatus.Name TripStatus,"
        MySQL = MySQL & " dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
        MySQL = MySQL & " dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
        MySQL = MySQL & " dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
        MySQL = MySQL & " dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice,  dbo.TblOrderUpload.Total,"
        MySQL = MySQL & " dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
        MySQL = MySQL & " TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
        MySQL = MySQL & " TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
        MySQL = MySQL & " dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
        MySQL = MySQL & " TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2"
MySQL = MySQL & "   , tripid="
MySQL = MySQL & "  ("
MySQL = MySQL & "   select  max( notes_all.NoteSerial1)  from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "  )"
MySQL = MySQL & "   , invoiceid="
MySQL = MySQL & "  ("
 MySQL = MySQL & "   SELECT   max(dbo.TblTravDueK.NoteSerial1)"
MySQL = MySQL & "  FROM         dbo.TblTravDueKDet RIGHT OUTER JOIN"
MySQL = MySQL & "                        dbo.TblTravDueK ON dbo.TblTravDueKDet.TravID = dbo.TblTravDueK.ID"
MySQL = MySQL & "  WHERE     dbo.TblTravDueKDet.TripNo in"
MySQL = MySQL & "      ("
MySQL = MySQL & "      select    max  ( notes_all.NoteSerial1) from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "       )"
MySQL = MySQL & "  )"
        
MySQL = MySQL & "   , RecNo="
MySQL = MySQL & "  ("
MySQL = MySQL & "   select  max( notes_all.RecNo)  from  notes_all where notetype=370 and BasedNo=TblOrderUpload.id"
MySQL = MySQL & "  )"
        
        
        ',dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
        'MySQL = MySQL & "  dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
        MySQL = MySQL & " FROM dbo.TblUnites RIGHT OUTER JOIN "
        MySQL = MySQL & " dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
        MySQL = MySQL & " dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
        MySQL = MySQL & " LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblTypesTripStatus ON dbo.TblOrderUpload.TripStatusID = dbo.TblTypesTripStatus.Id"
              MySQL = MySQL & "   Where  TblOrderUpload.id    in (select BasedNo from  notes_all where notetype=370 and BasedNo<>0) "
        
   
    ElseIf RepCarsDelayChk.value = True Then
    '    MySQL = " SELECT DISTINCT"
    '    MySQL = MySQL & " TblOrderUpload.Price, TblOrderUpload.DrievType, TblOrderUpload.LeaderName, TblOrderUpload.CountOrders, TblOrderUpload.ID, TblOrderUpload.RecordDate AS NoteDate, TblOrderUpload.BranchID,"
    '    MySQL = MySQL & " TblBranchesData.branch_name, TblBranchesData.branch_namee, CONVERT(char(10), TblOrderUpload.TimeOrder, 108) AS TimeOrder, TblTypesTripStatus.Name AS TripStatus, TblOrderUpload.DrievType AS Expr1,"
    '    MySQL = MySQL & " TblOrderUpload.EmpID, TblEmployee.Emp_Name, TblEmployee.Fullcode, TblEmployee.Emp_Namee, TblOrderUpload.IDNo, TblOrderUpload.LeaderName AS Expr2, TblOrderUpload.Nationality, TblOrderUpload.CarType,"
    '    MySQL = MySQL & " TblOrderUpload.CusID, TblCustemers.CusName, TblCustemers.CusNamee, TblCustemers.Fullcode AS CusFullcode, TblOrderUpload.TypGoods, TblOrderUpload.OrderNo, TblOrderUpload.Remarks, TblOrderUpload.PartPrice,"
    '    MySQL = MySQL & " TblOrderUpload.Total, TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, TblOrderUpload.CityID2, TblCountriesGovernments_1.GovernmentName AS ToCity, TblOrderUpload.CarID,"
    '    MySQL = MySQL & " TblCarsData.BoardNO, TblOrderUpload.CarID2, TblVendorCars_2.BoardNo AS BoardNo2, TblOrderUpload.SupplemID, FixedAssets.Name AS SupplemName, FixedAssets.namee AS SupplemNameE, TblOrderUpload.SupplemID2,"
    '    MySQL = MySQL & " TblVendorCars_1.accessory, TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2, TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2, TblOrderUpload.startDate,"
    '    MySQL = MySQL & " TblOrderUpload.distance, TblOrderUpload.chkStop, TblOrderUpload.carStatus, TblOrderUpload.orderStatus, TblOrderUpload.delayHours, TblOrderUpload.EDA, TblOrderUpload.ETA, TblOrderUpload.delayCuz,"
    '    MySQL = MySQL & " TblOrderUpload.supervisorName, TblOrderUpload.TimeOrder AS TimeOrder2, TblOrderUpload.IsTravel"
    '    MySQL = MySQL & " FROM TblUnites RIGHT OUTER JOIN"
    '    MySQL = MySQL & " TravKItemDet1 ON TblUnites.UnitID = TravKItemDet1.UnitID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblItems ON TravKItemDet1.ItemID = TblItems.ItemID RIGHT OUTER JOIN"
    '    MySQL = MySQL & " TblOrderUpload ON TravKItemDet1.MasterID = TblOrderUpload.ID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblCustemers AS TblCustemers_1 ON TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblVendorCars AS TblVendorCars_1 ON TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
    '    MySQL = MySQL & " FixedAssets ON TblOrderUpload.SupplemID = FixedAssets.id LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblVendorCars AS TblVendorCars_2 ON TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblCarsData ON TblOrderUpload.CarID = TblCarsData.id LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblCountriesGovernments AS TblCountriesGovernments_1 ON TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblCountriesGovernments AS TblCountriesGovernments_2 ON TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblCustemers ON TblOrderUpload.CusID = TblCustemers.CusID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblEmployee ON TblOrderUpload.EmpID = TblEmployee.Emp_ID LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblBranchesData ON TblOrderUpload.BranchID = TblBranchesData.branch_id LEFT OUTER JOIN"
    '    MySQL = MySQL & " TblTypesTripStatus ON TblOrderUpload.TripStatusID = TblTypesTripStatus.ID where 1 = 1 "
    
  MySQL = "SELECT DISTINCT "
      MySQL = MySQL & "                    dbo.TblOrderUpload.Price, dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.CountOrders,IsNull(dbo.TblOrderUpload.TypeRep,-1) TypeRep, dbo.TblOrderUpload.ID,"
   MySQL = MySQL & "                       dbo.TblOrderUpload.RecordDate AS NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    MySQL = MySQL & "                       CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) AS TimeOrder, dbo.TblTypesTripStatus.Name AS TripStatus, dbo.TblOrderUpload.DrievType AS Expr1,"
       MySQL = MySQL & "                                   dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblOrderUpload.IDNo,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.LeaderName AS Expr2, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID, TblCustemers_1.CusName,"
    MySQL = MySQL & "                                      TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods, dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.PartPrice, dbo.TblOrderUpload.Total, dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.CityID2, TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.CarID2, TblVendorCars_2.BoardNo AS BoardNO2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName,"
    MySQL = MySQL & "                                      dbo.FixedAssets.namee AS SupplemNameE, dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1,"
    MySQL = MySQL & "                                      TblCustemers_1.CusName AS CusName2, TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2, dbo.TblOrderUpload.startDate,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.distance, dbo.TblOrderUpload.chkStop, dbo.TblOrderUpload.carStatus, dbo.TblOrderUpload.orderStatus, dbo.TblOrderUpload.delayHours,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.EDA, dbo.TblOrderUpload.ETA, dbo.TblOrderUpload.delayCuz, dbo.TblOrderUpload.supervisorName,"
    MySQL = MySQL & "                                      dbo.TblOrderUpload.TimeOrder AS TimeOrder2, dbo.TblOrderUpload.IsTravel, dbo.TblOrderUpload.DcEmpSuper, TblEmployee_1.Emp_Code AS ForemanCode,"
    MySQL = MySQL & "                                      TblEmployee_1.Emp_Name AS ForemanName, TblEmployee_1.Emp_Namee AS ForemanNamee"
    MySQL = MySQL & "                FROM         dbo.TblEmployee TblEmployee_1 INNER JOIN"
    MySQL = MySQL & "                                      dbo.TblOrderUpload ON TblEmployee_1.Emp_ID = dbo.TblOrderUpload.DcEmpSuper LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblUnites RIGHT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID ON dbo.TblOrderUpload.ID = dbo.TravKItemDet1.MasterID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblCustemers TblCustemers_2 ON dbo.TblOrderUpload.CusID = TblCustemers_2.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                                      dbo.TblTypesTripStatus ON dbo.TblOrderUpload.TripStatusID = dbo.TblTypesTripStatus.ID"
    MySQL = MySQL & "                WHERE     (1 = 1)"
    End If
    

    
    MySQL = MySQL & StrSQL
  '  sqlStr2 = sqlStr2 & StrSQL
    If optRepMonth.value = True Then
        MySQL = MySQL & " Group By"
        MySQL = MySQL & "     YEAR(RecordDate),"
        MySQL = MySQL & "     MONTH(RecordDate),"
        MySQL = MySQL & "        TblOrderUpload.CarID,BoardNO"
'
'        sqlStr2 = sqlStr2 & " Order By"
'        sqlStr2 = sqlStr2 & "       YEAR(TblOrderUpload.RecordDate),"
'        sqlStr2 = sqlStr2 & "       MONTH(TblOrderUpload.RecordDate),"
'        sqlStr2 = sqlStr2 & "       CarID"
    End If
    
AVGTRip:
    If optOrderUploadRec.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUploadRec.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUploadRec.rpt"
        End If
    
    End If
    If RepUploadOrderChk.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod2.rpt"
        End If
        
    ElseIf notInvoicing.value = True Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaodnotInvoiced.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaodnotInvoiced.rpt"
        End If
    
    ElseIf optRepMonth.value = True Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMonth.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMonth.rpt"
        End If
    
    
    
    ElseIf Invoicing.value = True Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaodInvoiced.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaodInvoiced.rpt"
        End If
    
    
    ElseIf RepCarsDelayChk.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod3.rpt"
        End If
     ElseIf KPI.value = True Then
     
           If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TripSAVG.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TripSAVG.rpt"
        End If
        
    End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData


    If sqlStr2 <> "" And optRepMonth Then
        Dim RsData2  As New ADODB.Recordset
        Dim RsData3  As New ADODB.Recordset
         
        RsData2.Open sqlStr2, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("Det").Database.SetDataSource RsData2
        
'        RsData3.Open sqlStr2, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        xReport.OpenSubreport("TotalIq").Database.SetDataSource RsData3
        
    End If
    
    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    
    If Not IsNull(FromDate2.value) Then
        xReport.ParameterFields(4).AddCurrentValue FromDate2.value
    End If
       
    If Not IsNull(ToDate2.value) Then
        xReport.ParameterFields(5).AddCurrentValue ToDate2.value
    End If
       
    If dcBranch(1).Text <> "" Then
        xReport.ParameterFields(6).AddCurrentValue dcBranch(1).Text
    End If
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL, mFromDate, mToDate
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text, 2
        DBCboClientName2.BoundText = CUSTID
    End If
End Sub



  Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.Text = "" Then
            Me.DCboItemS.BoundText = ""
        Else
            Me.DCboItemS.BoundText = GetItemID(Trim$(Me.TxtItemCode.Text))
        End If
    End If
End Sub
Private Sub Form_Load()

    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
 
    ScreenNameArabic = "  ‹‹ﬁ‹‹«—Ì‹‹‹— «·«‰‰«Ã  "
    ScreenNameEnglish = "  Production Report "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    FromDate2.value = Date
    ToDate2.value = Date
    FromDate2.value = ""
    ToDate2.value = ""
    
    Dim StrSQL As String
 
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "select ItemID,ItemName from tblitems  where GroupID in ( "
    Else
        StrSQL = "select ItemID,ItemNamee from tblitems  where GroupID in ( "
    End If
    StrSQL = StrSQL & " SELECT     GroupID "
    StrSQL = StrSQL & " From dbo.Groups"
    StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"
    
    fill_combo DCboItemS, StrSQL
    'Dcombos.GetExpensesType XPCboExpensesType
    
     Dim i As Integer
     For i = 1 To 12
        cmbFromMonthName.AddItem MonthName(i)
        cmbToMonthName.AddItem MonthName(i)
     Next
     
     For i = 2008 To 2030
        cmbFromYear.AddItem i
        cmbToYear.AddItem i
     Next
     cmbFromMonthName.ListIndex = 0
     cmbToMonthName.ListIndex = 11
    cmbFromYear.Text = year(Date)
    cmbToYear.Text = year(Date)

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
            With carStatus
            .Clear
            .AddItem "€Ì— „Õœœ"
            .AddItem "»«·ÿ—Ìﬁ"
            .AddItem "»«·„Êﬁ⁄"
            .AddItem "›«—€"
            .AddItem "»«·Ê—‘…"
        End With
        With cmbTypeRep
            .Clear
            .AddItem "œ«Œ·Ï"
            .AddItem "·Êﬂ«·"
            .AddItem "Œ«—ÃÌ"
        End With
        With orderStatus
            .Clear
            .AddItem "„› ÊÕ"
            .AddItem " „"
            .AddItem "„€·ﬁ"
            .AddItem " √ŒÌ—"
        End With
    Else
        With orderStatus
            .Clear
            .AddItem "Open"
            .AddItem "Done"
            .AddItem "Closed"
            .AddItem "Delayed"
        End With
        With cmbTypeRep
            .Clear
            .AddItem "Internal"
            .AddItem "Local"
            .AddItem "External"
        End With
    End If
    
    LodR
    Dcombos.GetTypesTransport Me.DcbTypeTransport
    Dcombos.GetHarbors Me.DcbHarbor
    Dcombos.GetShips Me.DcbShip
    Dcombos.GetTblCarsDataGroup VehicleType(0)
    Dcombos.GetTblCarsDataGroup VehicleType(1)
    Dcombos.GetEmployees Me.Dcemp, , True
    Dcombos.GetCitiesDistance Me.DcCityFromId, 0
    Dcombos.GetCitiesDistance Me.DcCityToId, 1
    Dcombos.GetCars Me.Dccar
    Dcombos.GetEmployees Me.DcbEmployee(0), , True
    Dcombos.GetEmployees Me.DcbEmployee(1), , True
    Dcombos.GetBranches dcBranch(0)
    Dcombos.GetBranches dcBranch(1)
    Dcombos.GetCitiesDistance Me.DcCityFromId2, 0
    Dcombos.GetCitiesDistance Me.DcCityToId2, 1
    Dcombos.GetCars Me.DCCar2
    Dcombos.GetCarByVonder DCCar3
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName2
    Dcombos.GetEmployees Me.Dcemp, , True
    Dcombos.GetCars Me.DcbCar
    Dcombos.GetEmployees Me.DcEmpSuper
    
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName1

    Dcombos.GetCars Me.DcbCar
    
    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo
    FrmDate.value = Date
    todate.value = Date
    FrmDate.value = ""
    todate.value = ""
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub
 Sub LodR()
Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
   If SystemOptions.ShowDriverOnly = True Then
   str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   End If
    fill_combo DcEmployee2, str

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

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub TxtQtyDownLoad_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtQtyDownload.Text, 0)
End Sub

Private Sub TxtQtyUpload_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtQtyUpload.Text, 0)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub DcEmpSuper_Change()
    If val(DcEmpSuper.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcEmpSuper.BoundText, EmpCode
    Text3.Text = EmpCode
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcEmpSuper.BoundText = EmpID
    End If
End Sub

