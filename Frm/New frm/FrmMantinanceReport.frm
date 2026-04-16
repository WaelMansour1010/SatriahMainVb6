VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmMantinanceReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   Icon            =   "FrmMantinanceReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   873
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
   End
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   8070
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   7125
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11115
      Begin VB.CommandButton Command2 
         Caption         =   " ﬁ«—Ì— ⁄Âœ «·„ÊŸ›Ì‰"
         Height          =   435
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   6720
         TabIndex        =   4
         Top             =   120
         Width           =   4332
         Begin VB.Image Image1 
            Height          =   3672
            Left            =   0
            Picture         =   "FrmMantinanceReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4272
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1095
            Left            =   240
            TabIndex        =   5
            Top             =   3840
            Width           =   2895
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6855
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   6705
         _cx             =   11827
         _cy             =   12091
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
         Caption         =   " ﬁ—Ì— «·’Ì«‰…| ﬁ—Ì— «·„—ﬂ»« | ﬁ«—Ì— ’Ì«‰… 2| ﬁ«—Ì— «·Ê—œÌ« "
         Align           =   0
         CurrTab         =   1
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6435
            Left            =   45
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Width           =   6615
            _cx             =   11668
            _cy             =   11351
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
            Begin VB.CheckBox chkTafweed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ﬁ—Ì— «·„›Ê÷Ì‰"
               Height          =   285
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   5580
               Width           =   2445
            End
            Begin VB.Frame Frame8 
               Height          =   495
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   4170
               Width           =   3945
               Begin VB.OptionButton optAlll 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬂ·"
                  Height          =   285
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optAuthorization 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ›ÊÌ÷"
                  Height          =   285
                  Left            =   2910
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   180
                  Width           =   855
               End
               Begin VB.OptionButton optNotAuthorization 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·€«¡ «· ›ÊÌ÷"
                  Height          =   285
                  Left            =   1170
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   180
                  Width           =   1395
               End
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   2010
               Width           =   945
            End
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   375
               Left            =   4200
               TabIndex        =   58
               Top             =   5910
               Width           =   2175
               _Version        =   786432
               _ExtentX        =   3836
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«ŸÂ«— »Ì«‰«  ﬁ«∆œ «·„⁄œ…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.ComboBox DcbStuts 
               Height          =   315
               Left            =   357
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   1680
               Width           =   3375
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «‰ Â«¡ «·«” „«—…"
               Height          =   975
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   4710
               Width           =   5775
               Begin MSComCtl2.DTPicker FromDate 
                  Height          =   330
                  Left            =   2640
                  TabIndex        =   50
                  Top             =   270
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   51
                  Top             =   270
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin VB.PictureBox FromDateH 
                  BackColor       =   &H000000FF&
                  Height          =   1000
                  Left            =   -2040
                  ScaleHeight     =   945
                  ScaleWidth      =   945
                  TabIndex        =   54
                  Top             =   0
                  Width           =   1000
               End
               Begin VB.PictureBox ToDateH 
                  BackColor       =   &H000000FF&
                  Height          =   1000
                  Left            =   0
                  ScaleHeight     =   1005
                  ScaleWidth      =   45
                  TabIndex        =   55
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   7
                  Left            =   5010
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   300
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   6
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   330
                  Width           =   480
               End
            End
            Begin VB.TextBox txtModel 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   960
               Width           =   3372
            End
            Begin MSDataListLib.DataCombo DCGroup 
               Height          =   288
               Left            =   360
               TabIndex        =   14
               Top             =   120
               Width           =   3372
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmployee 
               Height          =   315
               Left            =   357
               TabIndex        =   15
               Top             =   570
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   495
               Index           =   1
               Left            =   360
               TabIndex        =   16
               Top             =   5670
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   873
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
            Begin MSDataListLib.DataCombo LocationID 
               Height          =   315
               Left            =   357
               TabIndex        =   47
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   1320
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEqup4 
               Height          =   315
               Left            =   360
               TabIndex        =   150
               Top             =   2010
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDept3 
               Height          =   315
               Left            =   360
               TabIndex        =   152
               Top             =   2340
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbKafelName 
               Height          =   315
               Left            =   360
               TabIndex        =   154
               Top             =   3465
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbCarModel 
               Bindings        =   "FrmMantinanceReport.frx":28E2
               Height          =   315
               Left            =   357
               TabIndex        =   156
               Top             =   3060
               Width           =   3375
               _ExtentX        =   5953
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
            Begin MSDataListLib.DataCombo DcbDepartment2 
               Height          =   315
               Left            =   360
               TabIndex        =   162
               Top             =   2700
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCOwner 
               Height          =   315
               Left            =   360
               TabIndex        =   165
               Top             =   3840
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lab 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„«·ﬂ"
               Height          =   360
               Index           =   0
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   3840
               Width           =   1455
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁ”„"
               Height          =   225
               Index           =   5
               Left            =   4260
               TabIndex        =   163
               Top             =   2760
               Width           =   885
            End
            Begin VB.Label lblModel 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·ÿ—«“ "
               Height          =   255
               Left            =   4305
               TabIndex        =   157
               Top             =   3120
               Width           =   855
            End
            Begin VB.Label Lab 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’«Õ» «·⁄„·"
               Height          =   360
               Index           =   19
               Left            =   3705
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   3525
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«œ«—…"
               Height          =   285
               Index           =   41
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   2430
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„⁄œ…"
               Height          =   285
               Index           =   40
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   2010
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·…«·„⁄œ…"
               Height          =   195
               Index           =   47
               Left            =   4470
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1680
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ê«ﬁ⁄ «·⁄„· "
               Height          =   348
               Index           =   27
               Left            =   3864
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1320
               Width           =   1296
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„⁄œ…"
               Height          =   264
               Index           =   103
               Left            =   3852
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   120
               Width           =   1344
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ«∆œ «·„⁄œ…"
               Height          =   312
               Index           =   104
               Left            =   3852
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   564
               Width           =   1344
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÊœÌ·"
               Height          =   348
               Index           =   107
               Left            =   3924
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   960
               Width           =   1296
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   6435
            Index           =   2
            Left            =   -7260
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   45
            Width           =   6615
            _cx             =   11668
            _cy             =   11351
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
            Begin VB.OptionButton optCover 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Õ÷— Ã—œ «·ﬂ›—« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   124
               ToolTipText     =   "«’€— „‰"
               Top             =   420
               Width           =   2175
            End
            Begin VB.OptionButton optAlarm 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ‰»ÌÂ«  «·’Ì«‰…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   -390
               RightToLeft     =   -1  'True
               TabIndex        =   107
               ToolTipText     =   "«’€— „‰"
               Top             =   450
               Width           =   1875
            End
            Begin VB.OptionButton ChCarExpen 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„’—Ê›«  «·„—ﬂ»…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   42
               ToolTipText     =   "«’€— „‰"
               Top             =   0
               Width           =   1755
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   3585
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   720
               Width           =   5775
               Begin VB.ComboBox cmbCoverStatus 
                  Height          =   315
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   2340
                  Width           =   1575
               End
               Begin VB.CheckBox CHKiSoUT 
                  Caption         =   "«·„’—Ê›«  «· Ï ·Â« ”‰œ«  ’—›"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   3150
                  Width           =   2595
               End
               Begin VB.TextBox Text6 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3330
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   960
                  Width           =   975
               End
               Begin VB.TextBox TXTOrderMaintID2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   3120
                  Width           =   1335
               End
               Begin VB.ComboBox DcbStutsMaint2 
                  Height          =   315
                  Left            =   2700
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   2370
                  Width           =   1575
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3300
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   1320
                  Width           =   1005
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   1680
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DcbBranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   4170
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbDept 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   32
                  Top             =   600
                  Width           =   4170
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEqup 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   33
                  Top             =   960
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbTypeMain 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   34
                  Top             =   1320
                  Width           =   3180
                  _ExtentX        =   5609
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmp 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1680
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbLeader 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   44
                  Top             =   2040
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCGroups3 
                  Height          =   315
                  Left            =   1410
                  TabIndex        =   125
                  Top             =   2700
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·ﬂ›—« "
                  Height          =   285
                  Index           =   35
                  Left            =   1500
                  TabIndex        =   146
                  Top             =   2370
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "·√„— „Õœœ"
                  Height          =   285
                  Index           =   33
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   3120
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ã„Ê⁄Â"
                  Height          =   315
                  Index           =   29
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   2700
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’Ì«‰…"
                  Height          =   285
                  Index           =   28
                  Left            =   4560
                  TabIndex        =   121
                  Top             =   2370
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁ«∆œ «·„⁄œ…"
                  Height          =   285
                  Index           =   5
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„”∆Ê· «·’Ì«‰Â"
                  Height          =   285
                  Index           =   2
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·’Ì«‰Â"
                  Height          =   285
                  Index           =   0
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "·›—⁄ „⁄Ì‰"
                  Height          =   285
                  Index           =   38
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "·«œ«—Â/·ﬁ”„ „⁄Ì‰"
                  Height          =   285
                  Index           =   37
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„⁄œ…"
                  Height          =   285
                  Index           =   39
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   960
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·› —Â"
               Height          =   1665
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   4290
               Width           =   5775
               Begin VB.Frame Frame7 
                  Caption         =   "«· «—ÌŒ"
                  Height          =   795
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   540
                  Width           =   4305
                  Begin VB.ComboBox DcbStutsMaint4 
                     Height          =   315
                     Left            =   60
                     RightToLeft     =   -1  'True
                     TabIndex        =   147
                     Top             =   450
                     Width           =   1575
                  End
                  Begin VB.OptionButton optBydate 
                     Caption         =   "ÿ»ﬁ« ··Œÿ…"
                     Height          =   225
                     Index           =   1
                     Left            =   120
                     TabIndex        =   110
                     Top             =   180
                     Value           =   -1  'True
                     Width           =   1605
                  End
                  Begin VB.OptionButton optBydate 
                     Caption         =   "ÿ»ﬁ« ··ÿ·»"
                     Height          =   225
                     Index           =   0
                     Left            =   2280
                     TabIndex        =   109
                     Top             =   180
                     Width           =   1605
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ«·…"
                     Height          =   285
                     Index           =   36
                     Left            =   1920
                     TabIndex        =   148
                     Top             =   450
                     Width           =   1125
                  End
               End
               Begin MSComCtl2.DTPicker DtpDateFrom 
                  Height          =   330
                  Left            =   2640
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DtpDateTo 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin XtremeSuiteControls.CheckBox chkHidden 
                  Height          =   375
                  Left            =   690
                  TabIndex        =   137
                  Top             =   1260
                  Width           =   3585
                  _Version        =   786432
                  _ExtentX        =   6324
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "«Œ›«¡ ÿ·»«  «·’Ì«‰… «· Ï  „ «·«‰ Â«¡ „‰Â«"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   3
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   240
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   4
                  Left            =   5010
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.OptionButton ChorderAnlys 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„— ‘€·  Õ·Ì·Ì"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   -240
               RightToLeft     =   -1  'True
               TabIndex        =   23
               ToolTipText     =   "«’€— „‰"
               Top             =   0
               Width           =   1755
            End
            Begin VB.OptionButton Chorder 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ê«„— «·‘€·"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   22
               ToolTipText     =   "«’€— „‰"
               Top             =   0
               Width           =   1395
            End
            Begin VB.OptionButton CHReq 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ·»«  «·’Ì«‰Â"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   21
               ToolTipText     =   "«’€— „‰"
               Top             =   0
               Width           =   1875
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   405
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   6000
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   714
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
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   6435
            Index           =   0
            Left            =   7350
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   45
            Width           =   6615
            _cx             =   11668
            _cy             =   11351
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
            Begin VB.OptionButton Order2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ «»⁄Â «·«Ê«„—"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   119
               ToolTipText     =   "«’€— „‰"
               Top             =   -120
               Width           =   1515
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·› —Â"
               Height          =   615
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   5490
               Width           =   5775
               Begin MSComCtl2.DTPicker DtpDateFrom2 
                  Height          =   330
                  Left            =   2640
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DtpDateTo2 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   80412675
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   15
                  Left            =   5010
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   14
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   240
                  Width           =   480
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   5055
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   450
               Width           =   5775
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3330
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.TextBox TxtSearchCode4 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.TextBox TxtSearchCode3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.ComboBox DcbStutsMaint 
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   3480
                  Width           =   1575
               End
               Begin VB.ComboBox DcbStatusMaint 
                  Height          =   315
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   3840
                  Width           =   1335
               End
               Begin VB.ComboBox EquipmentStatusid 
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   3840
                  Width           =   1575
               End
               Begin VB.TextBox TXTOrderMaintID 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.TextBox TxtSearchCode2 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.TextBox Text12 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   2760
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DcbBranch2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   62
                  Top             =   240
                  Width           =   4170
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbDept2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   63
                  Top             =   600
                  Width           =   4170
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEqup2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   65
                  Top             =   960
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbTypeMain2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   67
                  Top             =   1320
                  Width           =   3150
                  _ExtentX        =   5556
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbEmp2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1680
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbLeader2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   77
                  Top             =   2760
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbGroup 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   78
                  Top             =   3120
                  Width           =   4170
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcsupervisor 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   71
                  Top             =   2040
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dctechnical 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   73
                  Top             =   2400
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboUserName 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   102
                  Top             =   4245
                  Width           =   4140
                  _ExtentX        =   7303
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCMaintenanceTypes 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   105
                  Top             =   4560
                  Width           =   4140
                  _ExtentX        =   7303
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "›∆… «·’Ì«‰Â"
                  Height          =   285
                  Index           =   22
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   4680
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„” Œœ„ „Õœœ"
                  Height          =   270
                  Index           =   21
                  Left            =   4800
                  TabIndex        =   103
                  Top             =   4200
                  Width           =   900
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "›‰Ì „Õœœ"
                  Height          =   285
                  Index           =   19
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   2400
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‘—› «·ﬁ”„"
                  Height          =   285
                  Index           =   18
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’Ì«‰…"
                  Height          =   285
                  Index           =   63
                  Left            =   1650
                  TabIndex        =   99
                  Top             =   3480
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·ÿ·»"
                  Height          =   285
                  Index           =   20
                  Left            =   4560
                  TabIndex        =   98
                  Top             =   3840
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·Â «·„⁄œ…"
                  Height          =   285
                  Index           =   26
                  Left            =   1680
                  TabIndex        =   96
                  Top             =   3840
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "·√„— „Õœœ"
                  Height          =   285
                  Index           =   17
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Ã„Ê⁄…"
                  Height          =   285
                  Index           =   16
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   3120
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„⁄œ…"
                  Height          =   285
                  Index           =   13
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«œ«—…"
                  Height          =   285
                  Index           =   12
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "·›—⁄ „⁄Ì‰"
                  Height          =   285
                  Index           =   11
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·’Ì«‰Â"
                  Height          =   285
                  Index           =   10
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„”∆Ê· «·’Ì«‰Â"
                  Height          =   285
                  Index           =   9
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁ«∆œ «·„⁄œ…"
                  Height          =   285
                  Index           =   8
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2760
                  Width           =   1335
               End
            End
            Begin ImpulseButton.ISButton CmdPrint 
               Height          =   375
               Left            =   120
               TabIndex        =   92
               Top             =   6000
               Width           =   1365
               _ExtentX        =   2408
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
            Begin VB.OptionButton Emp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·”«∆ﬁ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   60
               ToolTipText     =   "«’€— „‰"
               Top             =   -120
               Width           =   1515
            End
            Begin VB.OptionButton Order 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ê«„— «·‘€·"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   90
               ToolTipText     =   "«’€— „‰"
               Top             =   -120
               Width           =   1515
            End
            Begin VB.OptionButton ReqMaint 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ·»«  «·’Ì«‰Â"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   91
               ToolTipText     =   "«’€— „‰"
               Top             =   -120
               Width           =   2115
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   6435
            Index           =   1
            Left            =   7650
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   45
            Width           =   6615
            _cx             =   11668
            _cy             =   11351
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
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   2040
               Width           =   1215
            End
            Begin VB.ComboBox DcbStutsMaint3 
               Height          =   315
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   1350
               Width           =   1575
            End
            Begin VB.ComboBox DcbStuts2 
               Height          =   315
               Left            =   2340
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   990
               Width           =   2505
            End
            Begin VB.CommandButton Command1 
               Caption         =   "⁄—÷ «· ﬁ—Ì—"
               Height          =   495
               Left            =   1890
               TabIndex        =   123
               Top             =   3480
               Width           =   1875
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "“„‰ «·’Ì«‰…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   -420
               RightToLeft     =   -1  'True
               TabIndex        =   115
               ToolTipText     =   "«’€— „‰"
               Top             =   450
               Width           =   2115
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ”·Ì„ Ê—œÌ…  ›’Ì·Ì"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   114
               ToolTipText     =   "«’€— „‰"
               Top             =   450
               Width           =   2115
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«” ·«„ Ê—œÌ…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   480
               Left            =   4170
               RightToLeft     =   -1  'True
               TabIndex        =   113
               ToolTipText     =   "«’€— „‰"
               Top             =   450
               Value           =   -1  'True
               Width           =   2115
            End
            Begin MSComCtl2.DTPicker txtToDate 
               Height          =   330
               Left            =   570
               TabIndex        =   112
               Top             =   3000
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   80412675
               CurrentDate     =   41640
            End
            Begin ImpulseButton.ISButton ISButton1 
               Default         =   -1  'True
               Height          =   375
               Left            =   5610
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   6270
               Visible         =   0   'False
               Width           =   1365
               _ExtentX        =   2408
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
            Begin MSComCtl2.DTPicker txtFromDate 
               Height          =   330
               Left            =   4080
               TabIndex        =   116
               Top             =   2940
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   80412675
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo cmbShiftMaintType 
               Height          =   315
               Left            =   2370
               TabIndex        =   131
               Top             =   1740
               Width           =   2490
               _ExtentX        =   4392
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkStatusCar 
               Height          =   375
               Left            =   4230
               TabIndex        =   133
               Top             =   6360
               Visible         =   0   'False
               Width           =   1725
               _Version        =   786432
               _ExtentX        =   3043
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "»„ÊÃ» Õ«·… «·„⁄œÂ/«·”Ì«—…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkStatusMaint 
               Height          =   375
               Left            =   2460
               TabIndex        =   134
               Top             =   6360
               Visible         =   0   'False
               Width           =   1725
               _Version        =   786432
               _ExtentX        =   3043
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "»„ÊÃ» Õ«·… «·’Ì«‰…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkShift 
               Height          =   375
               Left            =   1170
               TabIndex        =   135
               Top             =   6330
               Visible         =   0   'False
               Width           =   1275
               _Version        =   786432
               _ExtentX        =   2249
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "»„ÊÃ» «·Ê—œÌ…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox chkAll 
               Height          =   375
               Left            =   -570
               TabIndex        =   136
               Top             =   6330
               Visible         =   0   'False
               Width           =   1725
               _Version        =   786432
               _ExtentX        =   3043
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«·ﬂ·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEqup3 
               Height          =   315
               Left            =   690
               TabIndex        =   140
               Top             =   2070
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„⁄œ…"
               Height          =   285
               Index           =   34
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   2070
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ê—œÌ…"
               Height          =   270
               Index           =   32
               Left            =   5430
               TabIndex        =   132
               Top             =   1710
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’Ì«‰…"
               Height          =   285
               Index           =   31
               Left            =   5070
               TabIndex        =   130
               Top             =   1320
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·…«·„⁄œ…"
               Height          =   195
               Index           =   30
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   990
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   23
               Left            =   5460
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   3000
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   24
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   3000
               Width           =   480
            End
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "‘«‘…  ﬁ«—Ì— ÿ·»«  «·’Ì«‰Â Ê «Ê«„— «·‘€·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Index           =   25
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   5400
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   495
         Left            =   6840
         Top             =   5400
         Width           =   3975
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   8
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— ÿ·»«  «·’Ì«‰Â Ê √Ê«„— «·‘€·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   -90
      TabIndex        =   3
      Top             =   0
      Width           =   11265
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmMantinanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    lbl(16).Caption = "Group"
    ChCarExpen.Caption = "Expenses"
    lbl(5).Caption = "Driver"
    lbl(8).Caption = "Driver"
    Command2.Caption = "Reports Covenant"
    lbl(27).Caption = "Location"
    lbl(47).Caption = "Status"
    Frame2.Caption = "Date of License"
    lbl(7).Caption = "From"
    lbl(6).Caption = "To"
    CheckBox1.Caption = "Show Data of Employee"
   ' Set XPic = Me.btnFirst.ButtonImage
   ' Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
   ' Set Me.btnLast.ButtonImage = XPic
   ' Set XPic = Me.btnPrevious.ButtonImage
   ' Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
   ' Set Me.btnNext.ButtonImage = XPic
   CmdPrint.Caption = "View Report"
   lbl(14).Caption = "To"
   lbl(15).Caption = "From"
   Frame5.Caption = "Period"
    Label5.Caption = "Reports of Maintenance Requests and Orders"
    lbl(25).Caption = Label5.Caption
  CHReq.RightToLeft = False
  CHReq.Caption = "Maintenance Requests"
  Chorder.RightToLeft = False
   Me.ReqMaint.Caption = "Maintenance Requests"
  ReqMaint.RightToLeft = False
  Chorder.Caption = "Maintenance Orders"
  ChorderAnlys.Caption = "Orders Analytic"
  Emp.RightToLeft = False
  Emp.Caption = "Drivers "
  Order.Caption = "Maintenance Orders"
  Order.RightToLeft = False
  ChorderAnlys.Caption = "Orders Analytic"
  ChorderAnlys.RightToLeft = False
  lbl(38).Caption = "Branch"
  lbl(37).Caption = "Department"
   lbl(39).Caption = "Equipment"
   lbl(13).Caption = "Equipment"
   lbl(0).Caption = "Type Maintenance"
     lbl(2).Caption = "Maint. Manager"
     Frame1.Caption = "Priod"
     lbl(3).Caption = "To"
     lbl(4).Caption = "From"
     lbl(10).Caption = "Type Maintenance "
     lbl(9).Caption = "Maint. Manager"
btnClear.Caption = "Clear"
lbl(11).Caption = "Branch"
lbl(12).Caption = "Management"
Cmd(0).Caption = "Show Report"
Cmd(2).Caption = "Exit"
lblCompanyname.Caption = "AL SATTARYAH"
   
   lbl(103).Caption = " Vehicle Type "
   lbl(104).Caption = " Vehicle Driver "
 lbl(107).Caption = "Model"
 Cmd(1).Caption = " View Report "
 
 C1Tab1.Caption = "Maintainance | Vehicle|Maintainance2"
 
End Sub
Private Sub btnClear_Click()
clear_all Me
DcbTypeMain.Enabled = False
TxtSearchCode.Enabled = False
DcbEmp.Enabled = False
DcbDept.Enabled = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.FromDate.value = ""
Me.ToDate.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
End Sub

Private Sub ChCarExpen_Click()
CHKiSoUT.Enabled = True
TxtSearchCode.Enabled = True
DcbEmp.Enabled = True
DcbTypeMain.Enabled = True
Text3.Enabled = True
Frame7.Visible = False
End Sub

Private Sub Chorder_Click()
CHKiSoUT.Enabled = False
DcbTypeMain.Enabled = True
TxtSearchCode.Enabled = True
DcbEmp.Enabled = True
DcbDept.Enabled = False
Frame7.Visible = False
End Sub

Private Sub ChorderAnlys_Click()
CHKiSoUT.Enabled = False
Frame7.Visible = False
End Sub

Private Sub CHReq_Click()
CHKiSoUT.Enabled = False
DcbTypeMain.Enabled = False
TxtSearchCode.Enabled = False
DcbEmp.Enabled = False
DcbDept.Enabled = True
Frame7.Visible = False
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            If Me.Chorder.value = True Or Me.CHReq.value = True Or ChorderAnlys.value = True Then
                GetData
            ElseIf Me.ChCarExpen.value = True Then
                GetData1
            ElseIf optAlarm Then
                GetDataAlarm
            ElseIf optCover Then
                GetDataCover
            End If
    
        Case 1
            print_report2
    
        Case 2
            Unload Me
    Case 3
'print_report
    End Select

End Sub




Private Sub CmdPrint_Click()
    If ReqMaint.value = True Or Emp.value = True Or Order.value = True Or Order2.value = True Then
        GetData2
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
        Else
            MsgBox "Please select type of reports"
        End If
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
print_report4
End Sub

Private Sub Command2_Click()
 FixedAssetReportsEmp.show
End Sub

Private Sub DataCombo3_Click(Area As Integer)

End Sub

Private Sub DcbDept3_Change()
LoadDept
End Sub

  Sub LoadDept()
     Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcbDepartment2
     Dcombos.GetNewDwpartMent DcbDepartment2, True, val(DcbDept3.BoundText)
  End Sub

Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub

Private Sub DcbEmp_Click(Area As Integer)
    If val(DcbEmp.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
'
    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode

End Sub

Private Sub DcbEmp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype.Caption = 11020171
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcbEmp2_Change()
DcbEmp2_Click (0)
End Sub

Private Sub DcbEmp2_Click(Area As Integer)
    If val(DcbEmp2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp2.BoundText, EmpCode
    TxtSearchCode2.Text = EmpCode
End Sub

Private Sub DcbEmp2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype.Caption = 11020174
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcbEqup_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmMantinanceReport"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbEqup2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmMantinanceReport2"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbLeader_Change()
DcbLeader_Click (0)
End Sub

Private Sub DcbLeader_Click(Area As Integer)
  If val(DcbLeader.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbLeader.BoundText, EmpCode
    Text1.Text = EmpCode
End Sub

Private Sub DcbLeader_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype.Caption = 1102017
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcbLeader2_Change()
DcbLeader2_Click (0)
End Sub

Private Sub DcbLeader2_Click(Area As Integer)
  If val(DcbLeader2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbLeader2.BoundText, EmpCode
    Text12.Text = EmpCode
End Sub

Private Sub DcbLeader2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype.Caption = 11020172
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype.Caption = 11020173
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DCGroup_Change()
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DCGroup.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DCGroup.BoundText)
   End If
End Sub

Private Sub DCGroup_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DCGroup.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.DCGroup.BoundText)
   End If
End Sub


Private Sub dcsupervisor_Change()
dcsupervisor_Click (0)
End Sub

Private Sub dcsupervisor_Click(Area As Integer)
If val(dcsupervisor.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , dcsupervisor.BoundText, EmpCode
    TxtSearchCode3.Text = EmpCode
    
End Sub

Private Sub dctechnical_Change()
dctechnical_Click (0)
End Sub

Private Sub dctechnical_Click(Area As Integer)
If val(dctechnical.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , dctechnical.BoundText, EmpCode
    TxtSearchCode4.Text = EmpCode
    
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub



Private Sub Fromdateh_LostFocus()
 'FromDate.value = ToGregorianDate(FromDateH.value)
End Sub

Private Sub optAlarm_Click()
CHKiSoUT.Enabled = False
If optAlarm Then
    Frame7.Visible = True
    optBydate(1).value = True
Else
    Frame7.Visible = False
End If
End Sub

Private Sub optCover_Click()
CHKiSoUT.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        Me.DcbLeader.BoundText = EmpID
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text12.Text, EmpID
        Me.DcbLeader2.BoundText = EmpID
    End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    Dim rsDummy As New ADODB.Recordset
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID =  " & user_id & "))) AND  dbo.TblCarsData.Fullcode like '%" & Text2.Text & "%'  "

   
   rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
   Dcombos.GetEquipments DcbEqup2, str
   If Not rsDummy.EOF Then
    DcbEqup2.BoundText = val(rsDummy!FixedassetId)
   End If
    


End Sub

Private Sub Text3_Change()
'    StrSQL = "select * from TblMaintenanceType  "
'    StrSQL = StrSQL & " where ( MainType =0 or MainType is null) "
'    If Trim(Text3) <> "" Then
'        StrSQL = StrSQL & " and( name like '%" & (Text3) & "%'  or namee like '%" & (Text3) & "%' )"
'    ' StrSQL = StrSQL & "  and  name=" & (.TextMatrix(Row, .ColIndex("QuickSearch"))) & "   "
'    End If
'
    Dim Dcombos As New ClsDataCombos
    
    Dcombos.GetQuicSearch DcbTypeMain, Text3, "TblMaintenanceType"
    
    


End Sub

Private Sub Text4_Change()
    Dim Dcombos As New ClsDataCombos
    
    Dcombos.GetQuicSearch DcbTypeMain2, Text4, "TblMaintenanceType"
    

End Sub

Private Sub Text5_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    Dim rsDummy As New ADODB.Recordset
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & user_id & " ))) AND  dbo.TblCarsData.Fullcode like '%" & Text5.Text & "%'  "

   
   rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
   Dcombos.GetEquipments DcbEqup3, str
   If Not rsDummy.EOF Then
        DcbEqup3.BoundText = val(rsDummy!FixedassetId)
   End If
    

End Sub

Private Sub Text6_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    Dim rsDummy As New ADODB.Recordset
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & user_id & " ))) AND  dbo.TblCarsData.Fullcode like '%" & Text6.Text & "%'  "

   
   rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
   Dcombos.GetEquipments DcbEqup, str
   If Not rsDummy.EOF Then
        DcbEqup.BoundText = val(rsDummy!FixedassetId)
   End If
    
End Sub

Private Sub Text7_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    Dim rsDummy As New ADODB.Recordset
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID =  " & user_id & "))) AND  dbo.TblCarsData.Fullcode like '%" & Text7.Text & "%'  "

   
   rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
   Dcombos.GetEquipments DcbEqup4, str
   If Not rsDummy.EOF Then
    DcbEqup4.BoundText = val(rsDummy!FixedassetId)
   End If
    
End Sub

Private Sub ToDate_Change()
If Not IsNull(ToDate.value) Then
 'ToDateH.value = ToHijriDate(ToDate.value)
 End If
End Sub

Private Sub ToDateH_LostFocus()
'ToDate.value = ToGregorianDate(ToDateH.value)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        Me.DcbEmp.BoundText = EmpID
    End If

End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim I As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Dim str As String

       EquipmentStatusid.AddItem " ⁄„·"
            EquipmentStatusid.AddItem "„ Êﬁ›…"
            EquipmentStatusid.AddItem "⁄ÿ· ÿ—Ìﬁ"
            EquipmentStatusid.AddItem "«·ﬂ·"
            
            With DcbStatusMaint
.Clear
.AddItem "„› ÊÕ"
.AddItem "„—›Ê÷"
.AddItem " „ «·«‰ Â«¡"
 .AddItem " Õ  «· ‰›Ì–"
  .AddItem "«·ﬂ·"
End With

With cmbCoverStatus
    .Clear
    .AddItem "ÃœÌœ"
    .AddItem "„” ⁄„·"
    .AddItem " «·›"
End With



    With DcbStutsMaint4
    .Clear
    .AddItem " Õ  «· ‰›Ì–"
    .AddItem " „ «· ‰›Ì–"
    .AddItem "„—›Ê÷"
End With
    

            With DcbStutsMaint
.Clear
.AddItem "Ã«—Ì «·«’·«Õ"
.AddItem "Ã«Â“"
.AddItem "Œ—Ã"
   .AddItem "«·ﬂ·"
End With


            With DcbStutsMaint3
.Clear
.AddItem "Ã«—Ì «·«’·«Õ"
.AddItem "Ã«Â“"
.AddItem "Œ—Ã"
   .AddItem "«·ﬂ·"
End With


With DcbStutsMaint2
.Clear
.AddItem "Ã«—Ì «·«’·«Õ"
.AddItem "Ã«Â“"
.AddItem "Œ—Ã"
   .AddItem "«·ﬂ·"
End With

Frame7.Visible = False
    If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.ShiftMaintType.ID, dbo.ShiftMaintType.Name"
    Else
      str = " SELECT     dbo.ShiftMaintType.ID, dbo.ShiftMaintType.Namee"
    End If
    str = str & " From ShiftMaintType"
    fill_combo cmbShiftMaintType, str

If SystemOptions.UserInterface = ArabicInterface Then


 With DcbStuts2
    .Clear
    .AddItem "›Ï «·Ê—‘…"
    .AddItem "⁄ÿ· ›Ï «·ÿ—Ìﬁ"
    .AddItem "›Ï «·»«—ﬂÌ‰Ã"
    .AddItem "«·ﬂ·"
End With
Else
With DcbStuts
.Clear
.AddItem "Active"
.AddItem "Under Maintenance "
.AddItem "Sold"
End With
End If


If SystemOptions.UserInterface = ArabicInterface Then
With DcbStuts
.Clear
.AddItem "‰‘ÿ"
.AddItem " Õ  «·’Ì«‰…"
.AddItem "„»«⁄"
End With
Else
With DcbStuts
.Clear
.AddItem "Active"
.AddItem "Under Maintenance "
.AddItem "Sold"
End With
End If
    
    Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches DcbBranch
   Dcombos.GetEmpDepartments DcbDept
   Dcombos.GetEmpDepartments DcbDept3
   Dcombos.GetTblCarModels Me.DcbCarModel
   Dcombos.GetBranches DcbBranch2
   Dcombos.GetEmpDepartments DcbDept2
   Dcombos.GetNewDwpartMent DcbDepartment2
   Dcombos.GetEmployees DcbEmp
   Dcombos.GetEmployees DcbEmp2
   Dcombos.GetItemSGroups Me.DCGroups3, False
    Dcombos.GetEmployees Me.dcsupervisor
     Dcombos.GetEmployees dctechnical
   Dcombos.GetmaintennceType DcbTypeMain
   Dcombos.GetmaintennceType DcbTypeMain2
   Dcombos.GetEquipments DcbEqup
   Dcombos.GetEquipments DcbEqup4
   Dcombos.GetEquipments DcbEqup3
   Dcombos.GetEquipments DcbEqup2
   Dcombos.GetEmpLocations LocationID
   Dcombos.GetTblCarsDataGroup DCGroup
   Dcombos.GetItemSGroups DcbGroup
     Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetCarsMaintenanceTypes Me.DCMaintenanceTypes, , , 1
     
    
    str = "SELECT DISTINCT KafelName, KafelName AS KafelNames"
    str = str & " From dbo.TblEmployee"
    str = str & " WHERE     (NOT (KafelName IS NULL)) "
    fill_combo DcbKafelName, str
       
     
     
         str = "select  distinct OwnerName,OwnerName from TblCarsData where not( OwnerName='')"
   fill_combo DCOwner, str
    
     txtFromDate = Date
     txtToDate = Date
     
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
  fill_combo DcbLeader2, str
  fill_combo DcbLeader, str
  fill_combo DcEmployee, str
 DtpDateFrom.value = Date
DtpDateTo.value = Date
DtpDateFrom2.value = Date
DtpDateTo2.value = Date
Me.FromDate.value = Date
Me.ToDate.value = Date
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
Me.FromDate.value = ""
Me.ToDate.value = ""
DcbTypeMain.Enabled = False
TxtSearchCode.Enabled = False
DcbEmp.Enabled = False
DcbDept.Enabled = False
    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
    If Me.Chorder.value = True Or Me.CHReq.value = True Or ChorderAnlys.value = True Then
If CHReq.value = True Then
StrSQL = " SELECT     dbo.TblRequerMainten.ID, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id, dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime,"
StrSQL = StrSQL & "                      dbo.TblRequerMainten.StartTime, dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate,"
StrSQL = StrSQL & "                      dbo.TblRequerMainten.StopDate, dbo.TblRequerMainten.EquepID, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.namee,"
StrSQL = StrSQL & "                      dbo.FixedAssets.id AS EqID, dbo.TblRequerMaintenDet.PartID, FixedAssets_1.code AS Partcode, FixedAssets_1.Name AS PartName,"
StrSQL = StrSQL & "                      FixedAssets_1.namee AS PartNameE"
StrSQL = StrSQL & " FROM         dbo.FixedAssets FixedAssets_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRequerMaintenDet ON FixedAssets_1.id = dbo.TblRequerMaintenDet.PartID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRequerMainten ON dbo.TblRequerMaintenDet.ReqID = dbo.TblRequerMainten.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " where 1=1"
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.TblRequerMainten.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
 If Me.DcbDept.Text <> "" And val(DcbDept.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblEmpDepartments.DeparmentID = " & val(Me.DcbDept.BoundText)

End If
If Trim(TXTOrderMaintID2) <> "" Then
    StrSQL = StrSQL & " AND dbo.TblRequerMainten.ID=" & val(TXTOrderMaintID2)
End If
'TXTOrderMaintID2
End If

If Me.Chorder.value = True Or ChorderAnlys.value = True Then
StrSQL = " SELECT     dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate, dbo.TblOrderMaint.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id, dbo.TblOrderMaint.startmaintenanceTime, dbo.TblOrderMaint.endmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceTime,"
StrSQL = StrSQL & "                      dbo.TblOrderMaint.endmaintenanceDate, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks, dbo.TblOrderMaint.Remarks,"
StrSQL = StrSQL & "                      dbo.TblOrderMaint.Des, dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.EquepID, dbo.FixedAssets.id AS EqID, dbo.FixedAssets.code, dbo.FixedAssets.Name,"
StrSQL = StrSQL & "                      dbo.FixedAssets.namee, dbo.TblOrderMaint.SuperVisor, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Jiha, dbo.TblOrderMaint.ended,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.Qty, dbo.tblordermaintenancetypes.Remarks AS RemarksDet, dbo.tblordermaintenancetypes.maintenanceid,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceType.name AS nameDet, dbo.TblMaintenanceType.namee AS nameDetE, dbo.tblordermaintenancetypes.TypeTrans,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.Transaction_ID, dbo.tblordermaintenancetypes.Transaction_IDDet, dbo.tblordermaintenancetypes.PartName,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.CusMobile, dbo.tblordermaintenancetypes.CusName, dbo.tblordermaintenancetypes.BillNo, dbo.tblordermaintenancetypes.Total,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.Price, dbo.tblordermaintenancetypes.Company, dbo.TblOrderMaint.EnterDate, dbo.TblOrderMaint.EquepmentName,"
StrSQL = StrSQL & "                      dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.LeaderName, dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName, dbo.TblOrderMaint.Total AS HExpr1,"
StrSQL = StrSQL & "                      dbo.TblOrderMaint.StutsMaint, dbo.TblOrderMaint.TechNote, dbo.TblOrderMaint.reciverid, dbo.TblOrderMaint.LeaderID, TblEmployee_1.Emp_Name,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3, TblEmployee_1.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.SuperID, TblEmployee_1.Emp_Name AS SuperEmp_Name, TblEmployee_1.Fullcode AS SuperFullcode,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee AS SuperEmp_NameE, dbo.tblordermaintenancetypes.EmpID, TblEmployee_2.Emp_Name AS FiterEmp_Name,"
StrSQL = StrSQL & "                      TblEmployee_2.Fullcode AS FiterFullcode, TblEmployee_2.Emp_Namee AS FiterEmp_NameE, dbo.tblordermaintenancetypes.PartID,"
StrSQL = StrSQL & "                      FixedAssets_1.code AS PartCodeDet, FixedAssets_1.Name AS PartNameDet, FixedAssets_1.namee AS PartNameEDet"
StrSQL = StrSQL & " FROM         dbo.TblMaintenanceType RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.FixedAssets FixedAssets_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes ON FixedAssets_1.id = dbo.tblordermaintenancetypes.PartID ON"
StrSQL = StrSQL & "                      dbo.TblMaintenanceType.id = dbo.tblordermaintenancetypes.maintenanceid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.tblordermaintenancetypes.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.tblordermaintenancetypes.SuperID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.tblordermaintenancetypes.DeptID = dbo.TblEmpDepartments.DeparmentID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOrderMaint LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_3 ON dbo.TblOrderMaint.LeaderID = TblEmployee_3.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.ORderID = dbo.TblOrderMaint.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.FixedAssets ON dbo.TblOrderMaint.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblOrderMaint.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " where 1=1"

If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.SuperVisor = " & val(Me.DcbEmp.BoundText)
End If
If Me.DcbTypeMain.Text <> "" And val(Me.DcbTypeMain.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.tblordermaintenancetypes.maintenanceid = " & val(Me.DcbTypeMain.BoundText)
End If
If Me.DcbLeader.Text <> "" And val(DcbLeader.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.LeaderID = " & val(Me.DcbLeader.BoundText)
End If
     If Me.TXTOrderMaintID2 <> "" Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.id  in ( " & Me.TXTOrderMaintID2.Text & ")"
        End If
   If Me.DcbStutsMaint2.ListIndex <> -1 And Me.DcbStutsMaint2.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.StutsMaint = " & val(Me.DcbStutsMaint2.ListIndex)
        End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
End If
If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)

End If

If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.FixedAssets.id = " & val(Me.DcbEqup.BoundText)

End If


 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If
End If
End Sub
Public Sub GetData2()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
    
    On Error GoTo ErrTrap

    If ReqMaint.value = True Then
        StrSQL = " SELECT StatusMaint, EquipmentStatusid, dbo.TblRequerMainten.ID AS RequerMainteniD, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
        StrSQL = StrSQL & " dbo.TblBranchesData.branch_id, dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName,"
        StrSQL = StrSQL & " dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime,"
        StrSQL = StrSQL & " dbo.TblRequerMainten.StartTime, dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate,"
        StrSQL = StrSQL & " dbo.TblRequerMainten.StopDate, dbo.TblRequerMainten.Mobile, dbo.TblRequerMainten.BoardNO, dbo.TblRequerMainten.OperationNo,"
        StrSQL = StrSQL & " dbo.TblRequerMainten.LeaderType, dbo.TblRequerMainten.LeaderName, dbo.TblRequerMainten.LeaderID, dbo.TblEmployee.Emp_Name,"
        StrSQL = StrSQL & " dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode,"
        StrSQL = StrSQL & " dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4,"
        StrSQL = StrSQL & " dbo.TblEmployee.NationlID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TblRequerMainten.EquepID, dbo.FixedAssets.Name AS EquipName,"
        StrSQL = StrSQL & " dbo.FixedAssets.namee AS EquipNameE, dbo.TblEmployee.Emp_ID ,TblRequerMainten.NoOfLabs,TblRequerMainten.supervisorNotes,TblRequerMainten.RemainKmToArrive,TblRequerMainten.StopLocation"
        StrSQL = StrSQL & " FROM dbo.FixedAssets RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblRequerMainten ON dbo.FixedAssets.id = dbo.TblRequerMainten.EquepID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.Nationality RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee ON dbo.Nationality.id = dbo.TblEmployee.NationlID ON dbo.TblRequerMainten.LeaderID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id"
        StrSQL = StrSQL & " Where (1 = 1)"
 
        If Not IsNull(Me.DtpDateFrom2.value) Then
            StrSQL = StrSQL & " AND dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
        End If
       
        If Not IsNull(Me.DtpDateTo2.value) Then
            StrSQL = StrSQL & " AND dbo.TblRequerMainten.RecordDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
        End If
        
        If Me.DcbEqup2.Text <> "" And val(DcbEqup2.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND dbo.TblRequerMainten.EquepID = " & val(Me.DcbEqup2.BoundText)
        End If

        If Me.DCboUserName.Text <> "" And val(DCboUserName.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND dbo.TblRequerMainten.userid = " & val(Me.DCboUserName.BoundText)
        End If

        If Me.EquipmentStatusid.ListIndex <> -1 And Me.EquipmentStatusid.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND dbo.TblRequerMainten.EquipmentStatusid = " & val(Me.EquipmentStatusid.ListIndex)
        End If

        If Me.DcbStatusMaint.ListIndex <> -1 And Me.DcbStatusMaint.ListIndex <> 4 Then
            StrSQL = StrSQL & " AND   dbo.TblRequerMainten.StatusMaint = " & val(Me.DcbStatusMaint.ListIndex)
        End If
        'DcbStatusMaint
    End If

    If Emp.value = True Then
        StrSQL = " SELECT dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_id, dbo.TblEmpDepartments.DeparmentID, "
        StrSQL = StrSQL & " dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
        StrSQL = StrSQL & " dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
        StrSQL = StrSQL & " dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.NationlID,"
        StrSQL = StrSQL & " dbo.Nationality.name, dbo.Nationality.namee, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.BranchId, dbo.TblEmployee.JobTypeID,"
        StrSQL = StrSQL & " dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.DeptID2, dbo.TblEmpDepartmentsDet.Name AS DepartName,"
        StrSQL = StrSQL & " dbo.TblEmpDepartmentsDet.NameE AS DepartNameE, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile"
        StrSQL = StrSQL & " FROM  dbo.TblEmpDepartmentsDet RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TblEmpDepartmentsDet.ID = dbo.TblEmployee.DeptID2 LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id"
        StrSQL = StrSQL & " Where (1 = 1)"
    End If

    If Emp.value = True Or ReqMaint.value = True Then
        If Me.DcbDept2.Text <> "" And val(DcbDept2.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   dbo.TblEmpDepartments.DeparmentID = " & val(Me.DcbDept2.BoundText)
        End If
    End If

    If Me.Order.value = True Then
        StrSQL = "SELECT distinct  dbo.tblordermaintenancetypes.Remarks , StutsMaint,dbo.TblOrderMaint.ID AS TblOrderMaintid, dbo.TblOrderMaint.RecordDate, dbo.TblBranchesData.branch_id, dbo.TblOrderMaint.BranchID, "
        StrSQL = StrSQL & " dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,tblordermaintenancetypes.Qty ShowQty  ,dbo.Transaction_Details.ShowQty ShowQty2, dbo.Transaction_Details.showPrice,"
        StrSQL = StrSQL & " dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.LeaderName, dbo.TblOrderMaint.LeaderID, TblEmployee_2.Emp_ID, TblEmployee_2.Emp_Name,"
        StrSQL = StrSQL & " TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4, TblEmployee_2.Fullcode AS EmpFullcode,"
        StrSQL = StrSQL & " TblEmployee_2.Emp_Namee4, TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.SuperVisor, TblEmployee_1.Emp_Name AS SuperEmp_Name, TblEmployee_1.Fullcode AS SuperFullcode,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Name1 AS SuperEmp_Name1, TblEmployee_1.Emp_Name2 AS SuperEmp_Name2, TblEmployee_1.Emp_Name3 AS SuperEmp_Name3,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Name4 AS SuperEmp_Name4, TblEmployee_1.Emp_Namee4 AS SuperEmp_Namee4, TblEmployee_1.Emp_Namee3 AS SuperEmp_Namee3,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Namee2 AS SuperEmp_Namee2, TblEmployee_1.Emp_Namee1 AS SuperEmp_Namee1, TblEmployee_1.Emp_Namee AS SuperEmp_Namee,"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes.maintenanceid, dbo.tblordermaintenancetypes.TypeTrans, dbo.TblMaintenanceType.name AS Mainname,"
        StrSQL = StrSQL & " dbo.TblMaintenanceType.namee AS MainnameE, "
        StrSQL = StrSQL & " dbo.Groups.GroupNamee, dbo.Transaction_Details.ItemSerial, dbo.TblItems.barCodeNO, dbo.TblItems.ItemCode, dbo.TblItems.Fullcode AS ItemFullcode,"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes.ORderID, dbo.Transaction_Details.EqupID, dbo.TblCarsData.Fullcode AS CarFullcode, dbo.TblCarsData.BoardNO,"
        StrSQL = StrSQL & " dbo.TblCarsData.OperatorN, dbo.TblCarsData.fixedAssetid, dbo.tblordermaintenancetypes.Head_Details, dbo.tblordermaintenancetypesQry.PartID, TblCarsData_1.Fullcode AS FullCode2,"
        StrSQL = StrSQL & " TblCarsData_1.BoardNO AS BoardNO2"
        StrSQL = StrSQL & " FROM dbo.TblCarsData TblCarsData_1 INNER JOIN"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypesQry INNER JOIN"
        StrSQL = StrSQL & " dbo.TblMaintenanceType INNER JOIN"
        StrSQL = StrSQL & " dbo.TblOrderMaint INNER JOIN"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes ON dbo.TblOrderMaint.ID = dbo.tblordermaintenancetypes.ORderID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblOrderMaint.BranchID = dbo.TblBranchesData.branch_id ON"
        StrSQL = StrSQL & " dbo.TblMaintenanceType.id = dbo.tblordermaintenancetypes.maintenanceid INNER JOIN"
        StrSQL = StrSQL & " dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & " dbo.TblCarsData ON dbo.FixedAssets.id = dbo.TblCarsData.fixedAssetid ON dbo.TblOrderMaint.EquepID = dbo.FixedAssets.id ON"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypesQry.ORderID = dbo.TblOrderMaint.ID INNER JOIN"
        StrSQL = StrSQL & " dbo.FixedAssets FixedAssets_1 ON dbo.tblordermaintenancetypesQry.PartID = FixedAssets_1.id ON TblCarsData_1.fixedAssetid = FixedAssets_1.id LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee TblEmployee_2 ON dbo.TblOrderMaint.LeaderID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.Transaction_Details INNER JOIN"
        StrSQL = StrSQL & " dbo.Groups INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes.maintenanceid = dbo.Transaction_Details.MintID AND"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes.OrderID = dbo.Transaction_Details.orderNo"
 
        'newqry 19 06 2018
        StrSQL = "SELECT  distinct dbo.tblordermaintenancetypes.remarks  AS Gridremarks,  dbo.TblOrderMaint.StutsMaint, dbo.TblOrderMaint.ID AS TblOrderMaintidS, dbo.TblOrderMaint.RecordDate, dbo.TblBranchesData.branch_id, "
        StrSQL = StrSQL & " dbo.TblOrderMaint.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transaction_Details.ShowQty,"
        StrSQL = StrSQL & " dbo.Transaction_Details.showPrice, dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
        StrSQL = StrSQL & " dbo.TblItems.Fullcode, dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.LeaderName, dbo.TblOrderMaint.LeaderID, TblEmployee_2.Emp_ID,"
        StrSQL = StrSQL & " TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4,"
        StrSQL = StrSQL & " TblEmployee_2.Fullcode AS EmpFullcode, TblEmployee_2.Emp_Namee4, TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2,"
        StrSQL = StrSQL & " TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblOrderMaint.SuperVisor, TblEmployee_1.Emp_Name AS SuperEmp_Name,"
        StrSQL = StrSQL & " TblEmployee_1.Fullcode AS SuperFullcode, TblEmployee_1.Emp_Name1 AS SuperEmp_Name1, TblEmployee_1.Emp_Name2 AS SuperEmp_Name2,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Name3 AS SuperEmp_Name3, TblEmployee_1.Emp_Name4 AS SuperEmp_Name4, TblEmployee_1.Emp_Namee4 AS SuperEmp_Namee4,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Namee3 AS SuperEmp_Namee3, TblEmployee_1.Emp_Namee2 AS SuperEmp_Namee2, TblEmployee_1.Emp_Namee1 AS SuperEmp_Namee1,"
        StrSQL = StrSQL & " TblEmployee_1.Emp_Namee AS SuperEmp_Namee, dbo.tblordermaintenancetypes.maintenanceid, dbo.tblordermaintenancetypes.TypeTrans,"
        StrSQL = StrSQL & " dbo.TblMaintenanceType.name AS Mainname, dbo.TblMaintenanceType.namee AS MainnameE, "
        StrSQL = StrSQL & " dbo.Transaction_Details.ItemSerial, dbo.TblItems.barCodeNO,"
        StrSQL = StrSQL & " dbo.TblItems.ItemCode, dbo.TblItems.Fullcode AS ItemFullcode, dbo.tblordermaintenancetypes.ORderID, dbo.Transaction_Details.EqupID,"
        StrSQL = StrSQL & " TblCarsData_1.Fullcode AS CarFullcode, TblCarsData_1.BoardNO, TblCarsData_1.OperatorN, TblCarsData_1.fixedAssetid,"
        StrSQL = StrSQL & " dbo.tblordermaintenancetypes.Head_Details, dbo.tblordermaintenancetypesQry.PartID, TblCarsData_1.Fullcode AS FullCode2, TblCarsData_2.BoardNO AS BoardNO2,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.carendperiod1, dbo.TblOrderMaint.carendperiod, dbo.TblOrderMaint.report1des1, dbo.TblOrderMaint.report1des, dbo.TblOrderMaint.alarmsPeriod,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.alarms, dbo.TblOrderMaint.mangercomment, dbo.TblOrderMaint.separatedreport1, dbo.TblOrderMaint.separatedreport, dbo.TblOrderMaint.LastKM,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.CurrKM, dbo.tblordermaintenancetypes.EmpID, TblEmployee_3.Emp_Code AS technicalCode, TblEmployee_3.Emp_Name AS technicalName,"
        StrSQL = StrSQL & " dbo.TblOrderMaint.EnterTime"
        
        
        
        
       StrSQL = "   SELECT distinct dbo.tblordermaintenancetypes.remarks AS Gridremarks,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.StutsMaint,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.ID           AS TblOrderMaintidS,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.RecordDate,"
       StrSQL = StrSQL & " dbo.TblBranchesData.branch_id,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.BranchID,"
       StrSQL = StrSQL & " dbo.TblBranchesData.branch_name,"
       StrSQL = StrSQL & " dbo.TblBranchesData.branch_namee,"
       StrSQL = StrSQL & " tblordermaintenancetypes.Qty ShowQty,"
       StrSQL = StrSQL & " ShowQty2 = (           SELECT avg(ShowQty)           FROM   Transaction_Details                  INNER JOIN Transactions AS t                       ON  t.Transaction_ID = Transaction_Details.Transaction_ID           WHERE  Transaction_Details.Transaction_ID = t.Transaction_ID                  AND t.order_no = TblOrderMaint.ID       ),"
        StrSQL = StrSQL & " showPrice = (SELECT SUM(showPrice) FROM   Transaction_Details INNER JOIN Transactions AS t ON  t.Transaction_ID = Transaction_Details.Transaction_ID WHERE  Transaction_Details.Transaction_ID = t.Transaction_ID AND t.order_no = TblOrderMaint.ID),"
        StrSQL = StrSQL & " ItemSerial = (SELECT TOP 1 ItemSerial FROM   Transaction_Details INNER JOIN Transactions AS t ON  t.Transaction_ID = Transaction_Details.Transaction_ID WHERE  Transaction_Details.Transaction_ID = t.Transaction_ID AND t.order_no = TblOrderMaint.ID       ),"
       
       
      ' StrSQL = StrSQL & " ShowQty = 0,"
      ' StrSQL = StrSQL & " showPrice = 0,"
       'StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty,"
       'StrSQL = StrSQL & " dbo.Transaction_Details.showPrice,"
      ' StrSQL = StrSQL & " dbo.Transaction_Details.OperPrice,"
      ' StrSQL = StrSQL & " dbo.Transaction_Details.Item_ID,"
    
       StrSQL = StrSQL & " dbo.TblOrderMaint.LeaderType,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.LeaderName,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.LeaderID,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_ID,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Name,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Name1,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Name2,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Name3,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Name4,"
       StrSQL = StrSQL & " TblEmployee_2.Fullcode         AS EmpFullcode,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Namee4,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Namee3,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Namee2,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Namee1,"
       StrSQL = StrSQL & " TblEmployee_2.Emp_Namee,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.SuperVisor,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name         AS SuperEmp_Name,"
       StrSQL = StrSQL & " TblEmployee_1.Fullcode         AS SuperFullcode,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name1        AS SuperEmp_Name1,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name2        AS SuperEmp_Name2,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name3        AS SuperEmp_Name3,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name4        AS SuperEmp_Name4,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Namee4       AS SuperEmp_Namee4,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Namee3       AS SuperEmp_Namee3,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Namee2       AS SuperEmp_Namee2,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Namee1       AS SuperEmp_Namee1,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Namee        AS SuperEmp_Namee,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.maintenanceid,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.TypeTrans,"
       StrSQL = StrSQL & " dbo.TblMaintenanceType.name    AS Mainname,"
       StrSQL = StrSQL & " dbo.TblMaintenanceType.namee   AS MainnameE,"
       
   
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.ORderID,"
      ' StrSQL = StrSQL & " dbo.Transaction_Details.EqupID,"
       StrSQL = StrSQL & " TblCarsData_1.Fullcode         AS CarFullcode,"
       StrSQL = StrSQL & " TblCarsData_1.BoardNO,"
       StrSQL = StrSQL & " TblCarsData_1.OperatorN,"
       StrSQL = StrSQL & " TblCarsData_1.fixedAssetid,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.Head_Details,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypesQry.PartID,"
       StrSQL = StrSQL & " TblCarsData_1.Fullcode         AS FullCode2,"
       StrSQL = StrSQL & " TblCarsData_2.BoardNO          AS BoardNO2,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.carendperiod1,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.carendperiod,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.report1des1,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.report1des,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.alarmsPeriod,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.alarms,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.mangercomment,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.separatedreport1,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.separatedreport,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.LastKM,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.CurrKM,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.EmpID,"
       StrSQL = StrSQL & " TblEmployee_3.Emp_Code         AS technicalCode,"
       StrSQL = StrSQL & " TblEmployee_3.Emp_Name         AS technicalName,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.EnterTime"
'        StrSQL = StrSQL & " FROM            FixedAssets RIGHT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblCarsData AS TblCarsData_1 INNER JOIN"
'        StrSQL = StrSQL & "                          tblordermaintenancetypesQry INNER JOIN"
'        StrSQL = StrSQL & "                          FixedAssets AS FixedAssets_1 ON tblordermaintenancetypesQry.PartID = FixedAssets_1.id ON TblCarsData_1.fixedAssetid = FixedAssets_1.id RIGHT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblMaintenanceType INNER JOIN"
'        StrSQL = StrSQL & "                          tblordermaintenancetypes ON TblMaintenanceType.id = tblordermaintenancetypes.maintenanceid INNER JOIN"
'        StrSQL = StrSQL & "                          TblEmployee AS TblEmployee_3 ON tblordermaintenancetypes.EmpID = TblEmployee_3.Emp_ID RIGHT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblOrderMaint ON tblordermaintenancetypes.ORderID = TblOrderMaint.ID ON tblordermaintenancetypesQry.ORderID = TblOrderMaint.ID LEFT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblBranchesData ON TblOrderMaint.BranchID = TblBranchesData.branch_id ON FixedAssets.id = TblOrderMaint.EquepID LEFT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblCarsData AS TblCarsData_2 ON FixedAssets.id = TblCarsData_2.fixedAssetid LEFT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblEmployee AS TblEmployee_1 ON TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
'        StrSQL = StrSQL & "                          TblEmployee AS TblEmployee_2 ON TblOrderMaint.LeaderID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
'        StrSQL = StrSQL & "                          Transaction_Details INNER JOIN"
'        StrSQL = StrSQL & "                          Groups INNER JOIN"
'        StrSQL = StrSQL & "                          TblItems ON Groups.GroupID = TblItems.GroupID ON Transaction_Details.Item_ID = TblItems.ItemID ON"
'        StrSQL = StrSQL & "                          tblordermaintenancetypes.maintenanceid = Transaction_Details.MintID And tblordermaintenancetypes.OrderID = Transaction_Details.orderNo "


'        StrSQL = StrSQL & "    From Transaction_Details"
'        StrSQL = StrSQL & "           INNER JOIN Groups"
'        StrSQL = StrSQL & "           INNER JOIN TblItems"
'        StrSQL = StrSQL & "                ON  Groups.GroupID = TblItems.GroupID"
'        StrSQL = StrSQL & "                ON  Transaction_Details.Item_ID = TblItems.ItemID"
'        StrSQL = StrSQL & "           RIGHT OUTER JOIN Transactions     AS t"
'        StrSQL = StrSQL & "                ON  Transaction_Details.Transaction_ID = t.Transaction_ID"
'        StrSQL = StrSQL & "           RIGHT OUTER JOIN FixedAssets"
'        StrSQL = StrSQL & "           RIGHT OUTER JOIN TblCarsData      AS TblCarsData_1"
'        StrSQL = StrSQL & "           INNER JOIN tblordermaintenancetypesQry"
'        StrSQL = StrSQL & "           INNER JOIN FixedAssets            AS FixedAssets_1"
'        StrSQL = StrSQL & "                ON  tblordermaintenancetypesQry.PartID = FixedAssets_1.id"
'        StrSQL = StrSQL & "                ON  TblCarsData_1.fixedAssetid = FixedAssets_1.id"
'        StrSQL = StrSQL & "           RIGHT OUTER JOIN TblMaintenanceType"
'        StrSQL = StrSQL & "           INNER JOIN tblordermaintenancetypes"
'        StrSQL = StrSQL & "                ON  TblMaintenanceType.id = tblordermaintenancetypes.maintenanceid"
'        StrSQL = StrSQL & "           INNER JOIN TblEmployee            AS TblEmployee_3"
'        StrSQL = StrSQL & "                ON  tblordermaintenancetypes.EmpID = TblEmployee_3.Emp_ID"
'        StrSQL = StrSQL & "           RIGHT OUTER JOIN TblOrderMaint"
'        StrSQL = StrSQL & "                ON  tblordermaintenancetypes.ORderID = TblOrderMaint.ID"
'        StrSQL = StrSQL & "                ON  tblordermaintenancetypesQry.ORderID = TblOrderMaint.ID"
'        StrSQL = StrSQL & "           LEFT OUTER JOIN TblBranchesData"
'        StrSQL = StrSQL & "                ON  TblOrderMaint.BranchID = TblBranchesData.branch_id"
'        StrSQL = StrSQL & "                ON  FixedAssets.id = TblOrderMaint.EquepID"
'        StrSQL = StrSQL & "           LEFT OUTER JOIN TblCarsData       AS TblCarsData_2"
'        StrSQL = StrSQL & "                ON  FixedAssets.id = TblCarsData_2.fixedAssetid"
'        StrSQL = StrSQL & "           LEFT OUTER JOIN TblEmployee       AS TblEmployee_1"
'        StrSQL = StrSQL & "                ON  TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID"
'        StrSQL = StrSQL & "           LEFT OUTER JOIN TblEmployee       AS TblEmployee_2"
'        StrSQL = StrSQL & "                ON  TblOrderMaint.LeaderID = TblEmployee_2.Emp_ID"
'        StrSQL = StrSQL & "                ON  t.order_no = TblOrderMaint.ID"
'        StrSQL = StrSQL & "                AND t.BillBasedOn = 8"
'        StrSQL = StrSQL & " Where 1=1"

StrSQL = StrSQL & " FROM   "
StrSQL = StrSQL & "       FixedAssets"
StrSQL = StrSQL & "       RIGHT OUTER JOIN TblCarsData      AS TblCarsData_1"
StrSQL = StrSQL & "       INNER JOIN tblordermaintenancetypesQry"
StrSQL = StrSQL & "       INNER JOIN FixedAssets            AS FixedAssets_1"
StrSQL = StrSQL & "            ON  tblordermaintenancetypesQry.PartID = FixedAssets_1.id"
StrSQL = StrSQL & "            ON  TblCarsData_1.fixedAssetid = FixedAssets_1.id"
StrSQL = StrSQL & "       RIGHT OUTER JOIN TblMaintenanceType"
StrSQL = StrSQL & "       INNER JOIN tblordermaintenancetypes"
StrSQL = StrSQL & "            ON  TblMaintenanceType.id = tblordermaintenancetypes.maintenanceid"
StrSQL = StrSQL & "       INNER JOIN TblEmployee            AS TblEmployee_3"
StrSQL = StrSQL & "            ON  tblordermaintenancetypes.EmpID = TblEmployee_3.Emp_ID"
StrSQL = StrSQL & "       RIGHT OUTER JOIN TblOrderMaint"
StrSQL = StrSQL & "            ON  tblordermaintenancetypes.ORderID = TblOrderMaint.ID"
StrSQL = StrSQL & "            ON  tblordermaintenancetypesQry.ORderID = TblOrderMaint.ID"
StrSQL = StrSQL & "       LEFT OUTER JOIN TblBranchesData"
StrSQL = StrSQL & "            ON  TblOrderMaint.BranchID = TblBranchesData.branch_id"
StrSQL = StrSQL & "            ON  FixedAssets.id = TblOrderMaint.EquepID"
StrSQL = StrSQL & "       LEFT OUTER JOIN TblCarsData       AS TblCarsData_2"
StrSQL = StrSQL & "            ON  FixedAssets.id = TblCarsData_2.fixedAssetid"
StrSQL = StrSQL & "       LEFT OUTER JOIN TblEmployee       AS TblEmployee_1"
StrSQL = StrSQL & "            ON  TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID"
StrSQL = StrSQL & "       LEFT OUTER JOIN TblEmployee       AS TblEmployee_2"
StrSQL = StrSQL & "            ON  TblOrderMaint.LeaderID = TblEmployee_2.Emp_ID"

StrSQL = StrSQL & " Where (1 = 1)"
    
        If Me.TXTOrderMaintID <> "" Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.id  in ( " & Me.TXTOrderMaintID.Text & ")"
        End If

        If Me.DCboUserName.Text <> "" And val(DCboUserName.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.userid = " & val(Me.DCboUserName.BoundText)
        End If

        If Me.DcbTypeMain2.Text <> "" And val(Me.DcbTypeMain2.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND  (   dbo.tblordermaintenancetypes.maintenanceid = " & val(Me.DcbTypeMain2.BoundText)
            StrSQL = StrSQL & " or   dbo.tblordermaintenancetypes.maintenanceid in ( "
            StrSQL = StrSQL & "               select id from TblMaintenanceType"
            StrSQL = StrSQL & "         Where (MainType = 0 Or MainType Is Null)"
            StrSQL = StrSQL & "    and  FollowID=   " & val(Me.DcbTypeMain2.BoundText) & "))"
        End If

        'XX
        If Me.DcbStutsMaint.ListIndex <> -1 And Me.DcbStutsMaint.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.StutsMaint = " & val(Me.DcbStutsMaint.ListIndex)
        End If
        
        If Me.DcbEmp2.Text <> "" And val(DcbEmp2.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.SuperVisor = " & val(Me.DcbEmp2.BoundText)
        End If
    End If

    If Me.DcbLeader2.Text <> "" And val(DcbLeader2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblEmployee.Emp_ID = " & val(Me.DcbLeader2.BoundText)
    End If
    
    If Me.dcsupervisor.Text <> "" And val(dcsupervisor.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   tblordermaintenancetypes.SuperID= " & val(Me.dcsupervisor.BoundText)
    End If

    If Me.dctechnical.Text <> "" And val(dctechnical.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   tblordermaintenancetypes.EmpID= " & val(Me.dctechnical.BoundText)
    End If

    If Me.DCMaintenanceTypes.Text <> "" And val(DCMaintenanceTypes.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   tblordermaintenancetypes.GroupID= " & val(Me.DCMaintenanceTypes.BoundText)
    End If

    If Me.DcbBranch2.Text <> "" And val(DcbBranch2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch2.BoundText)
    End If

    If Me.Order.value = True Then
        If Me.DcbGroup.BoundText <> "" And val(DcbGroup.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   dbo.TblItems.GroupID = " & val(Me.DcbGroup.BoundText)
        End If
        
        'If Me.DcbEqup2.Text <> "" And val(DcbEqup2.BoundText) <> 0 Then
        'StrSQL = StrSQL & " AND  dbo.Transaction_Details.EqupID= " & val(Me.DcbEqup2.BoundText)
        'End If

        If Me.DcbEqup2.Text <> "" And val(DcbEqup2.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND  (TblCarsData_2.fixedAssetid= " & val(Me.DcbEqup2.BoundText) & " or dbo.TblOrderMaint.EquepID in ( select PARTID from TblCarsDataDet WHERE EQUPID =" & val(Me.DcbEqup2.BoundText) & "))"
        End If

        'If Me.DcbEqup2.Text <> "" And val(DcbEqup2.BoundText) <> 0 Then
            'StrSQL = StrSQL & " AND  dbo.Transactions.FixesAssetsID= " & val(Me.DcbEqup2.BoundText)
        'End If
 
        If Not IsNull(Me.DtpDateFrom2.value) Then
            StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
        End If
       
        If Not IsNull(Me.DtpDateTo2.value) Then
            StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
        End If

        StrSQL = StrSQL & "   ORDER BY dbo.TblOrderMaint.ID,dbo.tblordermaintenancetypes.Head_Details"
    End If


If Order2.value = True Then
StrSQL = " SELECT     TotalHours = DATEDIFF(hour, CAST(EnterTime AS DATETIME), CAST(endmaintenanceTime AS DATETIME)) + DATEDIFF(Hour, TblOrderMaint.EnterDate, TblOrderMaint.endmaintenanceDate) , dbo.TblOrderMaint.ID, dbo.TblOrderMaint.EquepID, dbo.TblOrderMaint.LeaderID, dbo.TblOrderMaint.SuperVisor, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.EnterDate, "
StrSQL = StrSQL & "                         dbo.TblOrderMaint.EnterTime, dbo.TblOrderMaint.endmaintenanceDate, dbo.TblOrderMaint.endmaintenanceTime, dbo.TblOrderMaint.StutsMaint,"
StrSQL = StrSQL & "                         dbo.TblEmployee.Emp_Name AS Supervisonname, dbo.TblEmployee.Emp_Namee AS SupervisonnameE, TblEmployee_1.Emp_Name AS Drivername,"
StrSQL = StrSQL & "                         TblEmployee_1.Emp_Namee AS Drivernamee, dbo.TblOrderMaint.RecordDate, dbo.FixedAssets.Name AS equName, dbo.FixedAssets.Fullcode AS equCode,"
StrSQL = StrSQL & "                         dbo.FixedAssets.namee AS equNamee"
StrSQL = StrSQL & "   FROM         dbo.TblOrderMaint INNER JOIN"
StrSQL = StrSQL & "                         dbo.TblEmployee ON dbo.TblOrderMaint.SuperVisor = dbo.TblEmployee.Emp_ID INNER JOIN"
StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID INNER JOIN"
StrSQL = StrSQL & "                         dbo.FixedAssets ON dbo.TblOrderMaint.EquepID = dbo.FixedAssets.id"
 

     If Not IsNull(Me.DtpDateFrom2.value) Then
            StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
        End If
       
        If Not IsNull(Me.DtpDateTo2.value) Then
            StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
        End If
        
        If Me.DcbEqup2.Text <> "" And val(DcbEqup2.BoundText) <> 0 Then
            StrSQL = StrSQL & " and   dbo.TblOrderMaint.EquepID = " & val(Me.DcbEqup2.BoundText)
        End If
        
        
            If Me.DcbLeader2.Text <> "" And val(DcbLeader2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblOrderMaint.LeaderID = " & val(Me.DcbLeader2.BoundText)
    End If
    
    If Me.dcsupervisor.Text <> "" And val(dcsupervisor.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   TblOrderMaint.SuperVisor= " & val(Me.dcsupervisor.BoundText)
    End If
    
            If Me.DcbStutsMaint.ListIndex <> -1 And Me.DcbStutsMaint.ListIndex <> 4 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.StutsMaint = " & val(Me.DcbStutsMaint.ListIndex)
                                                                                    
        End If
    
        
        
StrSQL = StrSQL & "   ORDER BY dbo.TblOrderMaint.ID"

        
End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        Else
            Msg = "No Data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        rs.MoveFirst
        print_report3 StrSQL
    End If
    Exit Sub
ErrTrap:
   If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÊÃœ Œÿ√ »«·‘—Êÿ"
   Else
        Msg = "No Data"
   End If
End Sub
Public Sub GetData1()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
          Msg = "Â–Â «·ÿ»«⁄…  Õ·Ì·Ì " & CHR(13)
        Msg = Msg + " Â·  —Ìœ ÿ»«⁄… «Ã„«·Ì"
        Dim isTotal As Boolean
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
        isTotal = True
        StrSQL = " SELECT DISTINCT dbo.TblOrderMaint.ID,"
       StrSQL = StrSQL & "  nooforders=(select count( EquepID) from TblRequerMainten  f where f.EquepID=TblOrderMaint.EquepID),"
        
       StrSQL = StrSQL & " dbo.TblOrderMaint.RecordDate,dbo.TblCarsData.OwnerName,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.BranchID,"
       StrSQL = StrSQL & " TblBranchesData_2.branch_name,"
       StrSQL = StrSQL & " TblBranchesData_2.branch_namee,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.UserID,"
       StrSQL = StrSQL & " dbo.FixedAssets.Name,"
       StrSQL = StrSQL & " dbo.FixedAssets.namee,"
       StrSQL = StrSQL & " TblEmployee_1.Emp_Name,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.TypeMaint,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.Cost,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.Des,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.EquepmentName,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.Total,"
       StrSQL = StrSQL & " dbo.tblordermaintenancetypes.BillNo,"
       StrSQL = StrSQL & " dbo.Transactions.Transaction_Type,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.TotalSand,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.TotalSpare,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.TotalMaint,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.BoardNO,"
       StrSQL = StrSQL & " dbo.FixedAssets.code,"
       StrSQL = StrSQL & " dbo.TblOrderMaint.EquepID,"
       StrSQL = StrSQL & " dbo.TblCarsData.Fullcode      AS CarFullcode,"
       StrSQL = StrSQL & " dbo.TblCarsData.Model,"
       StrSQL = StrSQL & " dbo.TblCarsData.CarsTypeId,"
       StrSQL = StrSQL & " dbo.TBLCarTypes.name          AS CrTypname,"
       StrSQL = StrSQL & " dbo.TBLCarTypes.namee         AS CrTypnameE,"
       
       'If CHKiSoUT.value = vbChecked Then
            StrSQL = StrSQL & " t.Transaction_Date,"
            StrSQL = StrSQL & " td.showPrice,"
            StrSQL = StrSQL & " td.ID                         AS TrnsID,"
            StrSQL = StrSQL & " td.OperPrice,"
            StrSQL = StrSQL & " Td.ShowQty,"
'        Else
'             StrSQL = StrSQL & "Transaction_Date = '' ,"
'            StrSQL = StrSQL & " showPrice = 0,"
'            StrSQL = StrSQL & " TrnsID = 0,"
'            StrSQL = StrSQL & " OperPrice = 0,"
'            StrSQL = StrSQL & " ShowQty = 0,"
'        End If
        
        StrSQL = StrSQL & " dbo.GetCarsExpensValue(dbo.TblOrderMaint.EquepID) AS OtherValue"
       StrSQL = StrSQL & " From dbo.TblCarModels"
        StrSQL = StrSQL & "        RIGHT OUTER JOIN dbo.TblCarsData"
        StrSQL = StrSQL & "             ON  dbo.TblCarModels.Id = dbo.TblCarsData.VModel"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TBLCarTypes"
        StrSQL = StrSQL & "             ON  dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id"
        StrSQL = StrSQL & "        RIGHT OUTER JOIN dbo.FixedAssets"
        StrSQL = StrSQL & "             ON  dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id"
        StrSQL = StrSQL & " RIGHT OUTER JOIN dbo.TblOrderMaint"
        StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblEmployee TblEmployee_4"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.DrievID = TblEmployee_4.Emp_ID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblMaintenanceType"
        StrSQL = StrSQL & "        RIGHT OUTER JOIN dbo.TblEmployee TblEmployee_5"
        StrSQL = StrSQL & "        RIGHT OUTER JOIN dbo.TblEmpDepartments"
        StrSQL = StrSQL & "        RIGHT OUTER JOIN dbo.tblordermaintenancetypes"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.Transaction_Details"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.Transactions"
        StrSQL = StrSQL & "             ON  dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
        StrSQL = StrSQL & "             ON  dbo.tblordermaintenancetypes.Transaction_IDDet = dbo.Transaction_Details.ID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblCustemers"
        StrSQL = StrSQL & "             ON  dbo.tblordermaintenancetypes.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL & "             ON  dbo.TblEmpDepartments.DeparmentID = dbo.tblordermaintenancetypes.DeptID"
        StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblEmployee TblEmployee_6"
        StrSQL = StrSQL & "             ON  dbo.tblordermaintenancetypes.SuperID = TblEmployee_6.Emp_ID"
        StrSQL = StrSQL & "             ON  TblEmployee_5.Emp_ID = dbo.tblordermaintenancetypes.EmpID"
        StrSQL = StrSQL & "             ON  dbo.TblMaintenanceType.id = dbo.tblordermaintenancetypes.maintenanceid"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.ID = dbo.tblordermaintenancetypes.ORderID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblEmployee TblEmployee_2"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.SuperVisor = TblEmployee_2.Emp_ID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblEmployee TblEmployee_1"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.reciverid = TblEmployee_1.Emp_ID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblEmployee TblEmployee_3"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.LeaderID = TblEmployee_3.Emp_ID"
        StrSQL = StrSQL & "        LEFT OUTER JOIN dbo.TblBranchesData TblBranchesData_1"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.DcbBranchFrom = TblBranchesData_1.branch_id"
        StrSQL = StrSQL & "             ON  dbo.FixedAssets.id = dbo.TblOrderMaint.EquepID"
        StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblBranchesData TblBranchesData_2"
        StrSQL = StrSQL & "             ON  dbo.TblOrderMaint.BranchID = TblBranchesData_2.branch_id"
        
        
           'If CHKiSoUT.value = vbChecked Then
                  StrSQL = StrSQL & "                       LEFT OUTER JOIN Transactions AS t ON T.order_no = TblOrderMaint.ID"
                  StrSQL = StrSQL & "      LEFT OUTER JOIN Transaction_Details AS td ON T.Transaction_ID= Td.Transaction_ID"
                  
                  
                  
           ' End If
        Else
        isTotal = False
        StrSQL = " SELECT   DISTINCT  dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate, dbo.TblOrderMaint.BranchID, TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, "
        StrSQL = StrSQL & "                     dbo.TblOrderMaint.UserID, dbo.FixedAssets.Name, dbo.FixedAssets.namee, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2,"
        StrSQL = StrSQL & "                     TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1,"
        StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Jiha, dbo.TblOrderMaint.Remarks,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.startmaintenanceTime, dbo.TblOrderMaint.endmaintenanceTime,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.RecmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks, dbo.TblOrderMaint.TechNote,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.reciverid, TblEmployee_1.Emp_Name AS ReciEmp_Name, TblEmployee_1.Emp_Name1 AS ReciEmp_Name1,"
        StrSQL = StrSQL & "                    TblEmployee_1.Emp_Name2 AS ReciEmp_Name2, TblEmployee_1.Emp_Name3 AS ReciEmp_Name3, TblEmployee_1.Fullcode AS ReciFullcode,"
        StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee4 AS ReciEmp_Namee4, TblEmployee_1.Emp_Namee3 AS ReciEmp_Namee3, TblEmployee_1.Emp_Namee2 AS ReciEmp_Namee2,"
        StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee1 AS ReciEmp_Namee1, TblEmployee_1.Emp_Namee AS RecieEmp_Namee, dbo.TblOrderMaint.endmaintenanceDate,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.ended, dbo.TblOrderMaint.ReqMainID, TblEmployee_1.Emp_Namee4,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Remarks AS RemarksDet, dbo.tblordermaintenancetypes.ID AS IDDet, dbo.tblordermaintenancetypes.ORderID,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.maintenanceid, dbo.TblMaintenanceType.name AS nameMType, dbo.TblMaintenanceType.namee AS nameMTypeE,"
        
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName, dbo.TblOrderMaint.EquepmentName, dbo.TblOrderMaint.Total,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.DcbBranchFrom, TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.LeaderID, TblEmployee_3.Emp_Name AS LeaderEmp_Name, TblEmployee_3.Fullcode AS LeaderFullcode,"
        StrSQL = StrSQL & "                    TblEmployee_3.Emp_Namee AS LeaderEmp_NameE, dbo.TblOrderMaint.SuperVisor, TblEmployee_2.Emp_Name AS SuperEmp_Name,"
        StrSQL = StrSQL & "                    TblEmployee_2.Fullcode AS SuperFullcode, TblEmployee_2.Emp_Namee AS SuperEmp_NameE, dbo.TblOrderMaint.DrievID,"
        StrSQL = StrSQL & "                    TblEmployee_4.Emp_Name AS DevEmp_Name, TblEmployee_4.Fullcode AS DevFullcode, TblEmployee_4.Emp_Namee AS DevEmp_NameE,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.LocaMaint, dbo.tblordermaintenancetypes.Company,"
        StrSQL = StrSQL & "                   "
        
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Price , "
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Total AS TotalDet, dbo.tblordermaintenancetypes.Qty,"
        
        StrSQL = StrSQL & "                     dbo.tblordermaintenancetypes.BillNo, dbo.tblordermaintenancetypes.CusMobile,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.PartName, dbo.tblordermaintenancetypes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
        StrSQL = StrSQL & "                    dbo.TblCustemers.Fullcode AS CusFullcode, dbo.tblordermaintenancetypes.EmpID, TblEmployee_5.Emp_Name AS FiterEmp_Name,"
        StrSQL = StrSQL & "                    TblEmployee_5.Fullcode AS FiterFullcode, TblEmployee_5.Emp_Namee AS FiterEmp_NameE, dbo.tblordermaintenancetypes.SuperID,"
        StrSQL = StrSQL & "                    TblEmployee_6.Emp_Name AS SuperEmp_NameDet, TblEmployee_6.Fullcode AS SuperFullcodeDet, TblEmployee_6.Emp_Namee AS SuperEmp_NameDetE,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Transaction_ID AS Transaction_IDH, dbo.tblordermaintenancetypes.Transaction_IDDet, "
        StrSQL = StrSQL & "                    dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate,"
        StrSQL = StrSQL & "                    dbo.Transactions.TransactionComment, dbo.tblordermaintenancetypes.TypeTrans, dbo.Transactions.OpOrderID, dbo.Transactions.OldOpOrderID,"
        StrSQL = StrSQL & "                     dbo.TblOrderMaint.TotalSand, dbo.TblOrderMaint.TotalSpare, dbo.TblOrderMaint.TotalMaint, dbo.TblOrderMaint.OperatorN,"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.BoardNO, dbo.TblOrderMaint.StutsMaint, dbo.TblOrderMaint.EnterDate, dbo.TblOrderMaint.EnterTime, dbo.TblOrderMaint.startmaintenanceDate,"
        StrSQL = StrSQL & "                    dbo.FixedAssets.code, dbo.TblOrderMaint.EquepID, dbo.TblCarsData.Fullcode AS CarFullcode, dbo.TblCarsData.Model, dbo.TblCarsData.CarsTypeId,"
        StrSQL = StrSQL & "                    dbo.TBLCarTypes.name AS CrTypname, dbo.TBLCarTypes.namee AS CrTypnameE, dbo.TblCarsData.VModel, dbo.TblCarModels.Model AS ModelName,"
        StrSQL = StrSQL & "                    dbo.TblCarModels.ModelE AS ModelNameE , dbo.GetCarsExpensValue(dbo.TblOrderMaint.EquepID) AS OtherValue,dbo.TblMaintenanceType.id AS MainID, dbo.TblOrderMaint.LeaderName, "
        If CHKiSoUT.value = vbChecked Then
              StrSQL = StrSQL & "                    td.OperPrice,"
              StrSQL = StrSQL & "                    td.Item_ID, ti.ItemCode, ti.ItemName, ti.ItemNamee,ti.itemSerials,t.Transaction_Date,"
              StrSQL = StrSQL & "                    td.showPrice, td.ShowQty, td.ID AS TrnsID"
        Else
              StrSQL = StrSQL & "                    dbo.Transaction_Details.OperPrice,dbo.Transactions.Transaction_Date,"
              StrSQL = StrSQL & "                    dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,TblItems.itemSerials,"
              StrSQL = StrSQL & "                    dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.ID AS TrnsID "
        End If
        StrSQL = StrSQL & "    FROM         dbo.TblCarModels RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblCarsData ON dbo.TblCarModels.Id = dbo.TblCarsData.VModel LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_4 ON dbo.TblOrderMaint.DrievID = TblEmployee_4.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblMaintenanceType RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_5 RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmpDepartments RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblItems RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.Transaction_Details LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID ON"
        StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Transaction_IDDet = dbo.Transaction_Details.ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.tblordermaintenancetypes.CusID = dbo.TblCustemers.CusID ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
        StrSQL = StrSQL & "                    dbo.TblEmpDepartments.DeparmentID = dbo.tblordermaintenancetypes.DeptID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_6 ON dbo.tblordermaintenancetypes.SuperID = TblEmployee_6.Emp_ID ON"
        StrSQL = StrSQL & "                    TblEmployee_5.Emp_ID = dbo.tblordermaintenancetypes.EmpID ON dbo.TblMaintenanceType.id = dbo.tblordermaintenancetypes.maintenanceid ON"
        StrSQL = StrSQL & "                    dbo.TblOrderMaint.ID = dbo.tblordermaintenancetypes.ORderID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.TblOrderMaint.SuperVisor = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.reciverid = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_3 ON dbo.TblOrderMaint.LeaderID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblBranchesData TblBranchesData_1 ON dbo.TblOrderMaint.DcbBranchFrom = TblBranchesData_1.branch_id ON"
        StrSQL = StrSQL & "                    dbo.FixedAssets.id = dbo.TblOrderMaint.EquepID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                    dbo.TblBranchesData TblBranchesData_2 ON dbo.TblOrderMaint.BranchID = TblBranchesData_2.branch_id"
        If CHKiSoUT.value = vbChecked Then
              StrSQL = StrSQL & "                       LEFT OUTER JOIN Transactions AS t ON T.order_no = TblOrderMaint.ID"
              StrSQL = StrSQL & "      LEFT OUTER JOIN Transaction_Details AS td ON T.Transaction_ID= Td.Transaction_ID"
              StrSQL = StrSQL & " LEFT OUTER JOIN TblItems AS ti ON ti.ItemID = td.Item_ID"
              
              
        End If
    End If
  StrSQL = StrSQL & " Where (dbo.TblOrderMaint.EquepID <> 0)"
If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.SuperVisor = " & val(Me.DcbEmp.BoundText)
End If
If Me.DcbTypeMain.Text <> "" And val(Me.DcbTypeMain.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.tblordermaintenancetypes.maintenanceid = " & val(Me.DcbTypeMain.BoundText)
   
End If
If CHKiSoUT.value = vbChecked Then
   StrSQL = StrSQL & " AND  iSnULL(T.order_no,0) <> 0 "
Else
End If

If Trim(TXTOrderMaintID2) <> "" Then
    StrSQL = StrSQL & " AND dbo.TblOrderMaint.ID=" & val(TXTOrderMaintID2)
End If

If Me.TXTOrderMaintID <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.id  in ( " & Me.TXTOrderMaintID.Text & ")"
End If

'If Trim(TXTOrderMaintID2) <> "" Then
'    StrSQL = StrSQL & " AND dbo.TblRequerMainten.ID=" & val(TXTOrderMaintID2)
'End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
If Me.DcbStutsMaint2.ListIndex <> -1 And Me.DcbStutsMaint2.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.StutsMaint = " & val(Me.DcbStutsMaint2.ListIndex)
        End If
        
       
        
If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.BranchID = " & val(Me.DcbBranch.BoundText)
End If
If Me.DcbLeader.Text <> "" And val(DcbLeader.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.LeaderID = " & val(Me.DcbLeader.BoundText)
End If


If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblOrderMaint.EquepID = " & val(Me.DcbEqup.BoundText)

End If


 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL, , , isTotal

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub


Public Sub GetDataAlarm()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

    StrSQL = " SELECT  te.Emp_Name,"
    If optBydate(0) Then
        StrSQL = StrSQL & "    FixedAssets.Name Car,"
        
        StrSQL = StrSQL & "    STATUS =  ISNULL(TblOrderMaint.StutsMaint, 0)"
    Else
        StrSQL = StrSQL & "    FFA.Name Car,"
        'StrSQL = StrSQL & "    STATUS = (CASE MaintPlan WHEN ISNULL(MaintPlan, 0) THEN 1 ELSE 0 END)"
        StrSQL = StrSQL & "    STATUS = TblOrderMaint.StutsMaint"
        
        
    End If
    StrSQL = StrSQL & "    ,TTType.name MaintType,"
    
    If optBydate(1) Then
        StrSQL = StrSQL & "    Pl.Remarks,tcmpd.Done,"
        StrSQL = StrSQL & "    Pl.RecordDate                AS PlanDate,"
    Else
        StrSQL = StrSQL & "    TblRequerMainten.Remarks,"
        StrSQL = StrSQL & "    TblRequerMainten.RecordDate                AS PlanDate,"
    
    End If
    StrSQL = StrSQL & "        TblOrderMaint.RecordDate,"
    
    If optBydate(1) Then
        StrSQL = StrSQL & "    Pl.Planid    MaintPlan,"
    Else
        StrSQL = StrSQL & "      TblOrderMaint.reqmainID as   MaintPlan,"
    End If
    StrSQL = StrSQL & "        LeaderName = ("
    StrSQL = StrSQL & "            CASE"
    StrSQL = StrSQL & "                 WHEN ISNULL(TblOrderMaint.LeaderName, '') = '' THEN te.Emp_Name"
    StrSQL = StrSQL & "                 Else TblOrderMaint.LeaderName"
    StrSQL = StrSQL & "            End"
    StrSQL = StrSQL & "        ),"
    StrSQL = StrSQL & "        TblOrderMaint.LeaderID,tcmpd.CancelReason "
       
       
    StrSQL = StrSQL & " From TblOrderMaint"
    StrSQL = StrSQL & "        LEFT OUTER JOIN TblEmployee  AS te"
    StrSQL = StrSQL & "             ON  te.Emp_ID = TblOrderMaint.LeaderID"
    If optBydate(1) Then
        StrSQL = StrSQL & "        Right Outer JOIN TblCarMaintenancePlan Pl"
        StrSQL = StrSQL & "             ON  Pl.Planid = MaintPlan"
    Else
        StrSQL = StrSQL & "        Left Outer JOIN TblCarMaintenancePlan Pl"
        StrSQL = StrSQL & "             ON  Pl.Planid = MaintPlan"
        StrSQL = StrSQL & "        Left Outer JOIN TblRequerMainten On TblRequerMainten.Id =  TblOrderMaint.reqmainID"
    End If
    
    StrSQL = StrSQL & "             LEFT OUTER JOIN FixedAssets ON FixedAssets.id = dbo.TblOrderMaint.EquepID"
    
    StrSQL = StrSQL & "             LEFT OUTER JOIN TblCarsData ON TblCarsData.id = Pl.CarId"
    StrSQL = StrSQL & "             LEFT OUTER JOIN FixedAssets FFA ON TblCarsData.fixedAssetid = FFA.ID"
    
    
    
            

    StrSQL = StrSQL & "             Left OUTER JOIN TblCarMaintenancePlanDetails AS tcmpd"
    StrSQL = StrSQL & "                         ON  Pl.PlanId= tcmpd.Planid"
    
    StrSQL = StrSQL & "                         Left OUTER JOIN TblMaintenanceType AS TTType"
    StrSQL = StrSQL & "                         ON  tcmpd.MaintenanceID= TTType.id"
    

  StrSQL = StrSQL & " Where  1 = 1"
  '(dbo.TblOrderMaint.EquepID <> 0)"
  
    
    If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblOrderMaint.SuperVisor = " & val(Me.DcbEmp.BoundText)
    End If
    If Me.DcbTypeMain.Text <> "" And val(Me.DcbTypeMain.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblOrderMaint.TypeMaint = " & val(Me.DcbTypeMain.BoundText)
    End If
    If Me.DcbStutsMaint2.ListIndex <> -1 And Me.DcbStutsMaint2.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.StutsMaint = " & val(Me.DcbStutsMaint2.ListIndex)
        End If
    If optBydate(0) Then
        If Not IsNull(Me.DtpDateFrom.value) Then
                          StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
             End If
              If Not IsNull(Me.DtpDateTo.value) Then
                          StrSQL = StrSQL & " AND dbo.TblOrderMaint.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
             End If
    Else
        If Not IsNull(Me.DtpDateFrom.value) Then
                          StrSQL = StrSQL & " AND Pl.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
             End If
              If Not IsNull(Me.DtpDateTo.value) Then
                          StrSQL = StrSQL & " AND Pl.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
             End If
    
    End If
    'STATUS
    If DcbStutsMaint4.ListIndex <> -1 Then
        
         If optBydate(0) Then
           ' StrSQL = StrSQL & " AND  ISNULL(TblOrderMaint.StutsMaint, 0) =  " & val(DcbStutsMaint4.ListIndex)
            
           
        Else
            
            
            If Me.DcbStutsMaint4.ListIndex = 1 Then
               StrSQL = StrSQL & " AND    IsNull(TblOrderMaint.StutsMaint, 0) = 2 And IsNull(tcmpd.Done,0) = 0"
            ElseIf Me.DcbStutsMaint4.ListIndex = 0 Then
                StrSQL = StrSQL & " AND   (IsNull(tcmpd.Done,0) = " & 0 & "  and IsNull(TblOrderMaint.StutsMaint, 0) = 0)"
            ElseIf Me.DcbStutsMaint4.ListIndex = 2 Then
                StrSQL = StrSQL & " AND   IsNull(tcmpd.Done,0) = " & IIf(val(Me.DcbStutsMaint4.ListIndex) = 2, 1, 0) & "  "
            End If
            
     '       StrSQL = StrSQL & " AND   IsNull(tcmpd.Done,0) = " & IIf(val(Me.DcbStutsMaint4.ListIndex) = 2, 1, 0)
        
            'StrSQL = StrSQL & " AND   IsNull(TblOrderMaint.StutsMaint,0) = " & IIf(val(Me.DcbStutsMaint4.ListIndex) = 2, 1, 0)
        End If
    End If
    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblOrderMaint.BranchID = " & val(Me.DcbBranch.BoundText)
    End If
    If Me.DcbLeader.Text <> "" And val(DcbLeader.BoundText) <> 0 Then
        If optBydate(0) Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.LeaderID = " & val(Me.DcbLeader.BoundText)
        Else
            StrSQL = StrSQL & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbLeader.BoundText)
        End If
    
    End If
    
    If optBydate(0) Then
        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   dbo.TblOrderMaint.EquepID = " & val(Me.DcbEqup.BoundText)
        
        End If
    Else
        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   TblCarsData.fixedAssetid= " & val(Me.DcbEqup.BoundText)
        
        End If
    
    End If
'    If chkHidden.value = vbChecked Then
'        If optBydate(0) Then
'            StrSQL = StrSQL & " AND  ISNULL(TblRequerMainten.StatusMaint, 0) <> 2"
'        Else
'            StrSQL = StrSQL & " AND  ISNULL(TblOrderMaint.StutsMaint, 0) <> 2"
'
'        End If
'    End If
    
    StrSQL = StrSQL & " Order By Pl.Planid"

 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL, 1

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub



Public Sub GetDataCover()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
Dim Sql2 As String

StrSQL = " SELECT FixesAssetsID,"
 StrSQL = StrSQL & "       td.ItemSerial NoteSerial1,"
  StrSQL = StrSQL & "      TblCarsData.Id                     CarID,TblCarsData.OwnerName,"
 StrSQL = StrSQL & "       TblCarsData.Emp_id,"
 StrSQL = StrSQL & "        te.Emp_Name,"
 StrSQL = StrSQL & "       ManualNO,"
 StrSQL = StrSQL & "       t.Transaction_Date,"
 StrSQL = StrSQL & "       ti.ItemID,"
 StrSQL = StrSQL & "       ti.ItemName,"
 StrSQL = StrSQL & "       Groups.GroupName,"
 StrSQL = StrSQL & "       fa.Name                            faName,"
 StrSQL = StrSQL & "       ti.ItemCase"
 StrSQL = StrSQL & " FROM   Transactions                    AS t"
 StrSQL = StrSQL & "       INNER JOIN Transaction_Details  AS td"
 StrSQL = StrSQL & "            ON  td.Transaction_ID = t.Transaction_ID"
 StrSQL = StrSQL & "       LEFT OUTER JOIN TblItems        AS ti"
 StrSQL = StrSQL & "            ON  ti.ItemID = td.Item_ID"
 StrSQL = StrSQL & "       LEFT OUTER JOIN Groups"
 StrSQL = StrSQL & "            ON  ti.GroupID = Groups.GroupID"
 StrSQL = StrSQL & "       LEFT OUTER JOIN FixedAssets     AS fa"
 StrSQL = StrSQL & "            ON  fa.id = t.FixesAssetsID"
 StrSQL = StrSQL & "       LEFT OUTER JOIN TblCarsData"
 StrSQL = StrSQL & "            ON  TblCarsData.fixedAssetid = t.FixesAssetsID"
 StrSQL = StrSQL & "       LEFT OUTER JOIN TblEmployee     AS te"
 StrSQL = StrSQL & "            ON  te.Emp_id = TblCarsData.Emp_id"
 StrSQL = StrSQL & " Where t.Transaction_Type = 19"


    


  '(dbo.TblOrderMaint.EquepID <> 0)"
    If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbEmp.BoundText)
    End If


        If Not IsNull(Me.DtpDateFrom.value) Then
                          StrSQL = StrSQL & " AND t.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
             End If
              If Not IsNull(Me.DtpDateTo.value) Then
                          StrSQL = StrSQL & " AND t.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
             End If
    


    
    If Me.cmbCoverStatus.ListIndex <> -1 And Me.cmbCoverStatus.ListIndex <> 3 Then
            StrSQL = StrSQL & " AND   ti.ItemCase = " & val(Me.cmbCoverStatus.ListIndex) + 1
        End If
        
'    mWhere = mWhere & "  1 = 0)"

    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   t.BranchID = " & val(Me.DcbBranch.BoundText)
    End If
    If Me.DcbLeader.Text <> "" And val(DcbLeader.BoundText) <> 0 Then
        If optBydate(0) Then
            StrSQL = StrSQL & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbLeader.BoundText)
        Else
            StrSQL = StrSQL & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbLeader.BoundText)
        End If
    
    End If
        If DCGroups3.Text <> "" Then
            StrSQL = StrSQL & " AND   Groups.GroupID = " & val(Me.DCGroups3.BoundText)
    End If
    If optBydate(0) Then
        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   TblCarsData.Id = " & val(Me.DcbEqup.BoundText)
        
        End If
    Else
        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
            StrSQL = StrSQL & " AND   TblCarsData.fixedAssetid= " & val(Me.DcbEqup.BoundText)
        
        End If
    
    End If
   'StrSQL = StrSQL & " Order By Pl.Planid"

 
 
 Sql2 = " SELECT distinct "
 Sql2 = Sql2 & "       td.ItemSerial NoteSerial1,"
  Sql2 = Sql2 & "      Fa.EqupID                  CarID,"
 Sql2 = Sql2 & "       TblCarsData.Emp_id,"
 Sql2 = Sql2 & "        te.Emp_Name,"
 Sql2 = Sql2 & "       ManualNO,"
 Sql2 = Sql2 & "       t.Transaction_Date,"
 Sql2 = Sql2 & "       ti.ItemID,"
 Sql2 = Sql2 & "       ti.ItemName,"
 Sql2 = Sql2 & "       Groups.GroupName,"
 
 Sql2 = Sql2 & "       ti.ItemCase"
 Sql2 = Sql2 & " FROM   Transactions                    AS t"
 Sql2 = Sql2 & "       INNER JOIN Transaction_Details  AS td"
 Sql2 = Sql2 & "            ON  td.Transaction_ID = t.Transaction_ID"
 Sql2 = Sql2 & "       LEFT OUTER JOIN TblItems        AS ti"
 Sql2 = Sql2 & "            ON  ti.ItemID = td.Item_ID"
 Sql2 = Sql2 & "       LEFT OUTER JOIN Groups"
 Sql2 = Sql2 & "            ON  ti.GroupID = Groups.GroupID"
 
Sql2 = Sql2 & "                   LEFT OUTER JOIN TblCarsDataDet     AS fa"
Sql2 = Sql2 & "                        ON  fa.PartID = t.FixesAssetsID"
            

 Sql2 = Sql2 & "       LEFT OUTER JOIN TblCarsData"
 Sql2 = Sql2 & "            ON  TblCarsData.fixedAssetid = t.FixesAssetsID"
 Sql2 = Sql2 & "       LEFT OUTER JOIN TblEmployee     AS te"
 Sql2 = Sql2 & "            ON  te.Emp_id = TblCarsData.Emp_id"
 Sql2 = Sql2 & " Where t.Transaction_Type = 19"


    
    If Me.cmbCoverStatus.ListIndex <> -1 And Me.cmbCoverStatus.ListIndex <> 3 Then
            Sql2 = Sql2 & " AND   ti.ItemCase = " & val(Me.cmbCoverStatus.ListIndex)
        End If

  '(dbo.TblOrderMaint.EquepID <> 0)"
'    If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
'        Sql2 = Sql2 & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbEmp.BoundText)
'    End If
'
'
'        If Not IsNull(Me.DtpDateFrom.value) Then
'                          Sql2 = Sql2 & " AND t.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
'             End If
'              If Not IsNull(Me.DtpDateTo.value) Then
'                          Sql2 = Sql2 & " AND t.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
'             End If
'
'
'    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
'        Sql2 = Sql2 & " AND   t.BranchID = " & val(Me.DcbBranch.BoundText)
'    End If
'    If Me.DcbLeader.Text <> "" And val(DcbLeader.BoundText) <> 0 Then
'        If optBydate(0) Then
'            Sql2 = Sql2 & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbLeader.BoundText)
'        Else
'            Sql2 = Sql2 & " AND   dbo.TblCarsData.Emp_id = " & val(Me.DcbLeader.BoundText)
'        End If
'
'    End If
'    If DCGroups3.Text <> "" Then
'            Sql2 = Sql2 & " AND   Groups.GroupID = " & val(Me.DCGroups3.BoundText)
'    End If
'    If optBydate(0) Then
'        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
'            Sql2 = Sql2 & " AND   TblCarsData.Id = " & val(Me.DcbEqup.BoundText)
'
'        End If
'    Else
'        If Me.DcbEqup.Text <> "" And val(DcbEqup.BoundText) <> 0 Then
'            Sql2 = Sql2 & " AND   TblCarsData.fixedAssetid= " & val(Me.DcbEqup.BoundText)
'
'        End If
'
'    End If


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Set rs2 = New ADODB.Recordset
    rs2.Open Sql2, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF And rs2.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
' rs.MoveFirst
 print_report StrSQL, 6, Sql2

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub

Function print_report(Optional NoteSerial As String, Optional ByVal mType As Long = 0, Optional ByVal Sql2 As String = "", Optional ByVal isTotal As Boolean = False)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If Me.CHReq.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceReq.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceReqE.rpt"
            
       End If
       End If
       If Me.Chorder.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrder.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrderE.rpt"
            
       End If

End If
       If Me.ChorderAnlys.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrderAnalysor.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrderAnalysorE.rpt"
          End If
          End If

 If Me.ChCarExpen.value = True Then
        
            If Not isTotal Then
                  If SystemOptions.UserInterface = ArabicInterface Then
             
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpenses.rpt"
                 Else
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpenses.rpt"
            End If
      
        Else
              If SystemOptions.UserInterface = ArabicInterface Then
             
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpensesTotal.rpt"
                 Else
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpensesTotal.rpt"
            End If
     
        End If
End If
If Me.optAlarm.value = True Then
        
        If optBydate(1) Then
             If SystemOptions.UserInterface = ArabicInterface Then
             
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "OrderMaint.rpt"
                 Else
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "OrderMaint.rpt"
            End If
        Else
             If SystemOptions.UserInterface = ArabicInterface Then
             
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "OrderMaintByRequest.rpt"
                 Else
                 StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "OrderMaintByRequest.rpt"
            End If
        End If
End If
If optCover Then
    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Cover.rpt"
    
    
   
End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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


If optCover Then
    
    
    
    If Sql2 <> "" Then
        Dim RsData2  As New ADODB.Recordset
        
         
        RsData2.Open Sql2, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("Sub2").Database.SetDataSource RsData2
    End If
End If
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
   If Not IsNull(DtpDateFrom.value) And Not IsNull(DtpDateTo.value) Then
'    xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value
'    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
    End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial
    RsData.Close
    
    Set RsData = Nothing
     Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_report3(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
    If Me.ReqMaint.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceReq2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceReq2E.rpt"
        End If
   End If
   
   If Emp.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceDriver.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceDriverE.rpt"
        End If
    End If
   
    If Order.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrder2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrder2E.rpt"
        End If
    End If
  
  
      If Order2.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrder3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinanceOrder3.rpt"
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   
    If Not IsNull(DtpDateFrom2.value) Then
        xReport.ParameterFields(8).AddCurrentValue DtpDateFrom2.value
    End If
    
    If Not IsNull(DtpDateTo2.value) Then
        xReport.ParameterFields(10).AddCurrentValue DtpDateTo2.value
    End If
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If chkTafweed.value = vbUnchecked Then

    MySQL = " SELECT  TblCarsData.OwnerName ,   dbo.TblCarsData.CarsTypeId, dbo.TblBranchesData.branch_id, dbo.TblCarsData.Emp_id, dbo.TblCarsData.VColor, dbo.TblCarsData.LocationID, "
    MySQL = MySQL & "                      dbo.TblCarsData.VModel, dbo.TblCarsData.id, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.LicenseNO,"
    MySQL = MySQL & "                      dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter,"
    MySQL = MySQL & "                      dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.InsuranceCompanyId, dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate,"
    MySQL = MySQL & "                      dbo.TblCarsData.Notes, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH,"
    MySQL = MySQL & "                      dbo.TblCarsData.fixedAssetid, dbo.TblCarsData.VehicleLong, dbo.TblCarsData.EquQty, dbo.TblCarsData.Capacity, dbo.TblCarsData.ContractID,"
    MySQL = MySQL & "                      dbo.TblCarsData.EndContractDate, dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.Rep,"
    MySQL = MySQL & "                      dbo.TblCarsData.EndAllocationDate, dbo.TblCarsData.MaxCap, dbo.TblCarsData.OperatorN, dbo.TblCarsData.EqupName, dbo.TblCarsData.TypeCar,"
    MySQL = MySQL & "                      dbo.TblCarsData.Gearno, dbo.TblCarsData.Gearno1, dbo.TblCarsData.Machineno, dbo.TblCarsData.Machineno1, dbo.TblCarsData.VType, dbo.TblCarsData.Chesis,"
    MySQL = MySQL & "                      dbo.TblCarsData.Total, dbo.TblCarsData.LetterCount, dbo.TblCarsData.LetterPrice, dbo.TblCarsData.FormOrignal, dbo.TblCarsData.authorizeLicense,"
    MySQL = MySQL & "                      dbo.TblCarsData.authorizeExamination, dbo.TblCarsData.cleaner, dbo.TblCarsData.sideMirror, dbo.TblCarsData.driverMirror, dbo.TblCarsData.InnerLights,"
    MySQL = MySQL & "                      dbo.TblCarsData.Pedals, dbo.TblCarsData.Recorder, dbo.TblCarsData.SunScreens, dbo.TblCarsData.Anntena, dbo.TblCarsData.Battery, dbo.TblCarsData.SpareTyre,"
    MySQL = MySQL & "                      dbo.TblCarsData.Crane, dbo.TblCarsData.CoverKey, dbo.TblCarsData.Guarantee, dbo.TblCarsData.Stickers, dbo.TblColor.name AS ColorName,"
    MySQL = MySQL & "                      dbo.TblCarModels.Model AS ModelName, dbo.EmpGroupDep.GroupName, dbo.TblBranchesData.branch_name, dbo.TBLCarTypes.name AS TypeName,"
    MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Emp_FullCode, dbo.FixedAssets.Name AS FixedAssetName, dbo.TblCarsData.code,"
    MySQL = MySQL & "                      dbo.TblCarsData.StutsID, dbo.TblCarsData.Job, dbo.TblCarsData.Natinality, dbo.TblCarsData.Department, dbo.TblCarsData.DriLicenseNo,"
    MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblCarsData.EmpType, dbo.TblCarsDataDet.PartID,"
    MySQL = MySQL & "                      FixedAssets_1.code AS Partcode, FixedAssets_1.Name AS PartName, FixedAssets_1.namee AS PartNameE"
    MySQL = MySQL & " FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblCarsDataDet.PartID = FixedAssets_1.id RIGHT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblCarsData INNER JOIN"
    MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id ON dbo.TblCarsDataDet.EqupID = dbo.TblCarsData.fixedAssetid LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID"
    MySQL = MySQL & "  WHERE 1 = 1 "
    
    
        If DCOwner.Text <> "" Then
            MySQL = MySQL & " AND  TblCarsData.OwnerName ='" & (DCOwner.Text) & "'"
    End If
    
    If DcbStuts.Text <> "" And val(DcbStuts.ListIndex) <> -1 Then
            MySQL = MySQL & " AND  dbo.TblCarsData.StutsID = " & val(DcbStuts.ListIndex)
    End If
    
    If DCGroup.BoundText <> "" Then
            MySQL = MySQL & " AND  TblCarsData.CarsTypeId = " & val(DCGroup.BoundText)
    End If
    If LocationID.Text <> "" And val(LocationID.BoundText) <> 0 Then
            MySQL = MySQL & " AND  TblCarsData.LocationID = " & val(LocationID.BoundText)
    End If
    
    If DcEmployee.BoundText <> "" Then
            MySQL = MySQL & " AND  TblCarsData.Emp_id = " & val(DcEmployee.BoundText)
    End If
    
      If Trim(DcbKafelName.Text) <> "" Then
            MySQL = MySQL & " AND  TblEmployee.KafelName= N'" & Trim(DcbKafelName.Text) & "'"
    End If
    
    If Trim(DcbDept3.Text) <> "" Then
        MySQL = MySQL & " AND TblEmployee.DepartmentID = " & val(DcbDept3.BoundText)
    End If
    
    
    If Trim(DcbDepartment2.Text) <> "" Then
        MySQL = MySQL & " AND TblEmployee.DeptID2 = " & val(DcbDepartment2.BoundText)
    End If
    
          If DcbEqup4.BoundText <> "" Then
            MySQL = MySQL & " AND  FixedAssets.id= " & val(DcbEqup4.BoundText)
    End If
    
      
    
    If txtModel.Text <> "" Then
            MySQL = MySQL & " AND TblCarsData.Model LIKE '%" & txtModel.Text & "%'"
    End If
    If Not IsNull(FromDate.value) Then
    MySQL = MySQL & " AND  TblCarsData.LicenseExpireDate <= " & SQLDate(FromDate.value, True) & ""
    End If
    If Not IsNull(Me.ToDate.value) Then
    MySQL = MySQL & " AND  TblCarsData.LicenseExpireDate >= " & SQLDate(ToDate.value, True) & ""
    End If
    MySQL = MySQL & " order by  TblCarsData.ID "
Else


    MySQL = " SELECT     dbo.TblEmpPassOver2.AdvanceID,TblEmpPassOver2.IsAuthorization,TblEmpPassOver2.DeparmentID,TblEmpPassOver2.KafelName, dbo.TblEmpPassOver2.NoteSerial, dbo.TblEmpPassOver2.ComputerNo, dbo.TblEmpPassOver2.Name AS EmpName, "
    MySQL = MySQL & "                      dbo.TblEmpPassOver2.EmpType,dbo.TblEmpPassOver2.DateCancel, dbo.TblEmpPassOver2.DcbLeaderID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
    MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,TblEmpPassOver2.Name as Mofawad, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
    MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmpPassOver2.LeaderName,"
    MySQL = MySQL & "                      dbo.TblEmpPassOver2.Nationality, dbo.TblEmpPassOver2.NumID, dbo.TblEmpPassOver2.Remarks2, dbo.TblEmpPassOver2.BoardNO, dbo.TblEmpPassOver2.OperatorN,"
    MySQL = MySQL & "                      dbo.TblEmpPassOver2.ColorID, dbo.TblColor.name AS Colorname, dbo.TblColor.namee AS ColornameE, dbo.TblEmpPassOver2.ModelID, dbo.TblCarModels.Model,"
    MySQL = MySQL & "                      dbo.TblCarModels.ModelE, dbo.TblEmpPassOver2.TypeEqupID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblEmpPassOver2.EquepmentID,"
    MySQL = MySQL & "                      dbo.FixedAssets.Name AS EqupName, dbo.FixedAssets.namee AS EqupNameE, dbo.TblEmpPassOver2.AdvanceDate,"
    MySQL = MySQL & "                      TblEmpDepartments.DepartmentName,TblEmpDepartmentsDet.Name as Expr1"

    MySQL = MySQL & "                      From dbo.TblEmpPassOver2"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.FixedAssets"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.EquepmentID = dbo.FixedAssets.id"
    MySQL = MySQL & "                             LEFT OUTER JOIN TblCarsData"
    MySQL = MySQL & "                                  ON  TblCarsData.fixedAssetid = dbo.FixedAssets.id"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TBLCarTypes"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.TypeEqupID = dbo.TBLCarTypes.id"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TblCarModels"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.ModelID = dbo.TblCarModels.Id"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TblColor"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.ColorID = dbo.TblColor.Id"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TblEmployee"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.DcbLeaderID = dbo.TblEmployee.Emp_ID"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TblEmpDepartments"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.DeparmentID2 = dbo.TblEmpDepartments.DeparmentID"
    MySQL = MySQL & "                             LEFT OUTER JOIN dbo.TblEmpDepartmentsDet"
    MySQL = MySQL & "                                  ON  dbo.TblEmpPassOver2.DeptID2 = dbo.TblEmpDepartmentsDet.Id"

      
     
      
    MySQL = MySQL & " Where  1 = 1"


    
    
    If DCGroup.BoundText <> "" Then
            MySQL = MySQL & " AND  TblCarsData.CarsTypeId = " & val(DCGroup.BoundText)
    End If
    
    
    If DcbEqup4.BoundText <> "" Then
            MySQL = MySQL & " AND  FixedAssets.id= " & val(DcbEqup4.BoundText)
    End If
    
    If LocationID.Text <> "" And val(LocationID.BoundText) <> 0 Then
            MySQL = MySQL & " AND  TblCarsData.LocationID = " & val(LocationID.BoundText)
    End If
    
    If DcEmployee.BoundText <> "" Then
            MySQL = MySQL & " AND  TblCarsData.Emp_id = " & val(DcEmployee.BoundText)
    End If
    
    If Trim(DcbCarModel.Text) <> "" Then
            MySQL = MySQL & " AND TblEmpPassOver2.ModelID " & val(DcbCarModel.BoundText)
    End If
    If Not IsNull(FromDate.value) Then
        MySQL = MySQL & " AND  TblEmpPassOver2.AdvanceDate >= " & SQLDate(FromDate.value, True) & ""
    End If
    If Not IsNull(Me.ToDate.value) Then
        MySQL = MySQL & " AND  TblEmpPassOver2.AdvanceDate <= " & SQLDate(ToDate.value, True) & ""
    End If
    
    If Trim(DcbKafelName.Text) <> "" Then
            MySQL = MySQL & " AND  TblEmpPassOver2.KafelName= N'" & Trim(DcbKafelName.Text) & "'"
    End If
    
    If Trim(DcbDept3.Text) <> "" Then
        MySQL = MySQL & " AND TblEmpPassOver2.DeparmentID2 = " & val(DcbDept3.BoundText)
    End If
    
    
    If Trim(DcbDepartment2.Text) <> "" Then
        MySQL = MySQL & " AND TblEmpPassOver2.DeptID2 = " & val(DcbDepartment2.BoundText)
    End If
    
    
    If optAuthorization.value Then
        MySQL = MySQL & " AND IsNull(TblEmpPassOver2.IsAuthorization,0) = 1"
    ElseIf optAuthorization.value Then
        MySQL = MySQL & " AND IsNull(TblEmpPassOver2.IsAuthorization,0) = 0"
    ElseIf optAlll.value Then
    
    End If
    MySQL = MySQL & " order by  TblCarsData.ID "
End If
If chkTafweed.value = vbUnchecked Then
    If CheckBox1.value = vbChecked Then
     If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsReoprt2.rpt"
         Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsReoprt2E.rpt"
           End If
     Else
      If SystemOptions.UserInterface = ArabicInterface Then
              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsReoprt.rpt"
         Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsReoprtE.rpt"
           End If
     End If
    Else
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpPassOver.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpPassOver.rpt"
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    Dim I As Long
      For I = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(I).Name
        Case "{@FromDate}"
            xReport.FormulaFields.Item(I).Text = CStr(FromDate.value)
        Case "{@ToDate}"
            xReport.FormulaFields.Item(I).Text = CStr(ToDate.value)
            
        End Select
    Next I
    
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
 
End Function

Private Sub TxtSearchCode2_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode2.Text, EmpID
        Me.DcbEmp2.BoundText = EmpID
    End If
End Sub

Function print_report4(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = " SELECT        dbo.ShiftRec.ID AS Expr3, dbo.ShiftRec.OrderMaintinNo,ShiftRec.StutsMaint, dbo.ShiftRec.CustTel, dbo.ShiftRec.ShiftMaintTypeID, dbo.ShiftRec.RecordDate, dbo.ShiftRec.BranchID AS Expr5, dbo.ShiftRec.typemaint AS Expr6, dbo.ShiftRec.DateRec,"
MySQL = MySQL & "                          dbo.ShiftRec.TimeRec, dbo.ShiftRec.Remarks, dbo.ShiftRec.NoteDone, dbo.ShiftRec.NoteStill, dbo.ShiftRec.NoteLate, dbo.ShiftRec.UserID, dbo.ShiftRec.TimeEnd, dbo.ShiftRec.DateEnd, dbo.ShiftRec.CarStatus,"
MySQL = MySQL & "                         TblOrderMaint.Des as InitialNotes, TblBranchesData_1.branch_name, dbo.TblBranchesData.branch_name AS branch_name2, dbo.TblOrderMaint.reciverRemarks AS reciverRemarks2, dbo.TblOrderMaint.DrievName AS DrievName2,"
MySQL = MySQL & "                          dbo.TblOrderMaint.EquepmentName AS EquepmentName2, dbo.TblOrderMaint.EquepID AS EquepID2, dbo.FixedAssets.Name, TblEmployee_1.Emp_Name, dbo.TblOrderMaint.LeaderName AS LeaderName2,"
MySQL = MySQL & "                          TblEmployee_2.Emp_Name AS Emp_Name2, dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate AS Expr1, dbo.TblOrderMaint.BranchID, dbo.TblOrderMaint.UserID AS Expr2, dbo.TblOrderMaint.EquepID,"
MySQL = MySQL & "                                 HoursM = DATEDIFF(hour, CAST(TimeRec AS DATETIME), CAST(TimeEnd AS DATETIME)),"
MySQL = MySQL & "                                 DaysMai = DATEDIFF(DAY, ShiftRec.DateRec, ShiftRec.DateEnd),"
MySQL = MySQL & "                                 TotalHours = DATEDIFF(hour, CAST(TimeRec AS DATETIME), CAST(TimeEnd AS DATETIME)) + DATEDIFF(Hour, ShiftRec.DateRec, ShiftRec.DateEnd),"

MySQL = MySQL & "                          dbo.TblOrderMaint.SuperVisor, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Remarks AS Expr4, dbo.TblOrderMaint.Jiha, dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.startmaintenanceTime,"
MySQL = MySQL & "                          dbo.TblOrderMaint.endmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceTime, dbo.TblOrderMaint.endmaintenanceDate, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks,"
MySQL = MySQL & "                          dbo.TblOrderMaint.reciverid, dbo.TblOrderMaint.ended, dbo.TblOrderMaint.ReqMainID, dbo.TblOrderMaint.TechNote, dbo.TblOrderMaint.DcbBranchFrom, dbo.TblOrderMaint.LeaderID, dbo.TblOrderMaint.LeaderType,"
MySQL = MySQL & "                          dbo.TblOrderMaint.LeaderName LeaderName6,TblEmployee_1.Emp_Name as LeaderName ,dbo.TblOrderMaint.DrievID, dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName, dbo.TblOrderMaint.EquepmentName, dbo.TblOrderMaint.Total,"
MySQL = MySQL & "                          dbo.TblOrderMaint.EnterDate, dbo.TblOrderMaint.EnterTime, dbo.TblOrderMaint.startmaintenanceDate, dbo.TblOrderMaint.BoardNO, dbo.TblOrderMaint.OperatorN, dbo.TblOrderMaint.TotalMaint, dbo.TblOrderMaint.TotalSpare,"
MySQL = MySQL & "                          dbo.TblOrderMaint.TotalSand, dbo.TblOrderMaint.DeptNotes, dbo.TblOrderMaint.InitialNotes ff, dbo.TblOrderMaint.CurrKM, dbo.TblOrderMaint.LastKM, dbo.TblOrderMaint.separatedreport, dbo.TblOrderMaint.separatedreport1,"
MySQL = MySQL & "                          dbo.TblOrderMaint.mangercomment, dbo.TblOrderMaint.alarms, dbo.TblOrderMaint.alarmsPeriod, dbo.TblOrderMaint.report1des, dbo.TblOrderMaint.report1des1, dbo.TblOrderMaint.carendperiod,"
MySQL = MySQL & "                          dbo.TblOrderMaint.carendperiod1, dbo.TblOrderMaint.MaintPlan, dbo.TblOrderMaint.BaisedOn, dbo.TblOrderMaint.MaintenanceTypesLineNo, dbo.ShiftMaintType.Name AS ShiftName,"
MySQL = MySQL & "                          dbo.TblEmployee.Emp_Name AS SuperVisorName"
MySQL = MySQL & "                          FROM            ShiftMaintType RIGHT OUTER JOIN"
MySQL = MySQL & "                                                   FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                                                   TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                                                   TblOrderMaint INNER JOIN"
MySQL = MySQL & "                                                   ShiftRec ON TblOrderMaint.ID = ShiftRec.OrderMaintinNo ON TblEmployee.Emp_ID = TblOrderMaint.SuperVisor LEFT OUTER JOIN"
MySQL = MySQL & "                                                   TblEmployee AS TblEmployee_2 ON TblOrderMaint.DrievID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                                                   TblEmployee AS TblEmployee_1 ON TblOrderMaint.LeaderID = TblEmployee_1.Emp_ID ON FixedAssets.id = TblOrderMaint.EquepID LEFT OUTER JOIN"
MySQL = MySQL & "                                                   TblBranchesData ON TblOrderMaint.DcbBranchFrom = TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                                                   TblBranchesData AS TblBranchesData_1 ON ShiftRec.BranchID = TblBranchesData_1.branch_id ON ShiftMaintType.ID = ShiftRec.ShiftMaintTypeID"

MySQL = MySQL & "  WHERE 1 = 1 "


If DcbStuts2.Text <> "" And val(DcbStuts2.ListIndex) <> -1 And val(DcbStuts2.ListIndex) <> 3 Then
        MySQL = MySQL & " AND  dbo.ShiftRec.CarStatus = " & val(DcbStuts2.ListIndex)
        StrReportTitle = " Õ«·… «·„⁄œÂ/«·”Ì«—… " & DcbStuts2.Text
End If
If Not IsNull(txtFromDate.value) Then
MySQL = MySQL & " AND  ShiftRec.RecordDate >= " & SQLDate(txtFromDate.value, True) & ""
End If
If Not IsNull(Me.txtToDate.value) Then
MySQL = MySQL & " AND  ShiftRec.RecordDate <= " & SQLDate(txtToDate.value, True) & ""
End If

If Me.DcbEqup3.Text <> "" And val(DcbEqup3.BoundText) <> 0 Then
    MySQL = MySQL & " AND   dbo.FixedAssets.id = " & val(Me.DcbEqup3.BoundText)

End If


If Trim(cmbShiftMaintType.Text) <> "" Then
    MySQL = MySQL & " AND  ShiftRec.ShiftMaintTypeID = " & val(cmbShiftMaintType.BoundText)
    StrReportTitle = StrReportTitle & " «·Ê—œÌ… " & cmbShiftMaintType.Text
End If

If Me.DcbStutsMaint3.ListIndex <> -1 And Me.DcbStutsMaint3.ListIndex <> 3 And Me.DcbStutsMaint3.Text <> "" Then
    MySQL = MySQL & " AND   dbo.ShiftRec.StutsMaint = " & val(Me.DcbStutsMaint3.ListIndex)
    StrReportTitle = StrReportTitle & " Õ«·… «·’Ì«‰…  " & DcbStutsMaint3.Text
End If

If Option3 Then
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec.rpt"
       End If
 ElseIf Option2 Then
  If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec2.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec2.rpt"
       End If
 ElseIf Option1 Then
  If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec3.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ShiftRec3.rpt"
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
       ' StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
      ' StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
      Dim I As Integer
      For I = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(I).Name
        Case "{@Title}"
            xReport.FormulaFields.Item(I).Text = "'" & Trim(StrReportTitle) & "'"
        Case "{@FromDate}"
            xReport.FormulaFields.Item(I).Text = "'" & Trim(txtFromDate.value) & "'"
        Case "{@ToDate}"
            xReport.FormulaFields.Item(I).Text = "'" & Trim(txtToDate.value) & "'"
        End Select
    Next I

 '   xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
 
End Function


