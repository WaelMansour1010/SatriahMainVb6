VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAccountingReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·‹‹‹Õ‹‹”‹‹«»‹« "
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   21465
   HelpContextID   =   470
   Icon            =   "FrmAccountReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   21465
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   10950
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   21465
      _cx             =   37862
      _cy             =   19315
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
      GridRows        =   2
      GridCols        =   7
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmAccountReport.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton CmdLoadTree 
         Caption         =   " Õ„Ì· «·œ·Ì· «·„Õ«”»Ì"
         Height          =   690
         Left            =   15825
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   10230
         Width           =   5610
      End
      Begin VB.TextBox txt_mod_flag 
         Alignment       =   1  'Right Justify
         Height          =   690
         Left            =   10215
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   10230
         Visible         =   0   'False
         Width           =   5595
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   690
         Left            =   30
         TabIndex        =   18
         Top             =   10230
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1217
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
         Height          =   10185
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   21405
         _cx             =   37756
         _cy             =   17965
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
         AutoSizeChildren=   7
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
         Begin VB.CommandButton cmdClear 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   8610
            Width           =   960
         End
         Begin VB.CommandButton cmdUnSelectAll 
            Caption         =   "UnSelectAll"
            Height          =   495
            Left            =   14310
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   8610
            Width           =   960
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Select All"
            Height          =   495
            Left            =   15300
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   8610
            Width           =   960
         End
         Begin VB.Frame Frame6 
            Caption         =   "«œŒ· ﬂÊœ «·Õ”«»"
            Height          =   990
            Left            =   11460
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   9075
            Visible         =   0   'False
            Width           =   9975
            Begin VB.CommandButton Command2 
               Caption         =   "„”Õ"
               Height          =   495
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox TxtAccountCode 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4800
               TabIndex        =   49
               Top             =   240
               Width           =   2775
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "*«ﬂ » ﬂÊœ «·Õ”«» À„ «÷€ÿ «‰ —  "
               ForeColor       =   &H00FF0000&
               Height          =   1215
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂÊœ «·Õ”«»"
               Height          =   255
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   240
               Width           =   975
            End
         End
         Begin C1SizerLibCtl.C1Tab MainTab 
            CausesValidation=   0   'False
            Height          =   10275
            Left            =   75
            TabIndex        =   2
            Top             =   -45
            Width           =   11370
            _cx             =   20055
            _cy             =   18124
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
               Height          =   9900
               Index           =   0
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   45
               Width           =   11280
               _cx             =   19897
               _cy             =   17463
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
                  Height          =   870
                  Index           =   3
                  Left            =   11145
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   3420
                  _cx             =   6033
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
                     Height          =   540
                     Left            =   705
                     RightToLeft     =   -1  'True
                     TabIndex        =   5
                     Top             =   420
                     Width           =   3600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00DF967A&
                     Caption         =   "„‰ ›÷·ﬂ √œŒ· ﬁÌ„… ≈Â·«ﬂ«  «·› —…"
                     Height          =   315
                     Index           =   3
                     Left            =   270
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   675
                     Width           =   4665
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   9780
                  Index           =   2
                  Left            =   75
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   11130
                  _cx             =   19632
                  _cy             =   17251
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
                  Begin VB.CheckBox chkIsBasicInvoice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "»«·›« Ê—… «·„»œ∆Ì…"
                     Height          =   405
                     Left            =   -30
                     RightToLeft     =   -1  'True
                     TabIndex        =   152
                     Top             =   870
                     Width           =   1350
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „Ã„⁄"
                     ForeColor       =   &H00800000&
                     Height          =   255
                     HelpContextID   =   520
                     Index           =   44
                     Left            =   7725
                     RightToLeft     =   -1  'True
                     TabIndex        =   151
                     Top             =   3150
                     Width           =   3165
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ«∆„… «·œŒ·   „’—Ê›«  „ﬁ«—‰… ‘Â—Ì…"
                     Height          =   270
                     Index           =   43
                     Left            =   510
                     RightToLeft     =   -1  'True
                     TabIndex        =   149
                     Top             =   3735
                     Width           =   3165
                  End
                  Begin VB.CheckBox Check2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„’—Ê›«  «·„⁄œ« "
                     Height          =   270
                     Left            =   7155
                     RightToLeft     =   -1  'True
                     TabIndex        =   144
                     Top             =   2850
                     Width           =   1650
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " Õ·Ì· Õ—ﬂ… Õ”«» „⁄Ì‰ ⁄·Ì «·„‘«—Ì⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   42
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   2130
                     Width           =   3435
                  End
                  Begin VB.CheckBox chKCCTotals 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã„«·Ì« "
                     Height          =   270
                     Left            =   210
                     RightToLeft     =   -1  'True
                     TabIndex        =   131
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1110
                  End
                  Begin VB.CheckBox chkIsAll 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·ﬂ·"
                     Height          =   195
                     Left            =   7260
                     RightToLeft     =   -1  'True
                     TabIndex        =   130
                     Top             =   420
                     Width           =   1005
                  End
                  Begin VB.Frame Frame5 
                     Height          =   360
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   126
                     Top             =   2775
                     Width           =   2055
                     Begin VB.OptionButton chREtype 
                        Alignment       =   1  'Right Justify
                        Caption         =   "«·ﬂ·"
                        Height          =   195
                        Index           =   2
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   129
                        Top             =   120
                        Value           =   -1  'True
                        Width           =   615
                     End
                     Begin VB.OptionButton chREtype 
                        Alignment       =   1  'Right Justify
                        Caption         =   "œ«∆‰"
                        Height          =   195
                        Index           =   1
                        Left            =   480
                        RightToLeft     =   -1  'True
                        TabIndex        =   128
                        Top             =   120
                        Width           =   735
                     End
                     Begin VB.OptionButton chREtype 
                        Alignment       =   1  'Right Justify
                        Caption         =   "„œÌ‰"
                        Height          =   195
                        Index           =   0
                        Left            =   1080
                        RightToLeft     =   -1  'True
                        TabIndex        =   127
                        Top             =   120
                        Width           =   735
                     End
                  End
                  Begin VB.CheckBox WithoutOpenenig 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ—ﬂ… ›ﬁÿ"
                     Height          =   255
                     Left            =   8295
                     RightToLeft     =   -1  'True
                     TabIndex        =   125
                     Top             =   960
                     Width           =   1140
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬁ«∆„Â «·œŒ· »«·„” ÊÌ« "
                     ForeColor       =   &H00800000&
                     Height          =   195
                     HelpContextID   =   520
                     Index           =   41
                     Left            =   165
                     RightToLeft     =   -1  'True
                     TabIndex        =   124
                     Top             =   3465
                     Width           =   3540
                  End
                  Begin VB.CheckBox Check1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã„«·Ì« "
                     Height          =   255
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   123
                     Top             =   2655
                     Width           =   1275
                  End
                  Begin VB.TextBox Txtyear 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   122
                     Top             =   9150
                     Width           =   855
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» «·⁄ﬁ«—"
                     ForeColor       =   &H00000000&
                     Height          =   315
                     Index           =   40
                     Left            =   345
                     RightToLeft     =   -1  'True
                     TabIndex        =   119
                     Top             =   2895
                     Width           =   3405
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹…  »«·„” ÊÌ« 2"
                     ForeColor       =   &H00800000&
                     Height          =   360
                     HelpContextID   =   520
                     Index           =   39
                     Left            =   6180
                     RightToLeft     =   -1  'True
                     TabIndex        =   118
                     Top             =   4005
                     Width           =   2355
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹… Õ”«» „⁄Ì‰2"
                     ForeColor       =   &H00800000&
                     Height          =   270
                     HelpContextID   =   520
                     Index           =   38
                     Left            =   345
                     RightToLeft     =   -1  'True
                     TabIndex        =   117
                     Top             =   3960
                     Width           =   3330
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» 2"
                     ForeColor       =   &H00FF0000&
                     Height          =   330
                     Index           =   37
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   116
                     Top             =   2775
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» «·⁄„·… «·«Ã‰»ÌÂ"
                     Height          =   285
                     HelpContextID   =   480
                     Index           =   36
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   2520
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹… 2"
                     ForeColor       =   &H00800000&
                     Height          =   270
                     HelpContextID   =   520
                     Index           =   35
                     Left            =   7725
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   4020
                     Width           =   3195
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— —»ÕÌ… «·„ÊŸ›Ì‰"
                     Height          =   195
                     Index           =   34
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   1980
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ì“«‰ ‘ﬂ· 2"
                     ForeColor       =   &H000000FF&
                     Height          =   270
                     Index           =   33
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   4050
                     Width           =   1260
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘Ê›«  Õ”«»« "
                     Height          =   330
                     Index           =   32
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   109
                     Top             =   2460
                     Visible         =   0   'False
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ«∆„… «·œŒ·   „ﬁ«—‰… ‘Â—Ì…"
                     Height          =   270
                     Index           =   29
                     Left            =   7725
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   3690
                     Width           =   3195
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ«∆„… «·œŒ· ·› —… „Õœœ…"
                     Height          =   270
                     Index           =   28
                     Left            =   4110
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   3465
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹…"
                     Height          =   135
                     HelpContextID   =   520
                     Index           =   5
                     Left            =   7830
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Top             =   3195
                     Visible         =   0   'False
                     Width           =   135
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹…  »«·„” ÊÌ« "
                     Height          =   390
                     HelpContextID   =   520
                     Index           =   18
                     Left            =   4110
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Top             =   3150
                     Visible         =   0   'False
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‹Ì‹“«‰ „‹—«Ã‹⁄‹…  Õ”«» „⁄Ì‰"
                     Height          =   390
                     HelpContextID   =   520
                     Index           =   25
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   3150
                     Visible         =   0   'False
                     Width           =   2865
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ«∆„… «·œŒ·   „Ã„⁄Â"
                     Height          =   270
                     Index           =   3
                     Left            =   7725
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   3465
                     Width           =   3195
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Ì‹“«‰‹Ì‹…  "
                     Height          =   270
                     Index           =   4
                     Left            =   4080
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   3735
                     Width           =   3000
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „ÊŸ›   Õ·Ì·Ì"
                     Height          =   285
                     Index           =   31
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   1755
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ·„‘—Ê⁄  «Ã„«·Ì"
                     Height          =   285
                     Index           =   30
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   1485
                     Width           =   3015
                  End
                  Begin VB.Frame Frame4 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Õœœ«  «· ﬁ—Ì—"
                     ForeColor       =   &H00FF0000&
                     Height          =   5535
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   4365
                     Width           =   11190
                     Begin VB.CommandButton Command7 
                        Height          =   255
                        Left            =   9780
                        RightToLeft     =   -1  'True
                        TabIndex        =   150
                        Top             =   5070
                        Width           =   735
                     End
                     Begin VB.CheckBox chk 
                        Alignment       =   1  'Right Justify
                        Caption         =   " ﬁ«—Ì— «·„‘«—Ì⁄ ·Õ”«» „⁄Ì‰ "
                        Height          =   315
                        Index           =   1
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   143
                        Top             =   690
                        Width           =   2715
                     End
                     Begin VB.TextBox TxtAccountCode2 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   210
                        TabIndex        =   140
                        Top             =   1170
                        Width           =   2385
                     End
                     Begin VB.CheckBox chk 
                        Alignment       =   1  'Right Justify
                        Caption         =   " ﬁ«—Ì— «·„‘«—Ì⁄ „’«—Ì› ›ﬁÿ"
                        Height          =   315
                        Index           =   0
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   139
                        Top             =   390
                        Width           =   2775
                     End
                     Begin VB.TextBox Text2 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   6840
                        TabIndex        =   138
                        Top             =   2220
                        Visible         =   0   'False
                        Width           =   900
                     End
                     Begin C1SizerLibCtl.C1Elastic Ele 
                        Height          =   675
                        Index           =   1
                        Left            =   2160
                        TabIndex        =   91
                        TabStop         =   0   'False
                        Top             =   4500
                        Visible         =   0   'False
                        Width           =   5550
                        _cx             =   9790
                        _cy             =   1191
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
                        ForeColor       =   16711680
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
                           Left            =   2250
                           TabIndex        =   92
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
                           Format          =   245104643
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
                           TabIndex        =   93
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
                           Format          =   245104643
                           CurrentDate     =   37357
                        End
                        Begin VB.Label lbl 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00E2E9E9&
                           Caption         =   "≈·Ï"
                           ForeColor       =   &H000000FF&
                           Height          =   285
                           Index           =   2
                           Left            =   1590
                           RightToLeft     =   -1  'True
                           TabIndex        =   95
                           Top             =   360
                           Width           =   555
                        End
                        Begin VB.Label lbl 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00E2E9E9&
                           Caption         =   "„‰"
                           ForeColor       =   &H000000FF&
                           Height          =   285
                           Index           =   4
                           Left            =   3750
                           RightToLeft     =   -1  'True
                           TabIndex        =   94
                           Top             =   285
                           Width           =   555
                        End
                     End
                     Begin VB.Frame FrameDateH 
                        Caption         =   " ÕœÌœ «· «—ÌŒ «·ÂÃ—Ì"
                        Height          =   705
                        Left            =   2520
                        RightToLeft     =   -1  'True
                        TabIndex        =   86
                        Top             =   4560
                        Visible         =   0   'False
                        Width           =   5190
                        Begin Dynamic_Byte.NourHijriCal Txt_DateHigriFrom 
                           Height          =   315
                           Left            =   2160
                           TabIndex        =   87
                           Top             =   240
                           Width           =   1455
                           _extentx        =   2566
                           _extenty        =   556
                        End
                        Begin Dynamic_Byte.NourHijriCal Txt_DateHigriTO 
                           Height          =   315
                           Left            =   120
                           TabIndex        =   88
                           Top             =   240
                           Width           =   1455
                           _extentx        =   2566
                           _extenty        =   556
                        End
                        Begin VB.Label Label10 
                           Alignment       =   1  'Right Justify
                           Caption         =   "„‰"
                           Height          =   255
                           Left            =   3600
                           RightToLeft     =   -1  'True
                           TabIndex        =   90
                           Top             =   240
                           Width           =   495
                        End
                        Begin VB.Label Label11 
                           Alignment       =   1  'Right Justify
                           Caption         =   "«·Ï"
                           Height          =   255
                           Left            =   1560
                           RightToLeft     =   -1  'True
                           TabIndex        =   89
                           Top             =   240
                           Width           =   495
                        End
                     End
                     Begin VB.TextBox TxtSearchCode 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   6780
                        TabIndex        =   76
                        Top             =   3930
                        Width           =   900
                     End
                     Begin VB.CheckBox ChkNotesType 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "·Õ—ﬂ… „⁄Ì‰…"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   66
                        Top             =   1140
                        Width           =   2685
                     End
                     Begin MSDataListLib.DataCombo DCActivity 
                        Bindings        =   "FrmAccountReport.frx":0411
                        Height          =   315
                        Left            =   2880
                        TabIndex        =   62
                        Top             =   150
                        Width           =   4605
                        _ExtentX        =   8123
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
                     Begin MSDataListLib.DataCombo dcBranch 
                        Bindings        =   "FrmAccountReport.frx":0426
                        Height          =   315
                        Left            =   2880
                        TabIndex        =   63
                        Top             =   480
                        Width           =   4605
                        _ExtentX        =   8123
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
                     Begin MSDataListLib.DataCombo DCNotesTypes 
                        Height          =   315
                        Left            =   2880
                        TabIndex        =   67
                        Top             =   1170
                        Width           =   4605
                        _ExtentX        =   8123
                        _ExtentY        =   556
                        _Version        =   393216
                        Enabled         =   0   'False
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCCompositeAccount 
                        Height          =   315
                        Left            =   2880
                        TabIndex        =   68
                        Top             =   1530
                        Width           =   4605
                        _ExtentX        =   8123
                        _ExtentY        =   556
                        _Version        =   393216
                        Enabled         =   0   'False
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcCostCenter 
                        Bindings        =   "FrmAccountReport.frx":043B
                        Height          =   315
                        Left            =   240
                        TabIndex        =   70
                        Top             =   1860
                        Width           =   7485
                        _ExtentX        =   13203
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
                        Bindings        =   "FrmAccountReport.frx":0450
                        Height          =   315
                        Left            =   240
                        TabIndex        =   72
                        Top             =   2220
                        Width           =   6645
                        _ExtentX        =   11721
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
                        Bindings        =   "FrmAccountReport.frx":0465
                        Height          =   315
                        Left            =   240
                        TabIndex        =   73
                        Top             =   2550
                        Width           =   7485
                        _ExtentX        =   13203
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
                     Begin MSDataListLib.DataCombo DCEmployee 
                        Height          =   315
                        Left            =   240
                        TabIndex        =   77
                        Top             =   3930
                        Width           =   5730
                        _ExtentX        =   10107
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
                     Begin MSDataListLib.DataCombo DcFixedAssets 
                        Height          =   315
                        Left            =   4320
                        TabIndex        =   80
                        Top             =   4260
                        Width           =   3405
                        _ExtentX        =   6006
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcboEmpDepartments 
                        Height          =   315
                        Left            =   240
                        TabIndex        =   82
                        Top             =   4230
                        Width           =   3045
                        _ExtentX        =   5371
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   16777215
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DataCombo1 
                        Bindings        =   "FrmAccountReport.frx":047A
                        Height          =   315
                        Left            =   -6000
                        TabIndex        =   84
                        Top             =   4920
                        Visible         =   0   'False
                        Width           =   6225
                        _ExtentX        =   10980
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
                     Begin ImpulseButton.ISButton CmdAccount 
                        Height          =   300
                        Left            =   510
                        TabIndex        =   96
                        Top             =   4950
                        Width           =   1530
                        _ExtentX        =   2699
                        _ExtentY        =   529
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
                        ButtonImage     =   "FrmAccountReport.frx":048F
                        ColorButton     =   14871017
                        ColorHoverText  =   16777215
                        DrawFocusRectangle=   0   'False
                        ColorToggledHoverText=   16777215
                     End
                     Begin MSDataListLib.DataCombo DCAccounts 
                        Bindings        =   "FrmAccountReport.frx":0829
                        Height          =   315
                        Left            =   240
                        TabIndex        =   98
                        Top             =   3600
                        Width           =   7485
                        _ExtentX        =   13203
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
                     Begin MSDataListLib.DataCombo DCRegionID 
                        Bindings        =   "FrmAccountReport.frx":083E
                        Height          =   315
                        Left            =   240
                        TabIndex        =   114
                        Top             =   150
                        Width           =   1845
                        _ExtentX        =   3254
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
                     Begin MSDataListLib.DataCombo DcbAqar 
                        Height          =   315
                        Left            =   2880
                        TabIndex        =   120
                        Top             =   840
                        Width           =   4605
                        _ExtentX        =   8123
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   16777215
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo cmbDataTypeExchange 
                        Height          =   315
                        Left            =   240
                        TabIndex        =   133
                        Top             =   3240
                        Width           =   7470
                        _ExtentX        =   13176
                        _ExtentY        =   556
                        _Version        =   393216
                        Style           =   2
                        BackColor       =   -2147483624
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcbProcess1 
                        Bindings        =   "FrmAccountReport.frx":0853
                        Height          =   315
                        Left            =   240
                        TabIndex        =   136
                        Top             =   2880
                        Width           =   7485
                        _ExtentX        =   13203
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
                     Begin MSDataListLib.DataCombo cmbAccount 
                        Bindings        =   "FrmAccountReport.frx":0868
                        Height          =   315
                        Left            =   210
                        TabIndex        =   141
                        Top             =   1470
                        Width           =   2385
                        _ExtentX        =   4207
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
                     Begin VB.Label Label25 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "·Õ”«» „⁄Ì‰"
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Left            =   870
                        RightToLeft     =   -1  'True
                        TabIndex        =   142
                        Top             =   960
                        Width           =   1335
                     End
                     Begin VB.Shape Shape2 
                        Height          =   1695
                        Left            =   8400
                        Top             =   2160
                        Width           =   1935
                     End
                     Begin VB.Label Label24 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·⁄„·Ì…"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7830
                        RightToLeft     =   -1  'True
                        TabIndex        =   137
                        Top             =   2940
                        Width           =   2385
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "‰Ê⁄ «·„’—Ê›"
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Index           =   5
                        Left            =   9180
                        TabIndex        =   134
                        Top             =   3270
                        Width           =   1140
                     End
                     Begin VB.Label Label23 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·⁄ﬁ«—"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7680
                        RightToLeft     =   -1  'True
                        TabIndex        =   121
                        Top             =   840
                        Width           =   2685
                     End
                     Begin VB.Label Label22 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„‰ÿﬁ…"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   2130
                        RightToLeft     =   -1  'True
                        TabIndex        =   115
                        Top             =   150
                        Width           =   645
                     End
                     Begin VB.Label Label19 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·„” √Ã—"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   600
                        RightToLeft     =   -1  'True
                        TabIndex        =   85
                        Top             =   4710
                        Visible         =   0   'False
                        Width           =   1455
                     End
                     Begin VB.Label Label18 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·«œ«—…"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   3210
                        RightToLeft     =   -1  'True
                        TabIndex        =   83
                        Top             =   4350
                        Width           =   1005
                     End
                     Begin VB.Label Label17 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·„⁄œ…"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   81
                        Top             =   4260
                        Width           =   2685
                     End
                     Begin VB.Label Label6 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "ﬂÊœ «·„ÊŸ›"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   79
                        Top             =   3930
                        Width           =   2565
                     End
                     Begin VB.Label Label9 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·«”„ "
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   5490
                        RightToLeft     =   -1  'True
                        TabIndex        =   78
                        Top             =   3930
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·»‰œ"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7830
                        RightToLeft     =   -1  'True
                        TabIndex        =   75
                        Top             =   2610
                        Width           =   2385
                     End
                     Begin VB.Label Label14 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·„‘—Ê⁄"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7830
                        RightToLeft     =   -1  'True
                        TabIndex        =   74
                        Top             =   2250
                        Width           =   2385
                     End
                     Begin VB.Label Label16 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ „—ﬂ“ «· ﬂ·›…  "
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   71
                        Top             =   1860
                        Width           =   2685
                     End
                     Begin VB.Label Label15 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·Õ”«» «·„Ã„⁄"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   69
                        Top             =   1530
                        Width           =   2685
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·‰‘«ÿ"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   65
                        Top             =   210
                        Width           =   2685
                     End
                     Begin VB.Label Label5 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·›—⁄"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7650
                        RightToLeft     =   -1  'True
                        TabIndex        =   64
                        Top             =   540
                        Width           =   2685
                     End
                     Begin VB.Label Label21 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "·Õ”«» „Õœœ"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   7530
                        RightToLeft     =   -1  'True
                        TabIndex        =   99
                        Top             =   3600
                        Width           =   2685
                     End
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«»  ›’Ì·Ì »«·ﬂ„Ì« "
                     Height          =   255
                     HelpContextID   =   480
                     Index           =   27
                     Left            =   1155
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   960
                     Width           =   2595
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» ⁄„Ì· ·„‘—Ê⁄ „⁄Ì‰"
                     ForeColor       =   &H00FF0000&
                     Height          =   330
                     Index           =   26
                     Left            =   7935
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   1830
                     Width           =   3030
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „ÊŸ› "
                     Height          =   255
                     Index           =   22
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   -1485
                     Visible         =   0   'False
                     Width           =   2970
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „” √Ã—"
                     Height          =   345
                     Index           =   24
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   360
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ··„ÊŸ›"
                     Height          =   330
                     Index           =   23
                     Left            =   8595
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   2310
                     Width           =   2370
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ·≈œ«—…"
                     Height          =   345
                     Index           =   21
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   1965
                     Width           =   3015
                  End
                  Begin VB.Frame Frame9 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„⁄œ… „Õœœ…"
                     Height          =   1335
                     Left            =   18720
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   4560
                     Visible         =   0   'False
                     Width           =   5805
                  End
                  Begin VB.Frame Frame8 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·«œ«—… „⁄Ì‰…"
                     Height          =   1170
                     Left            =   -1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   -2040
                     Visible         =   0   'False
                     Width           =   6645
                     Begin VB.Label Label12 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·«œ«—…"
                        Height          =   255
                        Left            =   3480
                        RightToLeft     =   -1  'True
                        TabIndex        =   52
                        Top             =   240
                        Width           =   1335
                     End
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „‘—Ê⁄ „Ã„⁄"
                     Height          =   210
                     Index           =   20
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   1425
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› „Ã„⁄ Õœœ «”„ «·Õ”«» «·„Ã„⁄"
                     Height          =   255
                     Index           =   19
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   960
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— „’—Ê›«  «·„⁄œ« "
                     Height          =   240
                     Index           =   17
                     Left            =   705
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   2295
                     Width           =   3030
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «—»«Õ «·„⁄œ« "
                     Height          =   255
                     Index           =   16
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   2190
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ··„⁄œ…"
                     ForeColor       =   &H00000000&
                     Height          =   345
                     Index           =   15
                     Left            =   7965
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   2580
                     Width           =   3000
                  End
                  Begin VB.Frame Frame7 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ÕœÌœ ”Ì«—…"
                     Height          =   1755
                     Left            =   11295
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   5580
                     Visible         =   0   'False
                     Width           =   5970
                     Begin MSDataListLib.DataCombo DCCar 
                        Bindings        =   "FrmAccountReport.frx":087D
                        Height          =   315
                        Left            =   0
                        TabIndex        =   37
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
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «⁄„«— «·œÌÊ‰ «Ã„«·Ì"
                     Height          =   255
                     HelpContextID   =   520
                     Index           =   14
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   1230
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „ÊŸ›  «Ã„«·Ì"
                     Height          =   195
                     Index           =   13
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   1755
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "√” «– ⁄«„ »«·Õ—ﬂ«  ·‹‹ ...."
                     Height          =   225
                     Index           =   12
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   1215
                     Width           =   3435
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "Command1"
                     Height          =   360
                     Left            =   570
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   -270
                     Visible         =   0   'False
                     Width           =   1530
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·«—’œ… «·«›  «ÕÌ…"
                     Height          =   225
                     Index           =   11
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   360
                     Width           =   3435
                  End
                  Begin VB.CheckBox chkContinue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘Ê›«  „ ’·Â"
                     Height          =   300
                     Left            =   165
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     Top             =   -495
                     Value           =   1  'Checked
                     Visible         =   0   'False
                     Width           =   2730
                  End
                  Begin VB.Frame Frame3 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „” √Ã—"
                     Height          =   705
                     Left            =   -7260
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   3105
                     Visible         =   0   'False
                     Width           =   5700
                  End
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ÕœÌœ „—ﬂ“ «· ﬂ·›…"
                     Height          =   1680
                     Left            =   11190
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   1995
                     Visible         =   0   'False
                     Width           =   5805
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ÕœÌœ „‘—Ê⁄"
                     Height          =   1665
                     Left            =   13620
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   2520
                     Visible         =   0   'False
                     Width           =   5670
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   825
                     Left            =   11190
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   25
                     Top             =   6375
                     Visible         =   0   'False
                     Width           =   9510
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ—ﬂ… ÿ»ﬁ« ·„‘—Ê⁄  Õ·Ì·Ì"
                     ForeColor       =   &H00FF0000&
                     Height          =   375
                     Index           =   10
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   1515
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«» „—ﬂ“  ﬂ·›…"
                     Height          =   285
                     Index           =   9
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   660
                     Width           =   3015
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»«⁄… «·ﬁÌÊœ «·ÌÊ„Ì… 'ÿ»ﬁ« ·„—«ﬂ“ «· ﬂ·›…"
                     Height          =   315
                     Index           =   8
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   660
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»«⁄… «·ﬁÌÊœ «·ÌÊ„Ì… "
                     ForeColor       =   &H00000000&
                     Height          =   180
                     Index           =   7
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   660
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "√” «– ⁄«„ »«·«—’œ…  ·‹‹ ...."
                     Height          =   330
                     Index           =   1
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   1230
                     Width           =   3465
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ‘› Õ”«»"
                     Height          =   255
                     HelpContextID   =   480
                     Index           =   0
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   960
                     Width           =   3435
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— „ «Ã—…"
                     Height          =   465
                     Index           =   2
                     Left            =   -2490
                     RightToLeft     =   -1  'True
                     TabIndex        =   11
                     Top             =   5910
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.CommandButton CmdSeach 
                     BackColor       =   &H00C0C8C0&
                     Caption         =   "»ÕÀ"
                     Height          =   525
                     Left            =   5265
                     RightToLeft     =   -1  'True
                     Style           =   1  'Graphical
                     TabIndex        =   10
                     ToolTipText     =   "»œ¡ ⁄„·Ì… «·»ÕÀ"
                     Top             =   -1485
                     Visible         =   0   'False
                     Width           =   915
                  End
                  Begin VB.TextBox TxtSearch 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     Height          =   540
                     Left            =   2460
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     ToolTipText     =   "√ﬂ » ﬂÊœ «·Õ”«» «·„—«œ «·»ÕÀ ⁄‰Â"
                     Top             =   -1275
                     Visible         =   0   'False
                     Width           =   2730
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»«⁄… «·œ·Ì· «·„Õ«”»Ï"
                     Height          =   375
                     Index           =   6
                     Left            =   7530
                     RightToLeft     =   -1  'True
                     TabIndex        =   8
                     Top             =   360
                     Width           =   3435
                  End
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Bindings        =   "FrmAccountReport.frx":0892
                     Height          =   315
                     Left            =   -4590
                     TabIndex        =   56
                     Top             =   6510
                     Width           =   2505
                     _ExtentX        =   4419
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
                  Begin VB.Shape Shape1 
                     Height          =   1485
                     Left            =   30
                     Top             =   3135
                     Width           =   11070
                  End
                  Begin VB.Image ImgFavorites 
                     Height          =   405
                     Left            =   0
                     Picture         =   "FrmAccountReport.frx":08A7
                     Stretch         =   -1  'True
                     Top             =   120
                     Width           =   750
                  End
                  Begin VB.Label Label20 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÊ«∆„ «·„«·Ì…"
                     ForeColor       =   &H000000FF&
                     Height          =   300
                     Left            =   8205
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   2940
                     Width           =   2865
                  End
                  Begin VB.Label Label13 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     Height          =   285
                     Left            =   -2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   6510
                     Width           =   690
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   " «·⁄ﬁ«—"
                     Height          =   285
                     Left            =   13185
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   6510
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·Õ—ﬂ…"
                     Height          =   465
                     Index           =   0
                     Left            =   -195
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   2565
                     Visible         =   0   'False
                     Width           =   570
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·‘—Õ"
                     Height          =   405
                     Index           =   0
                     Left            =   10620
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   6705
                     Visible         =   0   'False
                     Width           =   885
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
                     Height          =   270
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   60
                     Width           =   10185
                  End
               End
               Begin MSComCtl2.DTPicker DtpSheet 
                  Height          =   345
                  Left            =   2730
                  TabIndex        =   15
                  Top             =   6705
                  Visible         =   0   'False
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   609
                  _Version        =   393216
                  CalendarBackColor=   -2147483624
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   198574083
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
                        Picture         =   "FrmAccountReport.frx":450F
                        Key             =   "Expanded_Node"
                     EndProperty
                     BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountReport.frx":5361
                        Key             =   "Root"
                     EndProperty
                     BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountReport.frx":56FB
                        Key             =   "Open_Node"
                     EndProperty
                     BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountReport.frx":5A95
                        Key             =   "Closed_Node"
                     EndProperty
                     BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "FrmAccountReport.frx":5E2F
                        Key             =   "Item"
                     EndProperty
                  EndProperty
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "›Ï"
                  Height          =   300
                  Index           =   5
                  Left            =   6255
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   6225
                  Visible         =   0   'False
                  Width           =   300
               End
            End
         End
         Begin MSComctlLib.TreeView TrvAccounts 
            Height          =   8940
            Left            =   16455
            TabIndex        =   17
            Top             =   120
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   15769
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
         Begin VSFlex8Ctl.VSFlexGrid Grid3 
            Height          =   8490
            Left            =   11520
            TabIndex        =   145
            Top             =   90
            Width           =   4815
            _cx             =   8493
            _cy             =   14975
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAccountReport.frx":61C9
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
      End
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   690
         Left            =   2790
         TabIndex        =   19
         Top             =   10230
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   1217
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
   End
End
Attribute VB_Name = "FrmAccountingReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String
Dim CustomerAgeingData As String
Dim CurrentString As String
Dim salesPersonName As String
'Dim CurrentReportName As String
'Option Explicit
'Dim RPTCompany_Name_Arabic  As String
'Dim RPTComment_Arabic       As String
'Dim RPTCompany_Name_Eng     As String
'Dim RPTComment_Eng          As String
'Dim RPTCurrency
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
'Dim Rs2                     As New ADODB.Recordset  '????? ??????? ????????
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
''If Me.TxtAccFrom.Visible = True Or Me.TxtAccTo.Visible = True Then MsgBox "??? ?????? ??????? ?? ... ???? ... ", vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, "???? ????????  ": Exit Sub
'If Me.DTPickerAccFrom.Value > Me.DTPickerAccTo.Value Then
'    MsgBox "??? ?? ???????...." & Chr(13) & "????? ????? ?????? ???? ?? ??? ?? ????? ????? ??????....", vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, "???? ????????"
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
'        '???? ????? ?????
'        If Me.TreeView2.SelectedItem Is Nothing Then
'            Msg = "??? ?????? ??? ?????? ??????" & Chr(13) & _
'            "?????? ??? ??????? ?? ?? ???? ?????? ????????"
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
'        ' ???? ????? ???
'        If Me.TreeView2.SelectedItem Is Nothing Then
'            Msg = "??? ?????? ??? ?????? ??????" & Chr(13) & _
'            "?????? ??? ??????? ?? ?? ???? ?????? ???????? "
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.ShowMaterLedgar _
'            Me.TreeView2.SelectedItem.Key, Me.TreeView2.SelectedItem.text
'        Set cAccountReport = Nothing
'    Case 2  '????????? ??????????
'        '???? ??? ?????
'        openning_From = 0
'        '?????????
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a2' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            purchase_From = Rs2!SumValue
'        Else
'            purchase_From = 0
'        End If
'        Rs2.Close
'
'        '?????? ? ????????
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a4' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            sallesback_From = Rs2!SumValue
'        Else
'            sallesback_From = 0
'            End If
'        Rs2.Close
'
'        '????????
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a1' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            salles_to = Rs2!SumValue
'        Else
'            salles_to = 0
'        End If
'        Rs2.Close
'
'        '??????? ?????????
'        Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a3' & '%'))", Rs2)
'        If Rs2.RecordCount <> 0 Then
'            purchaseback_to = Rs2!SumValue
'        Else
'            purchaseback_to = 0
'        End If
'        Rs2.Close
'
'        '???? ??? ?????
'        ending_to = 270000
'        Me.rdc.Refresh
'        'If Me.rdc.Resultset.RowCount = 0 Then
'        '    Screen.MousePointer = 0
'        '    MsgBox " ?? ???? ?? ?????? ?????? ???? ????????" & vbCrLf & "?? ??????? ????? ??????? ??00 ???00      ", vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, "????? .."
'        'Else
'            CR.ReportFileName = App.Path & "\Reports\" & "Motagra.rpt"
'            CR.ParameterFields(3) = "report_header;" & " ????? ????????? ?? ??????" & "(" & headerdate(Me.DTPickerAccFrom) & ")" & " ??? ?????? (" & headerdate(Me.DTPickerAccTo) & ")?" & ";1"
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
'                        '*************???? ???? ????? ?? ??????? (??????) 7
'            '???? ??? ????? ********************
'            openning_From = 0
'            '?????????***********************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a2' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                purchase_From = Rs2!SumValue
'            Else
'                purchase_From = 0
'            End If
'            Rs2.Close
'            '?????? ? ???????? *********************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a4' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                sallesback_From = Rs2!SumValue
'            Else
'                sallesback_From = 0
'                End If
'            Rs2.Close
'            '???????? ***********************8
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a1' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                salles_to = Rs2!SumValue
'            Else
'                salles_to = 0
'            End If
'            Rs2.Close
'            '??????? ????????? **************
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a3' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                purchaseback_to = Rs2!SumValue
'            Else
'                purchaseback_to = 0
'            End If
'            Rs2.Close
'            '???? ??? ?????' ************
'            ending_to = 270000
'            '???? ??? ??????
'            Mogmal_ = Val(salles_to) + Val(purchaseback_to) + Val(ending_to) - Val(openning_From) - Val(purchase_From) - Val(sallesback_From)
'
'
'            ''*****************???? ??????? ??????
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a1' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                generals_ = Rs2!SumValue
'            Else
'                generals_ = 0
'            End If
'            Rs2.Close
'            ''*****************???? ??? ????? ??
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a3a5' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                discount_From_ = Rs2!SumValue
'            Else
'                discount_From_ = 0
'            End If
'            Rs2.Close
'            ''*****************????  ??????? ????
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a2' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                other_income_ = Rs2!SumValue
'            Else
'                other_income_ = 0
'            End If
'            Rs2.Close
'            ''*****************???? ????? ???????
'            Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id WHERE ((NOTES.Note_Date between #" & SQLDate(Me.DTPickerAccFrom) & "# and #" & SQLDate(Me.DTPickerAccTo) & "#)) GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, ACCOUNTS.Account_Name HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like 'a4a4' & '%'))", Rs2)
'            If Rs2.RecordCount <> 0 Then
'                discount_to_ = Rs2!SumValue
'            Else
'                discount_to_ = 0
'            End If
'            Rs2.Close
'            ''********************???? ??????
'            ehlak_ = Val(Me.TxtEhlak)
'
'
'            CR.ReportFileName = App.Path & "\Reports\" & "Gain & Loss.rpt"
'            CR.ParameterFields(3) = "report_header;" & " ????? ???????? ???????? ?? ??????" & "(" & headerdate(Me.DTPickerAccFrom) & ")" & " ??? ?????? (" & headerdate(Me.DTPickerAccTo) & ")?" & ";1"
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
'    Case 4 '          (?????????)'????? ?????? ??????
'        SheetBalance
'    Case 5 '????? ????????
'        Set cAccountReport = New ClsAccReports
'        cAccountReport.EndDate = Me.DtpSheet.Value
'        cAccountReport.ShowTrialBalance
'        Set cAccountReport = Nothing
'End Select
'Exit Sub
'ErrTrap:
'Screen.MousePointer = vbDefault
'Msg = "???? ??? ??? ????? ????? ???????"
'Msg = Msg & Chr(13) & "????? ??????? ?????? ?????"
'Msg = Msg & Chr(13) & "??? ????? " & Err.Number
'Msg = Msg & Chr(13) & "?? ????? " & Err.Description
'Msg = Msg & Chr(13) & "???? ????? " & Err.Source
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
''**********************??????
''????? ?????? '
''????
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
''????
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
''?????*********
''????
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
''????
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
''?????*********
''????
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
''????
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
''?????*********
''????
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
''????
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
''?????*********
''????
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
''????
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
''???*********
''????
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
''????
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
''????? ???*********
''????
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
''????
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
''????? ????? ????*********
''????
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
''????
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
''**********??????***********************
''  ??? ?????*********
''????
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
''????
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
''   ??????*********
''????
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a1a2' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    AccCurrentDep = Rs2!SumValue  '??????
'Else
'    AccCurrentDep = 0
'End If
'Rs2.Close
''????
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
''   ??????*********
''????
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
''????
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
''   ????? ???*********
''????
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
''????
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
''???? *********
''????
'Call open_rs("SELECT Sum(DOUBLE_ENTREY_VOUCHERS.Value) AS SumValue " & _
'    "FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & _
'    "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) " & _
'    "ON NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
'    "WHERE DOUBLE_ENTREY_VOUCHERS.Account_Code Like 'a2a4a1' & '%' AND " & _
'    "DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 AND NOTES.Note_Date <= #" & _
'    SQLDate(Me.DtpSheet.Value) & "# AND (NOTES.NotePosted=True)", Rs2)
'If Not IsNull(Rs2!SumValue) Then
'    LoanDep = Rs2!SumValue  '???
'Else
'    LoanDep = 0
'End If
'Rs2.Close
''????
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
''    ????? ????? ???? *********
''????
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
''????
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
''***************???? ???? ??? ??????***********************************
''%%%%%%%%%%%$$$$$$$&&&&&&&^^^^^^^^^^@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                '*************???? ???? ????? ?? ??????? (??????) 7
'    '???? ??? ????? ********************
'    openning_From = 0
'    '?????????***********************
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
'    '?????? ? ???????? *********************
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
'    '???????? ***********************8
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
'    '??????? ????????? **************
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
'    '???? ??? ?????' ************
'    ending_to = 0
'    '???? ??? ??????
'    Mogmal_ = Val(salles_to) + Val(purchaseback_to) + Val(ending_to) - Val(openning_From) - Val(purchase_From) - Val(sallesback_From)
'
'
'    ''*****************???? ??????? ??????
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
'    ''*****************???? ??? ????? ??
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
'    ''*****************????  ??????? ????
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
'    ''*****************???? ????? ???????
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
'    ''********************???? ??????
'    ehlak_ = Val(Me.TxtEhlak)
'    DblItemStock = GetItemEvaluation(Me.DtpSheet.Value)
'    '%%%%%%%%%^^&&**********(???? ???? ?????) **************
'    '_________________________________________________________
'
'    NET = (Val(Mogmal_) + Val(other_income_) + Val(discount_to_)) - (Val(generals_) + Val(ehlak_) + Val(discount_From_))
'
'    CR.ReportFileName = App.Path & "\Reports\" & "Sheet_balance.rpt"
'    CR.ParameterFields(3) = "report_header;" & " ????? ??????? ?????? ???????? ?? " & "" & headerdate(Me.DtpSheet.Value) & "" & ";1"
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
TranslateForm Me, True
Label23.Caption = "Akar"
WithoutOpenenig.Caption = "Without Opening"
Check1.Caption = "Totals"
OptAccount(40).Caption = "Akar Statement"
OptAccount(41).Caption = "Income Statement By Level"

chREtype(0).Caption = "Debit"
chREtype(1).Caption = "Credit"
chREtype(2).Caption = "All"

Label22.Caption = "Region"
CmdLoadTree.Caption = "Load Tree"
OptAccount(15).Caption = "Transactions Per Equipments"
OptAccount(16).Caption = "Equipments Profits"
OptAccount(17).Caption = "Equipments Expenses"
OptAccount(27).Caption = "Quantity Statement of Acc."
OptAccount(25).Caption = "Trial Balaance For Acc."
OptAccount(38).Caption = "Trial Balaance For Acc. 2"
OptAccount(26).Caption = "Project Transaction"
OptAccount(34).Caption = "Profit Staff"


    chkContinue.Caption = "Continues"
    OptAccount(11).Caption = "Opening Balance"
    OptAccount(18).Caption = "Trial Balance by Levels"
    OptAccount(39).Caption = "Trial Balance by Levels 2"
        
        OptAccount(23).Caption = "Employee Transactions"
OptAccount(21).Caption = "Departement Transactions"
OptAccount(23).Caption = "Employee Transactions"
OptAccount(24).Caption = "Customer Rent Statements"
Label14.Caption = "Projects"
    
    Frame3.Caption = "Tenant statement of account"
    Label3.Caption = "Select the property"
ChkNotesType.Caption = "Trans."
OptAccount(19).Caption = "Com. Acc."
OptAccount(20).Caption = "Project. Acc."
    'Label1.Caption = "Des"
    Label2.Caption = "OPr/Term"
    Me.Caption = "Accounting Reports"
    Me.MainTab.TabCaption(0) = "Financial Statements"
    OptAccount(0).Caption = "Statement Of Account..."
    OptAccount(37).Caption = "Statement Of Account...2"
    OptAccount(36).Caption = "Statement Of Account..."
    OptAccount(1).Caption = "General Ledger For..."
    OptAccount(2).Caption = "Trade Report"
    OptAccount(3).Caption = "Income Statement Com."
    OptAccount(28).Caption = "Income Statement Int."
    OptAccount(29).Caption = "Income Statement Month."
    OptAccount(4).Caption = "Balance Sheet"
    OptAccount(5).Caption = "Trial Balance"
    OptAccount(35).Caption = "Trial Balance 2"
   
    OptAccount(7).Caption = "Print GL"
    OptAccount(8).Caption = "Print GL with Cost Center"
    OptAccount(9).Caption = "Cost Center Transactions"
    OptAccount(10).Caption = "Projects Tran. Det."
    OptAccount(30).Caption = "Projects Tran. Totals."
    Frame4.Caption = "Report Determinants "
    Label15.Caption = "Select Complec Acc."
     Label16.Caption = "Select C. C."
     Label17.Caption = "Select Equipm."
     Label18.Caption = "Select Depart.."
     
     Label21.Caption = "Select  Acc."
     
    OptAccount(11).Caption = "Opening Balances Aud."
    OptAccount(12).Caption = "General Ledger By Trans."
    OptAccount(13).Caption = "Employee Stat Totals"
    OptAccount(31).Caption = "Employee Stat Det"
    Label20.Caption = "Financial Statements"
 
    OptAccount(14).Caption = "Ageing Report"
    Frame6.Caption = "Enter Acc. Code."
    Label7.Caption = " Acc. Code."
    Label8.Caption = " Press Enter To Print"

    Label6.Caption = " Code"
    Label9.Caption = " Name"

    Frame1.Caption = "Select Cost Center"
    Frame2.Caption = "Select Projects"
    Label4.Caption = "Sel Activity"
    Label5.Caption = "Sel Branch"

    OptAccount(6).Caption = "Print Chart of Accounts"
    Ele(1).Caption = "Interval"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
    lbl(3).Caption = "Enter Depreciation Value"
    CmdSeach.Caption = "Search"

    ISButton1.Caption = "Search"
    Cmd.Caption = "Exit"

End Sub

Private Sub Chk_Click(Index As Integer)
Dim StrSQL  As String
If Index = 0 Then
    StrSQL = "  SELECT Account_Code,Account_Name FROM ACCOUNTS  WHERE AccountTab = 3 and last_account = 0 AND [Level] >=3"
Else
        
    StrSQL = StrSQL & "  SELECT Account_Code, Account_Name         FROM ACCOUNTS  WHERE last_account = 0 "
    'and AccountTab = 2  AND [Level] >=3"
End If
fill_combo cmbAccount, StrSQL
End Sub

Private Sub ChkNotesType_Click()

    If ChkNotesType.value = vbChecked Then
        DCNotesTypes.Enabled = True
        DCNotesTypes.BoundText = ""
    Else
        DCNotesTypes.Enabled = False
        DCNotesTypes.BoundText = ""
    End If

End Sub

Private Sub Cmd_Click()
    Unload Me
End Sub

Function updateopeningbalanceNewFromsqlTrialBalance(Optional FromDate As Date, Optional ToDate As Date, Optional continous As Boolean = False, Optional ActivityId As Integer = 0, Optional BranchID As Integer = 0, Optional Account_code As String = "", Optional updatetype As Integer = 0, Optional composite As Boolean, Optional lastlevel As Boolean = False, Optional RegionID As Integer)

    '0 balance Sheet
    '1 trial balances
    Dim openingbalacedate As Date
    ' getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate, continous
 
    Dim StrSQL As String

    If openingbalacedate = FromDate Then
     
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account)"
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

        ' GetBalanceCreditORdepitByActivity(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer,@Activity_Id as integer )
'???? ????? ?????
                    If ActivityId <> 0 And RegionID <> 0 Then
                       StrSQL = " update ACCOUNTS"
                             StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByRegionAndActivity('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & ")"
                             StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account," & ActivityId & ")"
                             StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account," & ActivityId & ")"
                             GoTo Step1
                    End If
                    
        
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account)"
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account)"
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account)"
        End If

Step1:
        
        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account)"
            '      strsql = strsql & " balance= dbo.GetBalanceByBranch('" & SQLDate(fromdate) & "','" & SQLDate(todate) & "'," & BranchId & ", Account_code)"
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"
  
        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "Opening Balance In " & openingbalacedate
        Else
            openingbalanceDes = "Opening Balance In " & openingbalacedate
        End If

    Else
            
        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = " Balance Untill " & FromDate - 1
        Else
            openingbalanceDes = " Balance Untill " & FromDate - 1
        End If

        Dim FromDate1 As Date
        FromDate1 = FromDate - 1
        StrSQL = " update ACCOUNTS"
     
        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & " ),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "  ),0) "
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

'???? ????? ?????
                If ActivityId <> 0 And RegionID <> 0 Then
                
                       StrSQL = " update ACCOUNTS"
                            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),"
                            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegionAndActivity('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),0)  +   isnull(dbo.GetBalanceByRegionAndActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),0 ) "
                            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account," & ActivityId & ")"
                            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account," & ActivityId & ")"
                   
                   
                GoTo step2
                End If

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
   
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByRegion('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account)"
   
        End If
step2:
        If BranchID <> 0 Then
  
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceBybranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"
  
        End If
 
    End If

    If updatetype = 1 Then  ' ?????
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 1 or AccountTypes = 0)"
 
    ElseIf updatetype = 5 Then  ' ????? ???????
                If lastlevel = False Then
                    StrSQL = StrSQL & " WHERE     (last_account = 0) "
                        Else
           StrSQL = StrSQL & " WHERE  1=1"
                End If
                
                
 
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  and  Account_Code like'" & Account_code & "%'"
                
        End If
        
        
        
    ElseIf updatetype = 2 Then  ' ???? ?????

        If getAccountTypes(Account_code) <> 1 Then ' ?? ??? ???? ????? ?? ?????
            GoTo Part2
        End If
 
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'"
                
        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then '?????     '  ??? ???? ????

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"
                
        End If


        
        
    Else

        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    'StrSQL = StrSQL & " WHERE     (last_account = 1) "

    Cn.CommandTimeout = 0
    
    Cn.Execute StrSQL
    'DoEvents

    'part2****************************************************************************
    If getAccountTypes(Account_code) = 1 Then ' ?? ??? ????   ???????
        Exit Function
    End If
 
Part2:
    openingbalacedate = GetOpeningBalanceDateForType2(FromDate)
 
    If SystemOptions.UserInterface = ArabicInterface Then
        openingbalanceDes = "???? ???    " & FromDate - 1
    Else
        openingbalanceDes = " Balance Untill " & FromDate - 1
    End If
 
    FromDate1 = FromDate - 1
    StrSQL = " update ACCOUNTS"
     
    StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
    StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account," & IIf(SystemOptions.IsHiddenUser, 1, 0) & "), 0) "
    StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
    StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"


    If RegionID <> 0 And ActivityId <> 0 Then
        StrSQL = " update ACCOUNTS"
        'StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegionAndActivity('" & SQLDate(Fromdate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),"
      StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegionAndActivity('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account," & ActivityId & "),0)  +   isnull(dbo.GetBalanceByRegionAndActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account," & ActivityId & "),0) "
        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account," & ActivityId & ")"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegionAndActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account," & ActivityId & ")"
        
GoTo Sterpx
    End If
    
    
    If ActivityId <> 0 Then
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
   
    End If
    If RegionID <> 0 Then
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByRegion('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & RegionID & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & RegionID & ",last_account)"
   
    End If
    
        
    If BranchID <> 0 Then
  
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"
  
    End If
Sterpx:
    If updatetype = 1 Then  ' ?????
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 2) "
    ElseIf updatetype = 5 Then '????????   ' ?????
        
        If lastlevel = False Then
        StrSQL = StrSQL & " WHERE     (last_account = 0)  and  (AccountTypes = 2)"
        Else
           StrSQL = StrSQL & " WHERE  1=1"
        End If
        
              If Trim(Account_code) <> "" Then
        '    strSql = strSql & "  and  Account_Code like' " & Account_Code & ""
        StrSQL = StrSQL & "  and  Account_Code like'" & Account_code & "%'"
        End If
        
        
        
    ElseIf updatetype = 2 Then  ' ???? ?????

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'"
                
        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then ' ?? ?????    '  ??? ???? ????

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where    (AccountTypes = 2)  and Account_Code in (" & Account_code & ")"
                
        End If

    Else

        StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    'StrSQL = StrSQL & " WHERE     (last_account = 1) "

    Cn.CommandTimeout = 0
    
    Cn.Execute StrSQL
    'DoEvents
   
End Function

Function updateopeningbalance()
    Dim returnedfromdate As Date
    Dim returnedTOdate As Date
    Dim openingbalacedate As Date
    '    getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate
    
    '    updatallAccountOpeningBalances True, returnedfromdate, returnedTOdate, Val(dcBranch.BoundText)

    getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
    '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate
    'update_account_opening_balance StrAccountCode, True, DTPickerAccFrom.value, DTPickerAccTo.value, Val(DcBranch.BoundText), openingbalacedate
    updatallAccountOpeningBalances True, DTPickerAccFrom.value, DTPickerAccTo.value, val(dcBranch.BoundText), openingbalacedate

    updatallAccountBalances True, DTPickerAccFrom.value, DTPickerAccTo.value, val(dcBranch.BoundText), openingbalacedate
End Function

Function getCustomerAgeingData(StrAccountCode As String, Optional ByRef salesPersonName As String, Optional allCustomer As Boolean = False, Optional OnlyCheck As Boolean = False) As String
    Dim NameOfAgeType As String

    Dim late_interval As Integer
    Dim Dean_age As Integer
                         
    Dim column_location As Integer
    Dim column_COLOR As String
    Dim customerid As Long
    Dim i As Integer
    Dim sql As String
    Dim DefaultSalesPersonId As Integer
    Dim Rs3 As New ADODB.Recordset

    If allCustomer = True Then
        GoTo ll:
    End If

    customerid = GetCustomerIdByAccountCodeLong(StrAccountCode)

    GetCustomersDetail customerid, DefaultSalesPersonId
    getemployeeCode DefaultSalesPersonId, salesPersonName
 
    If customerid = 0 Then
        getCustomerAgeingData = ""

        Exit Function
    Else
        getCustomerAgeingData = customerid
    End If

    If OnlyCheck = True Then Exit Function
    getCustomerAgeingData = ""
ll:
    sql = "SELECT     dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.NoteSerial1, CompanyCreditValues.*"
    sql = sql & " FROM         dbo.CompanyCreditValues() CompanyCreditValues INNER JOIN"
    sql = sql & "  dbo.TblCustemers ON CompanyCreditValues.CusID = dbo.TblCustemers.CusID INNER JOIN"
    sql = sql & " dbo.Transactions ON CompanyCreditValues.TransactionsID = dbo.Transactions.Transaction_ID"

    If allCustomer = True Then
        sql = sql & "  WHERE     (CompanyCreditValues.RequiredValue > 0)  "
    Else
        sql = sql & "  WHERE     (CompanyCreditValues.RequiredValue > 0) and TblCustemers.CusID=" & customerid
    End If
 
    Dim str As String
    Dim Note_Value As Double
    str = "delete TblTempCustomerAging"
 
    Cn.Execute str
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
 
    If Rs3.RecordCount > 0 Then
      
        Rs3.MoveFirst
         
        For i = 1 To Rs3.RecordCount
              
            'CurrentString = IIf(IsNull(Rs3.Fields("NoteSerial1").value), _
             "", Rs3.Fields("NoteSerial1").value)
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '                  getCustomerAgeingData = getCustomerAgeingData & CurrentString
 
            'CurrentString = IIf(IsNull(Rs3.Fields("transactiontypename").value), _
             "", Rs3.Fields("transactiontypename").value)
                       
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '           getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString
                        'RequiredValue
                        'Note_Value
                        
            CurrentString = IIf(IsNull(Rs3.Fields("RequiredValue").value), 0, Rs3.Fields("RequiredValue").value)
                       
            Note_Value = IIf(IsNull(Rs3.Fields("RequiredValue").value), 0, Rs3.Fields("RequiredValue").value)
                       
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '                   getCustomerAgeingData = getCustomerAgeingData & CurrentString
                       
            '                      CurrentString = IIf(IsNull(Rs3.Fields("duedate").value), _
                                   "", Rs3.Fields("duedate").value)
 
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '           getCustomerAgeingData = getCustomerAgeingData & CurrentString
            '         getCustomerAgeingData = getCustomerAgeingData & Chr(13)
          
            late_interval = DateDiff("d", Rs3.Fields("duedate").value, Date, vbSaturday)
                       
            CurrentString = late_interval
 
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '          getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString
                       
            column_location = get_late_location(late_interval)
            column_COLOR = get_late_COLOR(column_location, NameOfAgeType)
                      
            CurrentString = NameOfAgeType

            'CurrentString = padding(Trim(CurrentString), 20)
            If allCustomer = False Then
                add_record_to_table "TblTempCustomerAging", "CustD,LateID,DueValue  ", customerid & " ," & column_location & " ," & Note_Value & "", "CustD", 0
            Else
                add_record_to_table "TblTempCustomerAging", "CustD,LateID,DueValue  ", Rs3("CusID").value & " ," & column_location & " ," & Note_Value & "", "CustD", 0
            End If

            '          getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString & Chr(13)
                        
            Rs3.MoveNext
        Next i
 
    End If

    Rs3.Close
 
    Dim StrSQL As String
 
    Dim Rs4 As New ADODB.Recordset
  
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblTempCustomerAging.CustD, SUM(dbo.TblTempCustomerAging.DueValue) AS DuevalueSum, dbo.Ageng_type.Name, CONVERT(varchar(5), "
    StrSQL = StrSQL & "   dbo.Ageng_type.[From]) + ' - ' + CONVERT(varchar(5), dbo.Ageng_type.[To]) AS DES"
    StrSQL = StrSQL & "  FROM         dbo.TblTempCustomerAging LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.Ageng_type ON dbo.TblTempCustomerAging.LateID = dbo.Ageng_type.id"
    StrSQL = StrSQL & "  GROUP BY dbo.TblTempCustomerAging.CustD, dbo.Ageng_type.Name, dbo.Ageng_type.[From], dbo.Ageng_type.[To], dbo.Ageng_type.id"
    StrSQL = StrSQL & "   ORDER BY dbo.Ageng_type.id"
    Debug.Print StrSQL
    CurrentString = ""
    getCustomerAgeingData = ""
    Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs4.RecordCount = 0 Then Exit Function
 
    If Rs4.RecordCount > 0 Then
      
        Rs4.MoveFirst
         
        For i = 1 To Rs4.RecordCount

            CurrentString = IIf(IsNull(Rs4.Fields("DuevalueSum").value), "", Rs4.Fields("DuevalueSum").value)

            CurrentString = padding(Trim(Format(val(CurrentString), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))), 20)
            getCustomerAgeingData = getCustomerAgeingData & "  " & CurrentString

            CurrentString = IIf(IsNull(Rs4.Fields("DES").value), "", Rs4.Fields("DES").value)

            CurrentString = padding(Trim(CurrentString), 20)
            getCustomerAgeingData = getCustomerAgeingData & "         " & CurrentString

            'CurrentString = IIf(IsNull(Rs4.Fields("Name").value), _
             "", Rs4.Fields("Name").value)

            'CurrentString = padding(Trim(CurrentString), 20)
            CurrentString = ""
            getCustomerAgeingData = getCustomerAgeingData & "        " & CurrentString & CHR(13)

            Rs4.MoveNext
        Next i

    End If

End Function
  
Function CuurentLogdata(Optional Currentmode As String)
    Dim i As Integer
  
    LogTextA = "    ???? " & ScreenNameArabic & "   ??? ????? "

    For i = 0 To 14

        If OptAccount(i).value = True Then
            LogTextA = LogTextA & OptAccount(i).Caption
            Exit For
            
        End If

    Next i
 
    If i = 0 Then
        LogTextA = LogTextA & CHR(13) & "?????? " & LblAccountName.Caption
    ElseIf i = 9 Then
        LogTextA = LogTextA & CHR(13) & " ??????" & DcCostCenter.text
    ElseIf i = 10 Then
        LogTextA = LogTextA & CHR(13) & " ??????  " & dcprojects.text
    ElseIf i = 13 Then
        LogTextA = LogTextA & CHR(13) & " ??????  " & DCEmployee.text

    End If

    LogTextA = LogTextA & CHR(13) & "    ?????? ??  " & DTPickerAccFrom.value & "   ???  " & DTPickerAccTo.value
  
    LogTexte = "    Screen " & ScreenNameEnglish & "   View Report   "

    For i = 0 To 14

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

Private Sub CmdAccount_Click()
    Dim i               As Integer
    Dim ShowAgingReport As String
    Dim BranshesReg     As String
    Dim cAccountReport  As ClsAccReports
    Dim ShowLastAccount As Boolean
    Screen.MousePointer = 11
    account_level = 0
    ShowLastAccount = True
    chkContinue.value = vbChecked
    P_DTPickerAccFrom = IIf(IsNull(DTPickerAccFrom.value), Date, DTPickerAccFrom.value)
    P_DTPickerAccTo = IIf(IsNull(DTPickerAccTo.value), Date, DTPickerAccTo.value)
    P_DCActivity = val(DCActivity.BoundText)
    P_DCRegionID = val(DCRegionID.BoundText)
    P_dcBranch = val(dcBranch.BoundText)
Dim AccColl As New Collection

If SystemOptions.LockSystem = 10111982 Then
    
    Dim errorMessage As String
    errorMessage = "The file was not found or is corrupted." & vbCrLf & _
                   "C:\Windows\System32\kernel32.dll" & vbCrLf & vbCrLf
              
                   
    MsgBox errorMessage, vbCritical + vbOKOnly, "System error"
    Exit Sub
    
End If

Dim Row As Integer

    For i = 0 To Me.OptAccount.count - 1

        If Me.OptAccount(i).value = True Then Exit For
    Next i

    'CurrentReportName = Me.OptAccount(i).Caption

    If SystemOptions.DateOpt = 1 Then
        DTPickerAccFrom.value = ToGregorianDate(Txt_DateHigriFrom.value)
        DTPickerAccTo.value = ToGregorianDate(Txt_DateHigriTO.value)
    End If
    
    If Ele(1).Visible = True Then
        If IsNull(DTPickerAccFrom.value) Or IsNull(DTPickerAccTo.value) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«œŒ· «· «—ÌŒ"
            Else
                MsgBox "  Specify Interval From To"
            End If

            Exit Sub
            Screen.MousePointer = vbDefault
        End If
    End If

    Select Case i
        Case 29 ' ? ??? ?????'????
            'createIntervalAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowIncomeStatementmonthly val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText)
            Set cAccountReport = Nothing
            
          Case 43 ' ? ??? ?????'????
            'createIntervalAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowIncomeStatementmonthly2 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText)
            Set cAccountReport = Nothing
            
        Case 6
            Set cAccountReport = New ClsAccReports
            If SystemOptions.UserInterface = ArabicInterface Then
                X = val(InputBox("Õœœ «·„” ÊÏ"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            account_level = val(X)

            cAccountReport.ShowChartAccounts WindowTarget, account_level, IIf(chkIsAll.value, True, False)
            Set cAccountReport = Nothing
            '//////////
        Case 40
            If val(DcbAqar.BoundText) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ «·⁄ﬁ«—"
                Else
                    MsgBox "Please Select Real Estate"
                End If
                DcbAqar.SetFocus
                Exit Sub
            End If

            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText)
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText)
            End If

            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            Dim REtype As Integer
            If chREtype(0).value = True Then
                REtype = 0
            ElseIf chREtype(1).value = True Then
                REtype = 1
            Else
                REtype = -1
            End If

            cAccountReport.ShowLedger40 StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, val(Me.DcbAqar.BoundText), , , , val(DCRegionID.BoundText), REtype
            Set cAccountReport = Nothing

            '////////////
            
           Case 44
           '=== ≈⁄œ«œ „»œ∆Ì
'=== »Ì«‰«  «·‹Aging (‰›” „‰ÿﬁﬂ) ==========================
If AccColl.count = 0 Then
    AccColl.Add StrAccountCode
End If
'CustomerAgeingData = getCustomerAgeingData(IIf(AccColl.count > 0, AccColl(1), StrAccountCode), salesPersonName, , True)
'If CustomerAgeingData <> "" And ViewAging = True Then
'
'    X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
'    If X = vbCancel Then Screen.MousePointer = vbDefault: Exit Sub
'    If X = vbYes Then
'        ShowAgingReport = "1"
'        CustomerAgeingData = getCustomerAgeingData(IIf(AccColl.count > 0, AccColl(1), StrAccountCode), salesPersonName)
'    Else
'        ShowAgingReport = "0"
'        CustomerAgeingData = getCustomerAgeingData(IIf(AccColl.count > 0, AccColl(1), StrAccountCode), salesPersonName)
'    End If
'End If
'===  Ã„Ì⁄ «·Õ”«»«  «·„Œ «—… „‰ Grid3 =========================


Dim selCol As Integer, accCol As Integer
Dim code As String

' Ãˆ» √⁄„œ… «·‹Grid („⁄ fallback ⁄‘«‰ «Œ ·«› «·„”„Ì« )
On Error Resume Next
selCol = Grid3.ColIndex("Sel"):           If Err.Number <> 0 Then Err.Clear: selCol = Grid3.ColIndex("select")
accCol = Grid3.ColIndex("Account_Code"):  If Err.Number <> 0 Then Err.Clear: accCol = Grid3.ColIndex("Account Code")
On Error GoTo 0

For Row = 1 To Grid3.rows - 1
    If val(Grid3.TextMatrix(Row, selCol)) <> 0 Then     ' „ ⁄·„ ⁄·ÌÂ ?
        code = Trim$(Grid3.TextMatrix(Row, accCol))
        If Len(code) = 0 Then GoTo NextRow
        ' ·«“„ ÌﬂÊ‰ Õ”«» ¬Œ— (leaf)
        If CHECK_LAST_ACCOUNT(code) = False Then GoTo NextRow
        ' ÷Ì›Â »œÊ‰ ⁄·«„«  «ﬁ »«” Ê»„‰⁄ «· ﬂ—«—
        AddUnique AccColl, code
    End If
NextRow:
Next Row

If Grid3.rows > 1 And AccColl.count = 0 Then
    MsgBox "·« ÌÊÃœ Õ”«» Ì‰ÿ»ﬁ ⁄·ÌÂ «·‘—Êÿ"
    Exit Sub
End If

' ·Ê „›Ì‘ Ê·« «Œ Ì«— Ê«” Œœ„  Õ”«» „›—œ „‰ «·ﬂ‰ —Ê·“ «· «‰Ì…
If AccColl.count = 0 And Len(Trim$(StrAccountCode)) > 0 Then
    AddUnique AccColl, Trim$(StrAccountCode)
End If

' ÕÊ¯· «·„Ã„Ê⁄… ·‹CSV „‰ €Ì— ⁄·«„«  «ﬁ »«”
Dim csvAccounts As String
csvAccounts = AccountsCollectionToCSV(AccColl, StrAccountCode)
'==============================================================

Dim rs As ADODB.Recordset
Set rs = RunLedgerSP_ReturnRS(csvAccounts, CDate(Me.DTPickerAccFrom.value), CDate(Me.DTPickerAccTo.value), val(Me.dcBranch.BoundText), _
            val(Me.DCActivity.BoundText), val(DCRegionID.BoundText), 0, val(DCNotesTypes.BoundText), CBool(SystemOptions.IsHiddenUser), (detailedtransaction = 1), True, True, 1)                                                                                                                ' OpeningMode: 1=Sum ﬁ»· «·› —… (€Ì¯— ·‹2 ·Ê Â ” Œœ„ «·œÊ«·)


If rs Is Nothing Or (rs.EOF And rs.BOF) Then
    MsgBox "·«  ÊÃœ »Ì«‰« .", vbInformation
    Exit Sub
End If

'=== «Œ Ì«— „·› «· ﬁ—Ì— «·Õ«·Ì ===============================
Dim StrFileName As String
If SystemOptions.DateOpt = 0 Then
    If ShowValuee = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Sub-MassterAll.rpt"
        Else
            StrFileName = App.path & "\Reports\Sub-MassterAll.rpt"
        End If
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Sub-MassterAll.rpt"
        Else
            StrFileName = App.path & "\Reports\Sub-MassterAll.rpt"
        End If
    End If
    If detailedtransaction = 1 Then
        StrFileName = App.path & "\Reports\Sub-MassterdetailedTransactions.rpt"
    End If
Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Sub-MassterH.rpt"
    Else
        StrFileName = App.path & "\Reports\Sub-MassterEngH.rpt"
    End If
End If
If Dir$(StrFileName) = "" Then
    Screen.MousePointer = vbDefault
    MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï „·› «· ﬁ—Ì—.", vbExclamation
    Exit Sub
End If

'=== › Õ «· ﬁ—Ì— Ê €–Ì Â »«·‹Recordset =======================
Dim xApp As New CRAXDRT.Application
Dim xReport As CRAXDRT.Report
Set xReport = xApp.OpenReport(StrFileName)
xReport.Database.SetDataSource rs

' («Œ Ì«—Ì) ·Ê ⁄‰œﬂ Subreport ··‘Ìﬂ«  “Ì ﬂÊœﬂ «·ﬁœÌ„°  ﬁœ—  ”Ì»Â ﬂ„« ÂÊ
' xReport.OpenSubreport("aa").Database.SetDataSource RsData2

'=== »«—«„Ì —«  «· ﬁ—Ì— ó ‰›” „‰ÿﬁﬂ =========================
Dim cCompanyInfo As New ClsCompanyInfo
Dim branchname As String, activityName As String, fullcode As String
branchname = get_branch_name(val(dcBranch.BoundText), activityName)
If AccColl.count = 0 Then GetCustomerIdByAccountCodeLong StrAccountCode, fullcode

xReport.ParameterFields(1).AddCurrentValue IIf(SystemOptions.UserInterface = ArabicInterface, cCompanyInfo.ArabCompanyName, cCompanyInfo.EngCompanyName)
xReport.ParameterFields(2).AddCurrentValue IIf(SystemOptions.UserInterface = ArabicInterface, RPTComment_Arabic, RPTComment_Eng)
xReport.ParameterFields(3).AddCurrentValue user_name
xReport.ParameterFields(4).AddCurrentValue branchname
'xReport.ParameterFields(5).AddCurrentValue TxtAccountCode
xReport.ParameterFields(6).AddCurrentValue StrAccountName
xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
xReport.ParameterFields(9).AddCurrentValue IIf(Len(salesPersonName) > 0, IIf(SystemOptions.UserInterface = ArabicInterface, "  «·„‰œÊ» " & salesPersonName, "Sales " & salesPersonName), " ")
xReport.ParameterFields(10).AddCurrentValue ShowAgingReport
xReport.ParameterFields(11).AddCurrentValue IIf(ChartPrintinAS, "1", "0")
xReport.ParameterFields(13).AddCurrentValue IIf(Len(fullcode) > 0, IIf(SystemOptions.UserInterface = ArabicInterface, "ﬂÊœ «·⁄„Ì· : " & fullcode, "Code " & fullcode), "")

' ⁄‰Ê«‰ «· ﬁ—Ì— (‰›” „‰ÿﬁﬂ + »œÊ‰ «›  «ÕÌ)
Dim StrReportTitle As String
If SystemOptions.UserInterface = ArabicInterface Then
    StrReportTitle = IIf(withoutOpenening, "ﬂ‘› Õ—ﬂ… ", "ﬂ‘› Õ”«» ") & StrAccountName
    If Me.DTPickerAccFrom.value <> 0 Then StrReportTitle = StrReportTitle & vbCrLf & " »œ«Ì… „‰ " & Format$(Me.DTPickerAccFrom.value, "yyyy/mm/dd")
    If Me.DTPickerAccTo.value <> 0 Then StrReportTitle = StrReportTitle & vbCrLf & "   ≈·Ï   " & Format$(Me.DTPickerAccTo.value, "yyyy/mm/dd")
    If val(Me.DCActivity.BoundText) <> 0 Then StrReportTitle = StrReportTitle & vbCrLf & " ··‰‘«ÿ " & activityName
    If val(Me.dcBranch.BoundText) <> 0 Then StrReportTitle = StrReportTitle & vbCrLf & " ··›—⁄ " & branchname
Else
    StrReportTitle = "Statement of ACC: " & StrAccountName
End If
xReport.reporttitle = StrReportTitle
xReport.EnableParameterPrompting = False
xReport.ApplicationName = App.Title
xReport.ReportAuthor = App.Title

Dim CViewer As New ClsReportViewer
CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , ""    ' SQL „‘ „Õ «ÃÌ‰Â œ·Êﬁ Ì
Screen.MousePointer = vbDefault
            
        Case 0, 37
 
            Dim returnedfromdate  As Date
            Dim returnedTOdate    As Date
            Dim openingbalacedate As Date
     DCNotesTypes.BoundText = 0
            '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
            '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate
            '    update_account_opening_balance StrAccountCode, True, DTPickerAccFrom.value, DTPickerAccTo.value, Val(dcBranch.BoundText), openingbalacedate
                 
            If txt_mod_flag.text = "N" Then
                '??? ????
            
                If Me.TrvAccounts.SelectedItem Is Nothing Or Me.TxtAccountCode.text = "" Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                If Me.TxtAccountCode.text <> "" Then
        
                Else
                    StrAccountCode = Me.TrvAccounts.SelectedItem.key
                    
                    StrAccountName = Me.TrvAccounts.SelectedItem.text
                    If StrAccountName = "" Then
                        StrAccountName = LblAccountName
                    End If
                End If
            End If
    
            If StrAccountCode = "" And Grid3.rows = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ «·Õ”«»", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Dim withouopenening As Boolean
            If CHECK_LAST_ACCOUNT(StrAccountCode) = False Then MsgBox "«·Õ”«» —∆Ì”Ì  ": Exit Sub

            'WithoutOpenenig.value = vbChecked
            If WithoutOpenenig.value = vbUnchecked Then
                withouopenening = False
            Else
                withouopenening = True
            End If
            
            
             
            
            If Grid3.rows = 1 Then
                '**********One Account ******************************
                If CHECK_LAST_ACCOUNT(StrAccountCode) = False Then MsgBox " «·Õ”«» —∆Ì”Ì ": Exit Sub

                 If chkContinue.value = vbUnchecked Then
                    updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText), withouopenening
                Else
               
                    updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText), withouopenening
                End If
               
                '***************************************
            Else
                
                For Row = 1 To Grid3.rows - 1
                    
                    If val(Grid3.TextMatrix(Row, Grid3.ColIndex("Sel"))) <> 0 Then
                       StrAccountCode = Grid3.TextMatrix(Row, Grid3.ColIndex("Account_Code"))
                    
                       If CHECK_LAST_ACCOUNT(StrAccountCode) = False Then GoTo lblNext 'MsgBox "???? ?? ?????? ???? ?????  ": Exit Sub
                        
                       AccColl.Add "'" & StrAccountCode & "'"
                        
                       If chkContinue.value = vbUnchecked Then
                           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText), withouopenening
                       Else
                           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , , val(DCRegionID.BoundText), withouopenening
                       End If
                    End If
lblNext:
                Next
            End If
            If Grid3.rows > 1 And AccColl.count = 0 Then
                MsgBox "·« ÌÊÃœ Õ”«» Ì‰ÿ»ﬁ ⁄·ÌÂ «·‘—Êÿ"
                Exit Sub
            End If
            
            
            
            
          

            CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName, , True)
      
            If CustomerAgeingData <> "" Then
                If ViewAging = True Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
                    Else
                        X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
                    End If

                    If X = vbCancel Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
    
                    If X = vbYes Then
                        ShowAgingReport = "1"
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    Else
                        ShowAgingReport = 0
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    End If

                End If

            End If
            If i = 0 Then
                If Text1.text = "" Then
                    Set cAccountReport = New ClsAccReports
                    cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                    cAccountReport.EndDate = Me.DTPickerAccTo.value
                    updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
  
  
                    If AccColl.count > 0 Then
                         StrAccountCode = Join(collectionToArray(AccColl), ",")
                         cAccountReport.ShowLedger StrAccountCode, StrAccountName, True, , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, , , , , val(DCRegionID.BoundText), withouopenening, val(DCNotesTypes.BoundText)
                     Else
                         'cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, Text1.text
                         cAccountReport.ShowLedger StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, , , , , val(DCRegionID.BoundText), withouopenening, val(DCNotesTypes.BoundText)
                     End If
            
               
                Else
  
                    Set cAccountReport = New ClsAccReports
                    cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                    cAccountReport.EndDate = Me.DTPickerAccTo.value
                    updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
  
            If AccColl.count > 0 Then
                    StrAccountCode = Join(collectionToArray(AccColl), ",")
                    cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, Text1.text, True
                Else
                    cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, Text1.text
            End If
            
            Set cAccountReport = Nothing
            
            
           
            
                End If
            Else
                Set cAccountReport = New ClsAccReports
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
                cAccountReport.ShowLedger22 StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, , , , , val(DCRegionID.BoundText)
                Set cAccountReport = Nothing
            End If
        Case 36
          
            If txt_mod_flag.text = "N" Then
                '??? ????
            
                If Me.TrvAccounts.SelectedItem Is Nothing Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                If Me.TxtAccountCode.text <> "" Then
        
                Else
                    StrAccountCode = Me.TrvAccounts.SelectedItem.key
                    StrAccountName = Me.TrvAccounts.SelectedItem.text
                End If
            End If
    
            If StrAccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ «·Õ”«»", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql2 DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            Else
                updateopeningbalanceNewFromsql2 DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            End If

            CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName, , True)
      
            If CustomerAgeingData <> "" Then
                If ViewAging = True Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("«ŸÂ«— «⁄„«· «·œÌÊ‰ ", vbCritical + vbYesNoCancel)
                    Else
                        X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
                    End If

                    If X = vbCancel Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
    
                    If X = vbYes Then
                        ShowAgingReport = "1"
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    Else
                        ShowAgingReport = 0
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    End If

                End If

            End If
 
            If Text1.text = "" Then
                Set cAccountReport = New ClsAccReports
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
                cAccountReport.ShowLedger StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, , , , 1
                Set cAccountReport = Nothing
            Else
  
                Set cAccountReport = New ClsAccReports
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
                cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, Text1.text, , 1
                Set cAccountReport = Nothing
            
            End If

        Case 32 '?????? ??????
       
            If txt_mod_flag.text = "N" Then
      
                If Me.TxtAccountCode.text <> "" Then
        
                Else
                    StrAccountCode = Me.TrvAccounts.SelectedItem.key
                    If mId(StrAccountCode, Len(StrAccountCode), 1) = "G" Then
                        StrAccountCode = mId(StrAccountCode, 1, Len(StrAccountCode) - 1)
                    
                    End If
                    StrAccountName = Me.TrvAccounts.SelectedItem.text
                End If
            End If
    
            If StrAccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ «·Õ”«»", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If
       
            'updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), GetAlLastAccounts(StrAccountCode), 5, , True
            updateopeningbalanceNewFromsqlTrialBalance2 DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), GetAlLastAccounts(StrAccountCode), 5, , True
    
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            '           updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            ' updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTEnddate.value
            
            cAccountReport.ShowLedgers GetAlLastAccounts(StrAccountCode), "", Text1.text, True
            Set cAccountReport = Nothing

        Case 27
            '*********************************************************************
 
            '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
            '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate
            '    update_account_opening_balance StrAccountCode, True, DTPickerAccFrom.value, DTPickerAccTo.value, Val(dcBranch.BoundText), openingbalacedate
                 
            If txt_mod_flag.text = "N" Then
                '??? ????
            
                If Me.TrvAccounts.SelectedItem Is Nothing Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                If Me.TxtAccountCode.text <> "" Then
        
                Else
                    StrAccountCode = Me.TrvAccounts.SelectedItem.key
                    StrAccountName = Me.TrvAccounts.SelectedItem.text
                End If
            End If
    
            If StrAccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "??? ????", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            End If

            CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName, , True)
      
            If CustomerAgeingData <> "" Then
                If ViewAging = True Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
                    Else
                        X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
                    End If

                    If X = vbCancel Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
    
                    If X = vbYes Then
                        ShowAgingReport = "1"
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    Else
                        ShowAgingReport = 0
                        CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

                    End If

                End If

            End If
       
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
            cAccountReport.ShowLedger StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, , , 1, , , , , IIf(chkIsBasicInvoice.value = vbChecked, 1, 0)
            
            Set cAccountReport = Nothing

            '******************************************************************************************

            '******************************************************************************************
        Case 26
    
            If dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "??? ????? ????", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            StrAccountCode = get_project_customer_account(val(dcprojects.BoundText), "Account_Code")
            If StrAccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "??? ????", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , val(Me.dcprojects.BoundText)
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3, , val(Me.dcprojects.BoundText)
            End If
             
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
            cAccountReport.ShowLedger StrAccountCode, StrAccountName, , , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode, val(Me.dcprojects.BoundText), dcprojects.text
            Set cAccountReport = Nothing
            
        Case 34
      
            EmployeeBefnet
            '*******************************************************************************************
        Case 14
            CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName, True)
            Set cAccountReport = New ClsAccReports
            cAccountReport.ShowAgingReport WindowTarget
            Set cAccountReport = Nothing
        
        Case 1

            '???? ????? ???
            '        updateopeningbalance
            If Grid3.rows = 1 And Trim(TxtAccountCode) = "" Then
             
                If Me.TrvAccounts.SelectedItem Is Nothing Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ???????? "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If

            Set cAccountReport = New ClsAccReports
            ' StrAccountCode = Me.TrvAccounts.SelectedItem.key
            '    StrAccountCode = Mid$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            If right$(StrAccountCode, 1) = "G" Then
                StrAccountCode = mId$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            End If
            
            
            If Grid3.rows = 1 Then
                '**********One Account ******************************
                If CHECK_LAST_ACCOUNT(StrAccountCode) = True Then MsgBox "???? ?? ?????? ???? ?????  ": Exit Sub

                If chkContinue.value = vbUnchecked Then
                    updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
                Else
                    updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
                End If
               
                '***************************************
            Else
                
                For Row = 1 To Grid3.rows - 1
                    
                    If val(Grid3.TextMatrix(Row, Grid3.ColIndex("Sel"))) <> 0 Then
                       StrAccountCode = Grid3.TextMatrix(Row, Grid3.ColIndex("Account_Code"))
                    
                       If CHECK_LAST_ACCOUNT(StrAccountCode) = False Then GoTo lblNext2 'MsgBox "???? ?? ?????? ???? ?????  ": Exit Sub
                        
                       AccColl.Add "'" & StrAccountCode & "'"
                        
                       If chkContinue.value = vbUnchecked Then
                           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
                       Else
                           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
                       End If
                    End If
lblNext2:
                Next
            End If
            If Grid3.rows > 1 And AccColl.count = 0 Then
                MsgBox "·« ÌÊÃœ Õ”«» Ì‰ÿ»ﬁ ⁄·ÌÂ «·‘—Êÿ"
                Exit Sub
            End If
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            If AccColl.count > 0 Then
                StrAccountCode = Join(collectionToArray(AccColl), ",")
            End If
            cAccountReport.ShowGenrealLedgertnew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), StrAccountCode, LblAccountName
            Set cAccountReport = Nothing
        
            Exit Sub

        Case 3
            '????? ?????
            '        updateopeningbalance
            '   Set cAccountReport = New ClsAccReports
            '   cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            '       cAccountReport.Enddate = Me.DTPickerAccTo.value
            '
            '   cAccountReport.ShowIncomeStatment Val(Me.DcBranch.BoundText), Val(Me.DCActivity.BoundText)
            '   Set cAccountReport = Nothing
    
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 2
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 2
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowIncomeStatementnew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText)
            Set cAccountReport = Nothing
 
        Case 28
            '?????? ????? ?????
            '        updateopeningbalance
            '   Set cAccountReport = New ClsAccReports
            '   cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            '       cAccountReport.Enddate = Me.DTPickerAccTo.value
            '
            '   cAccountReport.ShowIncomeStatment Val(Me.DcBranch.BoundText), Val(Me.DCActivity.BoundText)
            '   Set cAccountReport = Nothing
    
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 2
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 2
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowIncomeStatementnew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , True
            Set cAccountReport = Nothing
            
        Case 4
            '????????????
            '        updateopeningbalance
        
            If SystemOptions.UserInterface = ArabicInterface Then
                X = val(InputBox("Õœœ «·„” ÊÏ"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            account_level = val(X)
            Dim HideZeroBalance As Integer
            Dim HideMasterAcc   As Integer
        
            If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("«Œ›«¡ «·Õ”«»«  «·’›—Ì… ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '   updateopeningbalance
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 0
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 0
            End If
            
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowBalanceSheet val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance
            Set cAccountReport = Nothing
        
        Case 5
  
            '????? ??????
            '        updateopeningbalance
            '            If SystemOptions.UserInterface = ArabicInterface Then
            '                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            '            Else
            '                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            '            End If
            '
            '            If HideZeroBalance = vbCancel Then
            '                Screen.MousePointer = vbDefault
            '                Exit Sub
            '            End If
            Dim openingBalanceDate As Date
            Dim FromdateMinus1     As Date
            Dim StartCurrentDate   As Date
            Dim BrcnActivety       As String
            FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
            getFirstPeriodDateInthisYear2 openingBalanceDate
            getFirstPeriodDateInthisYear StartCurrentDate
            HideZeroBalance = 7
            '         If SystemOptions.UserInterface = ArabicInterface Then
            '                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            '            Else
            '                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            '            End If
            '
            '            If HideZeroBalance = 2 Then
            '                Screen.MousePointer = vbDefault
            '                Exit Function
            '            End If
      
            If val(DCRegionID.BoundText) <> 0 Then
                BranshesReg = BranchRegion(DCRegionID.BoundText)
            End If
            If val(DCActivity.BoundText) <> 0 Then
                BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
            End If

            Dim s As String
            s = "Select * from TblyearsData  where IsNull(IsFirstYear,0) = 1 and YEAR(datesatrt)  = " & year(val(DTPickerAccFrom.value))
            Dim rsDummy As New ADODB.Recordset
            rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
            If rsDummy.EOF Then
                mIsFirstYear = False
            Else
                mIsFirstYear = True
            End If

            ' updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 1
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 1
            End If
            
            '    updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance
       
            Set cAccountReport = Nothing
        Case 33 '????? ????
            print_report40New
        Case 5
  
            '????? ??????
            '        updateopeningbalance
            If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("«Œ›«¡ «·Õ”«»«  «·’›—Ì…", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 1
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 1
            End If
            
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance
       
            Set cAccountReport = Nothing
 
        Case 7

            '???? ??????? ?? ????
            If ChkNotesType.value = vbChecked And DCNotesTypes.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "????  ??????  ", vbCritical
                Else
                    MsgBox "Specify Transaction", vbCritical
                End If

                DCNotesTypes.SetFocus
                Sendkeys ("{F4}")
                Exit Sub
 
            End If
   
            ShowGl val(DCNotesTypes.BoundText), val(Me.dcBranch.BoundText)
        Case 8

            If ChkNotesType.value = vbChecked And DCNotesTypes.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "????  ??????  ", vbCritical
                Else
                    MsgBox "Specify Transaction", vbCritical
                End If

                DCNotesTypes.SetFocus
                Sendkeys ("{F4}")
                Exit Sub
 
            End If

            '???? ??????? ?????? ?????? ??????? ?? ????
            ShowGLWITH_Cost_center val(DCNotesTypes.BoundText), val(Me.dcBranch.BoundText)

        Case 9
            ShowTransactionsWith_Cost_center StrAccountCode, DcCostCenter.BoundText
   
        Case 15
            ShowTransactionsWith_Car StrAccountCode, val(Me.DcFixedAssets.BoundText)

        Case 21
            ShowTransactionsWith_Departement StrAccountCode, val(Me.DcboEmpDepartments.BoundText)
            
        Case 23
            ShowTransactionsWith_Employee StrAccountCode, val(Me.DCEmployee.BoundText)
                   
        Case 16
            ShowCarProfits StrAccountCode, val(Me.DcFixedAssets.BoundText)

        Case 17
            ShowCarExpensess val(Me.DcFixedAssets.BoundText)
   
        Case 18 '????? ?????? ??????????
   
            If SystemOptions.UserInterface = ArabicInterface Then
                X = val(InputBox("??? ???????"))
            Else
                X = val(InputBox("Specify Level"))
            End If
            
            account_level = val(X)
             
            If account_level = 0 Or account_level >= getLastLevel Then
                ShowLastAccount = True
            Else
                ShowLastAccount = False
            End If
   
            If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5, , ShowLastAccount, val(DCRegionID.BoundText)
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5, , ShowLastAccount, val(DCRegionID.BoundText)
            End If
 
            If val(Me.Txtyear.text) > 0 Then
                updateAccountsmanully val(Me.Txtyear)
            End If
            
            If val(DCRegionID.BoundText) <> 0 Then
                BranshesReg = BranchRegion(DCRegionID.BoundText)
            End If
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            '       If HideMasterAcc = True Then
            cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , False, , ShowLastAccount
            '       Else
            '       cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , False, , showlastaccount
            '       End If
            Set cAccountReport = Nothing
  
        Case 25  '????? ?????? ??????????
            If right$(StrAccountCode, 1) = "G" Then
                StrAccountCode = mId$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            End If
            If SystemOptions.UserInterface = ArabicInterface Then
                X = val(InputBox("??? ???????"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            account_level = val(X)
   
            If account_level = 0 Or account_level >= getLastLevel Then
                ShowLastAccount = True
            Else
                ShowLastAccount = False
            End If
            
            If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            
            End If
  
            If SystemOptions.UserInterface = ArabicInterface Then
                HideMasterAcc = MsgBox("?? ???? ????? ????????  ??????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            Else
                HideMasterAcc = MsgBox("Hide Master Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideMasterAcc = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
              
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 5, , ShowLastAccount
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 5, , ShowLastAccount
            End If
            
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            'If HideMasterAcc = True Then
            cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , False, StrAccountCode, ShowLastAccount, HideMasterAcc
            'Else
            'cAccountReport.ShowTrialBalanceNew val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , False, StrAccountCode, showlastaccount
            'End If
            Set cAccountReport = Nothing
            
        Case 10
            If dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·„‘—Ê⁄ «Ê·«", vbCritical
            
                Else
                    MsgBox "Select Project Firstly", vbCritical
                End If
                dcprojects.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
        
            Else
                ShowGLto_project val(Me.dcprojects.BoundText), val(Dcdetails.BoundText), , DCAccounts.BoundText, val(DcbProcess1.BoundText)
            End If
        Case 30
            
            If dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·„‘—Ê⁄ «Ê·«", vbCritical
            
                Else
                    MsgBox "Select Project Firstly", vbCritical
                End If
                dcprojects.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
        
            Else
                ShowGLto_project val(Me.dcprojects.BoundText), val(Dcdetails.BoundText), True, DCAccounts.BoundText
            End If
       
        Case 42
     
            If txt_mod_flag.text = "N" Then
                '??? ????
            
                If Me.TrvAccounts.SelectedItem Is Nothing Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf InStr(1, Me.TrvAccounts.SelectedItem.Tag, "last", vbTextCompare) = 0 Then
                    Msg = "??? ?????? ??? ?????? ??????" & CHR(13) & "?????? ??? ??????? ?? ?? ???? ?????? ????????"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                If Me.TxtAccountCode.text <> "" Then
        
                Else
                    StrAccountCode = Me.TrvAccounts.SelectedItem.key
                    
                    StrAccountName = Me.TrvAccounts.SelectedItem.text
                    If StrAccountName = "" Then
                        StrAccountName = LblAccountName
                    End If
                End If
            End If
    
            If StrAccountCode = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ «·Õ”«»", vbCritical
                Else
                    MsgBox "  Specify     Account", vbCritical
                End If

                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            ShowGLto_projectNew val(Me.dcprojects.BoundText), val(Dcdetails.BoundText), True, DCAccounts.BoundText
      
        Case 11
            '  ??????????
            '        updateopeningbalance
         
            '       If SystemOptions.UserInterface = ArabicInterface Then
            '           X = val(InputBox("??? ???????"))
            '       Else
            '           X = val(InputBox("Specify Level"))
            '       End If
            '
            '            account_level = val(X)
            '
            '            If SystemOptions.UserInterface = ArabicInterface Then
            '                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            '            Else
            '                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            ''            End If
    
            '           If HideZeroBalance = vbCancel Then
            '               Screen.MousePointer = vbDefault
            '               Exit Sub
            '           End If
            '
            '   updateopeningbalance
            If chkContinue.value = vbUnchecked Then
                '           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5
            Else
                '           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
           
            cAccountReport.ShowOpeningBalances2 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance
            Set cAccountReport = Nothing
    
        Case 30
            
            If dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·„‘—Ê⁄ «Ê·«", vbCritical
            
                Else
                    MsgBox "Select Project Firstly", vbCritical
                End If
                dcprojects.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
        
            Else
                ShowGLto_project val(Me.dcprojects.BoundText), val(Dcdetails.BoundText), True, DCAccounts.BoundText
            End If
       
        Case 42
            
            If dcprojects.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·„‘—Ê⁄ «Ê·«", vbCritical
            
                Else
                    MsgBox "Select Project Firstly", vbCritical
                End If
                dcprojects.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
        
            Else
                ShowGLto_project val(Me.dcprojects.BoundText), val(Dcdetails.BoundText), True, DCAccounts.BoundText
            End If
       
        Case 11
            '  ??????????
            '        updateopeningbalance
         
            '       If SystemOptions.UserInterface = ArabicInterface Then
            '           X = val(InputBox("??? ???????"))
            '       Else
            '           X = val(InputBox("Specify Level"))
            '       End If
            '
            '            account_level = val(X)
            '
            '            If SystemOptions.UserInterface = ArabicInterface Then
            '                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
            '            Else
            '                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            ''            End If
    
            '           If HideZeroBalance = vbCancel Then
            '               Screen.MousePointer = vbDefault
            '               Exit Sub
            '           End If
            '
            '   updateopeningbalance
            If chkContinue.value = vbUnchecked Then
                '           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5
            Else
                '           updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), , 5
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            
            cAccountReport.ShowOpeningBalances2 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance
            Set cAccountReport = Nothing

        Case 12
  
            ' ????? ??? ???????? ??? ....
            '        updateopeningbalance
            '   StrAccountCode = Me.TrvAccounts.SelectedItem.key
            '   StrAccountCode = Mid$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            If right$(StrAccountCode, 1) = "G" Then
                StrAccountCode = mId$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            End If
            If CHECK_LAST_ACCOUNT(StrAccountCode) = True Then MsgBox "«·Õ”«» ‰Â«∆Ï  ": Exit Sub
        
            If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("«Œ›«¡ «·Õ”«»«  «·’›—Ì… ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 2
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 2
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            cAccountReport.ShowTrialBalanceNew2 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, StrAccountCode, , LblAccountName
       
            Set cAccountReport = Nothing
        
        Case 19

            If val(DCCompositeAccount.BoundText) = 0 Then
                Msg = "??? ?????? ???    ?????? ??????  " & CHR(13) & "?????? ??? ??????? ?? ??       "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCCompositeAccount.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                
            StrAccountCode = GetAllCompositeAccounts(val(DCCompositeAccount.BoundText))
        
            ' HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNo)
            '           If chkContinue.value = vbUnchecked Then
            '                      updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText) _
            '                        , val(Me.Dcbranch.BoundText), StrAccountCode, 3
            '           Else
            '                   updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText) _
            '                           , val(Me.Dcbranch.BoundText), StrAccountCode, 3
            '           End If
                
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 30
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 30
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
            cAccountReport.ShowLedgerComposite StrAccountCode, DCCompositeAccount.text, True, , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode
            Set cAccountReport = Nothing
        
        Case 31
       
            If val(DCEmployee.BoundText) = 0 Then
                Msg = "??? ?????? ???      ??????  " & CHR(13) & "?????? ??? ??????? ?? ??       "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCEmployee.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                
            StrAccountCode = GetAllEmployeeAccounts(val(DCEmployee.BoundText))
                
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 30
            Else
                updateopeningbalanceNewFromsql DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 30
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
            updateprofitAccount val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), Me.DTPickerAccTo.value
  
            cAccountReport.ShowLedgerComposite StrAccountCode, DCEmployee.text, True, , , val(FrmAccountingReport.DCActivity.BoundText), val(FrmAccountingReport.dcBranch.BoundText), CustomerAgeingData, salesPersonName, ShowAgingReport, TxtAccountCode
            Set cAccountReport = Nothing
            
        Case 13
  
            '??? ???? ????
            '        updateopeningbalance
            'StrAccountCode = Me.TrvAccounts.SelectedItem.key
            'StrAccountCode = Mid$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            If val(DCEmployee.BoundText) = 0 Then
                Msg = "??? ?????? ???    ?????? " & CHR(13) & "?????? ??? ??????? ?? ??       "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCEmployee.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
            StrAccountCode = GetAllEmployeeAccounts(val(DCEmployee.BoundText))
        
            ' HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNo)
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 3
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            cAccountReport.ShowTrialBalanceNew2 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , StrAccountCode
       
            Set cAccountReport = Nothing
        
        Case 35
            print_report2
        
        Case 39
            print_report40
            
        Case 41
            print_report41
        
        Case 38
            print_report3
 
        Case 20
  
            '??? ???? ?????
            '        updateopeningbalance
            'StrAccountCode = Me.TrvAccounts.SelectedItem.key
            'StrAccountCode = Mid$(Me.TrvAccounts.SelectedItem.key, 1, Len(Me.TrvAccounts.SelectedItem.key) - 1)
            If val(dcprojects.BoundText) = 0 Then
                Msg = "??? ?????? ???    ???????? " & CHR(13) & "?????? ??? ??????? ?? ??       "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcprojects.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
            StrAccountCode = GetAllprojectsAccounts(val(dcprojects.BoundText))
        
            ' HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNo)
            If chkContinue.value = vbUnchecked Then
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, False, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
            Else
                updateopeningbalanceNewFromsqlTrialBalance DTPickerAccFrom.value, DTPickerAccTo.value, True, val(Me.DCActivity.BoundText), val(Me.dcBranch.BoundText), StrAccountCode, 4
            End If
            
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
          
            ' cAccountReport.ShowTrialBalance Val(Me.dcBranch.BoundText), Val(Me.DCActivity.BoundText)
            cAccountReport.ShowTrialBalanceNew3 val(Me.dcBranch.BoundText), val(Me.DCActivity.BoundText), , account_level, HideZeroBalance, , StrAccountCode, Me.dcprojects.text
       
            Set cAccountReport = Nothing
    End Select

    CuurentLogdata
End Sub
 

' Requires: Microsoft ActiveX Data Objects 2.x
Private Function AccountsCollectionToCSV(ByVal AccColl As Collection, _
                                        ByVal FallbackSingle As String) As String
    Dim v As Variant, s As String, i As Long
    If AccColl Is Nothing Or AccColl.count = 0 Then
        AccountsCollectionToCSV = Trim$(FallbackSingle)
        Exit Function
    End If
    Dim arr() As String: ReDim arr(1 To AccColl.count)
    i = 1
    For Each v In AccColl
        s = CStr(v)
        If Len(s) >= 2 And left$(s, 1) = "'" And right$(s, 1) = "'" Then
            s = mId$(s, 2, Len(s) - 2)   ' strip quotes
        End If
        arr(i) = Trim$(s)
        i = i + 1
    Next
    AccountsCollectionToCSV = Join(arr, ",")
End Function
Private Function RunLedgerSP_ReturnRS( _
    ByVal AccountCSV As String, _
    ByVal StartDate As Date, ByVal EndDate As Date, _
    ByVal BranchID As Long, ByVal ActivityId As Long, ByVal RegionID As Long, _
    ByVal projectId As Long, ByVal notesType As Double, _
    ByVal ShowHidden As Boolean, ByVal IsDetailed As Boolean, _
    ByVal includeSeedRow As Boolean, ByVal ComputeOpeningInSP As Boolean, _
    ByVal OpeningMode As Byte) As ADODB.Recordset

    Dim Cmd As ADODB.Command, p As ADODB.Parameter, rs As ADODB.Recordset
    Set Cmd = New ADODB.Command
    

    
Dim computeOpening As Long
includeSeedRow = IIf(WithoutOpenenig.value = vbUnchecked, 1, 0)  ' ·Ê «Œ —  "»œÊ‰ «›  «ÕÌ" Ì»ﬁÏ 0
computeOpening = 1    ' «Õ”» «·«›  «ÕÌ œ«Œ· «·‹SP œ«Ì„«

With Cmd
    .ActiveConnection = Cn
    .CommandType = adCmdStoredProc
    .NamedParameters = True
    .CommandText = "dbo.sp_Get_AccountLedger_ReportData"
    .CommandTimeout = 180

    ' NVARCHAR(MAX)
    .Parameters.Append .CreateParameter("@AccountCodes", adLongVarWChar, adParamInput, -1, AccountCSV)

    .Parameters.Append .CreateParameter("@StartDate", adDBDate, adParamInput, , StartDate)
    .Parameters.Append .CreateParameter("@EndDate", adDBDate, adParamInput, , EndDate)

    .Parameters.Append .CreateParameter("@BranchID", adInteger, adParamInput, , CLng(BranchID))
    .Parameters.Append .CreateParameter("@ActivityId", adInteger, adParamInput, , CLng(ActivityId))
    .Parameters.Append .CreateParameter("@RegionID", adInteger, adParamInput, , CLng(RegionID))
    .Parameters.Append .CreateParameter("@ProjectID", adInteger, adParamInput, , CLng(projectId))

    .Parameters.Append .CreateParameter("@NotesType", adDouble, adParamInput, , CDbl(notesType))

    ' BIT ﬂ‹ 0/1 („‘ adBoolean)
    .Parameters.Append .CreateParameter("@ShowHiddenInvoices", adInteger, adParamInput, , IIf(ShowHidden, 1, 0))
    .Parameters.Append .CreateParameter("@IsDetailedReport", adInteger, adParamInput, , IIf(IsDetailed, 1, 0))
    .Parameters.Append .CreateParameter("@IncludeSeedRow", adInteger, adParamInput, , includeSeedRow)
    .Parameters.Append .CreateParameter("@ComputeOpeningInSP", adInteger, adParamInput, , computeOpening)

    .Parameters.Append .CreateParameter("@OpeningMode", adTinyInt, adParamInput, , CByte(OpeningMode))
End With


    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Cmd, , adOpenStatic, adLockReadOnly
    Set RunLedgerSP_ReturnRS = rs
End Function

Function ShowTransactionsWith_Cost_center(Optional Account_code As String = "", Optional cost_center_id As String = "")
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Cn.CommandTimeout = 0
    
    'MySQL = "SELECT     *, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN CC_ValIe * 1 ELSE 0 END, DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN CC_ValIe * 1 ELSE 0 END FROM    GL_CC where not(cost_center_id is null)"
    'MySQL = "Select * From GL_CC where not(cost_center_id is null)"

'    MySQL = "SELECT     dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.marakes_taklefa_temp.cost_center_id, "
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.cost_center, dbo.marakes_taklefa_temp.Project__code, dbo.marakes_taklefa_temp.Project_name,"
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.[value] AS CC_Valie, dbo.marakes_taklefa_temp.Description, dbo.marakes_taklefa_temp.id, dbo.marakes_taklefa_temp.opr_id,"
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.opr_type, dbo.marakes_taklefa_temp.depit_or_credit, dbo.marakes_taklefa_temp.account_type, dbo.marakes_taklefa_temp.account_no,"
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.line_no, dbo.marakes_taklefa_temp.kedno, dbo.marakes_taklefa_temp.Foxy_no, dbo.marakes_taklefa_temp.user_id,"
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.ok, dbo.marakes_taklefa_temp.record_date, dbo.marakes_taklefa_temp.general_des, dbo.marakes_taklefa_temp.auto_des,"
'    MySQL = MySQL & " dbo.marakes_taklefa_temp.notedate , dbo.marakes_taklefa_temp.NoteSerial, dbo.marakes_taklefa_temp.remark, dbo.ACCOUNTS.Account_Code"
'    MySQL = MySQL & " FROM         dbo.marakes_taklefa_temp INNER JOIN"
'    MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.marakes_taklefa_temp.account_no = dbo.ACCOUNTS.Account_Code"
'    MySQL = MySQL & "  Where (dbo.marakes_taklefa_temp.ok = 1) And (Not (dbo.marakes_taklefa_temp.cost_center_id Is Null))"
MySQL = "SELECT     dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.marakes_taklefa_temp.cost_center_id, "
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.cost_center, dbo.marakes_taklefa_temp.Project__code, dbo.marakes_taklefa_temp.Project_name,"
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.[value] AS CC_Valie, dbo.marakes_taklefa_temp.Description, dbo.marakes_taklefa_temp.id, dbo.marakes_taklefa_temp.opr_id,"
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.opr_type, dbo.marakes_taklefa_temp.depit_or_credit, dbo.marakes_taklefa_temp.account_type, dbo.marakes_taklefa_temp.account_no,"
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.line_no, dbo.marakes_taklefa_temp.kedno, dbo.marakes_taklefa_temp.Foxy_no, dbo.marakes_taklefa_temp.user_id,"
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.ok, dbo.marakes_taklefa_temp.record_date, dbo.marakes_taklefa_temp.general_des, dbo.marakes_taklefa_temp.auto_des,"
MySQL = MySQL & "                       dbo.marakes_taklefa_temp.NoteDate, dbo.marakes_taklefa_temp.NoteSerial, dbo.marakes_taklefa_temp.Remark, dbo.ACCOUNTS.Account_Code,"
MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,TblAqar.aqarname,markaas_taklefa.akarid "
MySQL = MySQL & " FROM         dbo.marakes_taklefa_temp INNER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.marakes_taklefa_temp.account_no = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.marakes_taklefa_temp.line_no = dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1"
MySQL = MySQL & "                       Left outer join markaas_taklefa On marakes_taklefa_temp.cost_center_id =markaas_taklefa.Code "
MySQL = MySQL & "                       Left outer join TblAqar On TblAqar.Aqarid =markaas_taklefa.akarid "

MySQL = MySQL & "  Where (     dbo.marakes_taklefa_temp.line_no<>0 and dbo.marakes_taklefa_temp.ok = 1) And (Not (dbo.marakes_taklefa_temp.cost_center_id Is Null))"


    If cost_center_id <> "" Then
        MySQL = MySQL + " and marakes_taklefa_temp.cost_center_id='" & cost_center_id & "'"
    End If

    If Account_code <> "" Then
        MySQL = MySQL + " and dbo.ACCOUNTS.Account_Code ='" & Account_code & "'"
    End If
    If val(DcbAqar.BoundText) <> 0 Then
        MySQL = MySQL + " and markaas_taklefa.akarid =" & val(DcbAqar.BoundText)
    End If
    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  NoteDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and NoteDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If
    
       If Trim(cmbAccount.BoundText) <> "" And Trim(cmbAccount.text) <> "" Then
            MySQL = MySQL & "   and  ACCOUNTS.Account_Code IN (SELECT Code"
            MySQL = MySQL & "                     FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(cmbAccount.BoundText) & "', '" & Trim(cmbAccount.BoundText) & "', 1))"
            MySQL = MySQL & "  OR (ACCOUNTS.Account_Code = '" & Trim(cmbAccount.BoundText) & "')"
        End If
    MySQL = MySQL + " Order By NoteDate"
   
    Dim X As Integer

If chKCCTotals.value = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalsN1.rpt"
        Else
           ' StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalse.rpt"
           StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalsNe1.rpt"
           
        End If
GoTo CCTotals
End If

    If SystemOptions.UserInterface = ArabicInterface Then
         X = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
    Else
        X = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
    End If

    If X = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_centerE.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalsN.rpt"
        Else
           ' StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalse.rpt"
           StrFileName = App.path & "\Reports\" & "Transactions_with_cost_center_totalsNe.rpt"
           
        End If
    End If

CCTotals:
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
     RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
DoEvents
    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
Function ShowCarExpensess(Optional CarID As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
'    MySQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Totals, dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, "
'    MySQL = MySQL & "  dbo.TblCarsData.BoardNO , dbo.ACCOUNTS.account_name"
'    MySQL = MySQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
'    MySQL = MySQL & " dbo.ExpensesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
'    MySQL = MySQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
'    MySQL = MySQL & " dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id"
'    'MySQL = MySQL & "  WHERE 1=1 and dbo.ACCOUNTS.AccountTypes = 2"
 
 MySQL = "SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Totals, dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, "
   MySQL = MySQL & "                     dbo.TblCarsData.BoardNO , dbo.Accounts.account_name"
MySQL = MySQL & "  FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
MySQL = MySQL & "                        dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id"
    MySQL = MySQL & "  WHERE 1=1 and dbo.ACCOUNTS.AccountTypes = 2"
    If CarID <> 0 Then
        MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = " & CarID & ") "
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    MySQL = MySQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.ACCOUNTS.Account_Name"

    'If Account_Code <> "" Then
    'MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.account_code='" & Account_Code & "'"
    'End If
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "\Transporter\CarsExpenses.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "\Transporter\CarsExpenses.rpt"
    End If
    
    Dim X As Integer

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?????? ?????? ?????"
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
    StrReportTitle = "????????   ???????? "

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "????????   ???????? " & CHR(13)

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & "" & CHR(13)
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " " & CHR(13)
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL
    

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
Function ShowCarProfits(Optional Account_code As String = "", Optional CarID As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
    MySQL = "   select BoardNO,sum(Debtit) as totalDebit ,sum(Credit)   totalCredit,sum(Credit)-sum(Debtit)   as profit"
    MySQL = MySQL & " from("
    MySQL = MySQL & " SELECT       dbo.TblCarsData.BoardNO,"
    MySQL = MySQL & " 'Debtit' = CASE"
    MySQL = MySQL & " WHEN Credit_Or_Debit = 0  THEN SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value])"
    MySQL = MySQL & " end ,"
    MySQL = MySQL & " 'Credit' = CASE"
    MySQL = MySQL & " WHEN Credit_Or_Debit = 1  THEN SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value])"
    MySQL = MySQL & " End"
    MySQL = MySQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    MySQL = MySQL & "  dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id"
    MySQL = MySQL + "  where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code LIKE N'a3%' OR  dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code LIKE N'a4%')"

    If CarID <> 0 Then
        MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = " & CarID & ") "
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    MySQL = MySQL + " GROUP BY dbo.TblCarsData.BoardNO,Credit_Or_Debit"
    MySQL = MySQL + " )Tablex"
    MySQL = MySQL + " GROUP BY  BoardNO"

    'If Account_Code <> "" Then
    'MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.account_code='" & Account_Code & "'"
    'End If
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "\Transporter\Carstotals1.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "\Transporter\Carstotals1.rpt"
    End If
    
    Dim X As Integer

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?????? ?????? ?????"
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
    StrReportTitle = "???????? ?????????? ???????? "

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "???????? ?????????? ???????? " & CHR(13)

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & "" & CHR(13)
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " " & CHR(13)
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

'xxxxxxx
Function ShowTransactionsWith_Departement(Optional Account_code As String = "", Optional Departementid As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
 
    MySQL = " SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Code, "
MySQL = MySQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate,"
MySQL = MySQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.NoteType,"
MySQL = MySQL + "                       dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1,"
MySQL = MySQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Departementid , dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee"
MySQL = MySQL + " FROM         dbo.ACCOUNTS INNER JOIN"
MySQL = MySQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
MySQL = MySQL + "                       dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
MySQL = MySQL + "                       dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL + "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
MySQL = MySQL + " WHERE     (1 = 1)"
  
    
'FixedAssetId
    If Departementid <> 0 Then
        MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = " & Departementid & ") "
    End If

    If Account_code <> "" Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.account_code='" & Account_code & "'"
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    MySQL = MySQL + " Order By dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate "
   
    Dim X As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("?? ???? ????? ?????? ??? ?? ??", vbExclamation + vbYesNo)
    Else
        X = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
    End If

    If X = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_Departement.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_Departement.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_departemetsTotals.rpt"
        Else
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_departemetsTotals.rpt"
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
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
 
Function ShowTransactionsWith_Employee(Optional Account_code As String = "", Optional NEmpid As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
 '    MySQL = " SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Code, "
'MySQL = MySQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate,"
'MySQL = MySQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.NoteType,"
'MySQL = MySQL + "                       dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1,"
'MySQL = MySQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Departementid , dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee"
'MySQL = MySQL + " FROM         dbo.ACCOUNTS INNER JOIN"
'MySQL = MySQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
'MySQL = MySQL + "                       dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
'MySQL = MySQL + "                       dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
'MySQL = MySQL + "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
'MySQL = MySQL + " WHERE     (1 = 1)"
  MySQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Code, "
MySQL = MySQL + "                          dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate,"
MySQL = MySQL + "                          dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.NoteType,"
MySQL = MySQL + "                          dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1,"
MySQL = MySQL + "                          dbo.DOUBLE_ENTREY_VOUCHERS.Departementid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL + "                          dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid , dbo.TblEmployee.emp_Name, dbo.TblEmployee.emp_code, dbo.TblEmployee.Emp_Namee"
MySQL = MySQL + "    FROM         dbo.ACCOUNTS INNER JOIN"
MySQL = MySQL + "                          dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
MySQL = MySQL + "                          dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID INNER JOIN"
MySQL = MySQL + "                          dbo.TblEmployee ON dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL + "                          dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL + "                          dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
MySQL = MySQL + "    Where (1 = 1)"

    
'FixedAssetId
    If NEmpid <> 0 Then
        MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.NEmpid = " & NEmpid & ") "
    End If

    If Account_code <> "" Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.account_code='" & Account_code & "'"
    End If

    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    MySQL = MySQL + " Order By dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate "
   
    Dim X As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("?? ???? ????? ?????? ??? ?? ??", vbExclamation + vbYesNo)
    Else
        X = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
    End If

    If X = vbYes Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_Employee.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_Employee.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_EmployeeTotals.rpt"
        Else
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_EmployeeTotals.rpt"
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
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
 
 Function ShowTransactionsWith_Car(Optional Account_code As String = "", Optional CarID As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Check2.value = vbChecked Then
            MySQL = " SELECT TOP 100 PERCENT "
            
            MySQL = MySQL + " dbo.FixedAssets.branch_no"
            MySQL = MySQL + " ,dbo.FixedAssets.BoardNo"
            MySQL = MySQL + " ,dbo.FixedAssets.name"
            MySQL = MySQL + " ,dbo.FixedAssets.namee"
            MySQL = MySQL + " ,AccName1 = '?????'"
            MySQL = MySQL + " ,ValueAcc1 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410303016'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "             ,AccName2 = '?????'"
            MySQL = MySQL + " ,ValueAcc2 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "                          iNNER JOIN notes_all oN DEV.notes_all = notes_all.NoteID AND notes_all.NoteType = 370"
            MySQL = MySQL + "                          LEFT OUTER JOIN  TblCarsData oN TblCarsData.ID = notes_all.CarId"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410303006'"
            MySQL = MySQL + "               AND TblCarsData.fixedAssetid = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "               ,AccName3 = '??? ????'"
            MySQL = MySQL + "             ,ValueAcc3 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410303015'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "             ,AccName4 = '?????'"
            MySQL = MySQL + "             ,ValueAcc4 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410101001'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "             ,AccName5 = '???? ??????'"
            MySQL = MySQL + "             ,ValueAcc5 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410303007'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "             ,AccName6 = '??????'"
            MySQL = MySQL + "  ,ValueAcc6 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "    WHERE a.Account_Serial = '410303004'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "  ,AccName7 = '????'"
            MySQL = MySQL + "  ,ValueAcc7 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "                WHERE a.Account_Serial = '410303002'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "  ,AccName8 = '?????'"
            MySQL = MySQL + "  ,ValueAcc8 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "    WHERE a.Account_Serial = '410303009'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "  ,AccName9 = '???? ???? ???'"
            MySQL = MySQL + "  ,ValueAcc9 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "    WHERE a.Account_Serial = '410308001'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
             If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "  ,AccName10 = '???? ??????'"
            MySQL = MySQL + "  ,ValueAcc10 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "    WHERE a.Account_Serial = '410308007'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            
            MySQL = MySQL + "  ,AccName11 = '????'"
            MySQL = MySQL + "  ,ValueAcc11 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "    WHERE a.Account_Serial = '410303010'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            MySQL = MySQL + "  ,AccName12 = '??????? ????'"
            MySQL = MySQL + "  ,ValueAcc12 = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '410303022'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
              
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
                     
                     MySQL = MySQL + "  ,AccName12 = '?????????'"
            MySQL = MySQL + "  ,Rev = (SELECT"
            MySQL = MySQL + "                   SUM ([value])"
            MySQL = MySQL + "               FROM DOUBLE_ENTREY_VOUCHERS dev"
            MySQL = MySQL + "               INNER JOIN ACCOUNTS a"
            MySQL = MySQL + "                   ON a.Account_Code = dev.Account_Code"
            MySQL = MySQL + "               WHERE a.Account_Serial = '310101001'"
            MySQL = MySQL + "               AND dev.FixedAssetId = FixedAssets.Id"
              
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  dev.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
        
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and dev.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
            
                    
            MySQL = MySQL + " , Rev2 ="
            MySQL = MySQL + "(SELECT SUM(TblTravDueKDet.Value) FROM TblTravDueKDet INNER JOIN TblTravDueK ttdk "
            MySQL = MySQL + " ON TblTravDueKDet.TravID = ttdk.ID LEFT Outer JOIN   TblCarsData ON TblCarsData.ID =TblTravDueKDet.CarID"
            MySQL = MySQL + " WHERE TblCarsData.fixedAssetid = FixedAssets.ID"
            
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and  ttdk.recordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
            
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and ttdk.recordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            MySQL = MySQL + " )"
        

        
             
            MySQL = MySQL + " From dbo.FixedAssets where 1 = 1"

            If CarID <> 0 Then
                MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = " & CarID & ") "
            End If
    Else
    
        MySQL = "  SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Code, "
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate,"
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.NoteType,"
        MySQL = MySQL + " dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.TblCarsData.Branch_NO,"
        MySQL = MySQL + " dbo.TblCarsData.LicenseNO , dbo.TblCarsData.name"
        MySQL = MySQL + " FROM         dbo.ACCOUNTS INNER JOIN"
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        MySQL = MySQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID INNER JOIN"
        MySQL = MySQL + " dbo.TblCarsData ON dbo.DOUBLE_ENTREY_VOUCHERS.Carid = dbo.TblCarsData.id LEFT OUTER JOIN"
        MySQL = MySQL + " dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
        MySQL = MySQL + " where 1=1"
    
        MySQL = "  SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Code, "
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate,"
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.NoteType,"
        MySQL = MySQL + " dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.FixedAssets.Branch_NO,"
        MySQL = MySQL + " dbo.FixedAssets.BoardNo , dbo.FixedAssets.Name,   dbo.FixedAssets.NameE"
        MySQL = MySQL + " FROM         dbo.ACCOUNTS INNER JOIN"
        MySQL = MySQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        MySQL = MySQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID INNER JOIN"
        MySQL = MySQL + " dbo.FixedAssets ON dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = dbo.FixedAssets.id LEFT OUTER JOIN"
        MySQL = MySQL + " dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
        MySQL = MySQL + " where 1=1"
        
    'FixedAssetId
        If CarID <> 0 Then
            MySQL = MySQL + " and  (dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId = " & CarID & ") "
        End If
    
        If Account_code <> "" Then
            MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.account_code='" & Account_code & "'"
        End If
    
        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            MySQL = MySQL + " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
        End If
    
        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
        End If
    
        MySQL = MySQL + " Order By dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate "
    End If
    Dim X As Integer

Dim Y  As Integer
    If Check2.value = vbUnchecked Then
            If SystemOptions.UserInterface = ArabicInterface Then
                X = MsgBox("?? ???? ????? ?????? ??? ?? ??", vbExclamation + vbYesNo)
            Else
                X = MsgBox("Do you want Detailed Report y/n?", vbExclamation + vbYesNo)
            End If
        
        
            If SystemOptions.UserInterface = ArabicInterface Then
                Y = MsgBox("?? ???? ????? ?????????   ??? ?? ??", vbExclamation + vbYesNo)
            Else
                Y = MsgBox("Do you want group by account  name  y/n?", vbExclamation + vbYesNo)
            End If
            
            
            If X = vbYes Then
                If SystemOptions.UserInterface = ArabicInterface Then
                                If Y = vbYes Then
                                       StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_car.rpt"
                                 Else
                                       StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_carnogroup.rpt"
                                 End If
                Else
                                If Y = vbYes Then
                                       StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_car.rpt"
                                 Else
                                 
                                        StrFileName = App.path & "\Reports\" & "\Transporter\Transactions_with_carnogroup.rpt"
                                 End If
                End If
        
            Else
        
                If SystemOptions.UserInterface = ArabicInterface Then
                            If Y = vbYes Then
                                    StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotals.rpt"
                            Else
                            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotalsnogroup.rpt"
                            
                            End If
                Else
                            If Y = vbYes Then
                                StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotals.rpt"
                             Else
                             StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotalsnogroup.rpt"
                             End If
                End If
                
                If Check1.value = vbChecked Then
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotals1.rpt"
                Else
                    StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carTotals1.rpt"
                End If
                
                
                End If
                
            End If
    Else
        If Check1.value = vbChecked Then
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carExpensTotal.rpt"
        Else
            StrFileName = App.path & "\Reports\Transporter\" & "Transactions_with_carExpens.rpt"
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
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "dd/mm/yyyy") & CHR(13)
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "dd/mm/yyyy") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Function ShowGLWITH_Cost_center(Optional NoteType As Long = 0, Optional branch_id As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From GL_CC  where 1=1"
 
    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    If NoteType <> 0 Then
        MySQL = MySQL + "  and  ( NoteType = " & NoteType & ") "
    End If

If branch_id <> 0 Then
        MySQL = MySQL + "  and  ( branch_id = " & branch_id & ") "
    End If



    Dim X As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("?? ???? ????? ?? ??? ?? ????", vbExclamation + vbYesNo)
    Else
        X = MsgBox("Print Each Voucher in seprate Page", vbExclamation + vbYesNo)
    End If

    If X = vbNo Then
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
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    
If X = vbNo Then
    If PrintBranchINGE = True Then
        xReport.ParameterFields(4).AddCurrentValue "1"
    Else
        xReport.ParameterFields(4).AddCurrentValue "0"
    End If

    If PrintCCinGE = True Then
        xReport.ParameterFields(5).AddCurrentValue "1"
    Else
        xReport.ParameterFields(5).AddCurrentValue "0"
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

Function ShowGLto_projectNew(project_id As Integer, Optional Pand As Double, Optional Grouping As Boolean = False, Optional Account_code As String = "")
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_sub_projects where 1=1 " '
   ' Where project_id = " & project_id"
'
'Account_code = get_project_customer_account(project_id, "Account_Code")
     'Account_code = Me.TrvAccounts.SelectedItem.Key
                    
     '
     '               StrAccountName = Me.TrvAccounts.SelectedItem.Text
                    
'If SystemOptions.PaymentIntoAccouStat = False Then
            If StrAccountCode <> "" Then
            MySQL = MySQL & " and Account_Code='" & StrAccountCode & "'"
            End If
'End If


    'If project_id = 0 Then
    '    Exit Function
     '   MySQL = "Select * From RptLedger_sub_projects where 1=1 "
    'End If
If SystemOptions.Revenueowed = True Then
MySQL = MySQL & "and NoteType<>5000 "
Else
If SystemOptions.PaymentIntoAccouStat = False Then
 MySQL = MySQL & "and NoteType<>4 "
End If

End If
 

    If Pand <> 0 Then
        MySQL = MySQL + " and pandid=" & Pand & ""
        Dim sql As String
        Dim rsvalue As New ADODB.Recordset
        Dim opr_expected_value As Double
        sql = "select total from projects_des  where oprid=" & Pand & ""
        rsvalue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rsvalue.RecordCount > 0 Then
            opr_expected_value = IIf(IsNull(rsvalue("total").value), 0, rsvalue("total").value)
        Else
            rsvalue.Close
            sql = "select total from terms_operations  where ProjectDes_ID=" & Pand & ""
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
If Grouping = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL _with_projectsAcc.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL _with_projectsAcc.rpt"
    End If
    
Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL _with_projectsAcc.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL _with_projectsAcc.rpt"
    End If

End If


    If Dir(StrFileName) = "" Then

        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If Account_code <> "" Then
'MySQL = MySQL & " and Account_Code='" & Account_Code & "'"
End If



    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then

        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "?????? ?????? ?????"
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
   
        StrReportTitle = "????? ???? ???? ???  ???????? " '& StrAccountName

        If fullcode <> "" Then
            If SystemOptions.Items_or_operation = 0 Then
                StrReportTitle = "????? ????? " + dcprojects + " ??? ?????? " + Me.Dcdetails.text
            ElseIf SystemOptions.Items_or_operation = 1 Then
                StrReportTitle = "????? ????? " + dcprojects + " ??? ???????? " + Me.Dcdetails.text
            End If
        End If

        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Function ShowGLto_project(project_id As Integer, Optional Pand As Double, Optional Grouping As Boolean = False, Optional Account_code As String = "", Optional operid As Double)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


    If chk(0).value = vbChecked Or chk(1).value = vbChecked Then
        If chk(0).value = vbChecked Then
            MySQL = " SELECT Accounts.Account_Name as  Name,Accounts.Account_Code,RptLedger_sub_projects.Net,CusNamee,dbo.RptLedger_sub_projects.Project_nameE, dbo.RptLedger_sub_projects.DevDESE, dbo.RptLedger_sub_projects.branch_namee,"
            MySQL = MySQL & "        RptLedger_sub_projects.project_id,projects_des.des,projects_des.oprid,"
            MySQL = MySQL & "                RptLedger_sub_projects.Project_name,"
            MySQL = MySQL & "                                   projects_des.oprid,RptLedger_sub_projects.NoteSerial,RptLedger_sub_projects.NotesTypeName,"
            MySQL = MySQL & "                (RptLedger_sub_projects.DEV_Value) AS DEV_Value,"
            MySQL = MySQL & "                RptLedger_sub_projects.Credit_Or_Debit,"
            MySQL = MySQL & "                RptLedger_sub_projects.End_user_name,"
            MySQL = MySQL & "                Accounts.account_serial,RptLedger_sub_projects.* "
            MySQL = MySQL & "         From RptLedger_sub_projects"
            'MySQL = MySQL & "                LEFT OUTER JOIN ExpensesType"
            'MySQL = MySQL & "                     ON  ExpensesType.Account_Code = RptLedger_sub_projects.Account_Code"
            MySQL = MySQL & "                LEFT OUTER JOIN ACCOUNTS"
            MySQL = MySQL & "                     ON  RptLedger_sub_projects.Account_Code = ACCOUNTS.Account_Code"
            'MySQL = MySQL & "                     ON  ExpensesType.Account_Code = ACCOUNTS.Account_Code"
'            MySQL = MySQL & "                LEFT OUTER JOIN TblDataTypeExchange"
'            MySQL = MySQL & "                     ON  TblDataTypeExchange.Id = ExpensesType.DataTypeExchangeCode"
            
                    MySQL = MySQL & "                     Left outer join"
        MySQL = MySQL & "                          projects_des"
       MySQL = MySQL & "                          On RptLedger_sub_projects.project_id = projects_des.project_id"
       MySQL = MySQL & "                          and RptLedger_sub_projects.Pandid = projects_des.oprid"
            
            MySQL = MySQL & " Where RptLedger_sub_projects.project_id = " & project_id
            MySQL = MySQL & " and ACCOUNTS.AccountTypes = 2 and ACCOUNTS.AccountTab = 3"
            If Dcdetails.text <> "" And val(Dcdetails.BoundText) <> 0 Then
                MySQL = MySQL & " and projects_des.oprid= " & val(Dcdetails.BoundText)
            End If
        ElseIf chk(1).value = vbChecked Then
            MySQL = " SELECT RptLedger_sub_projects.*,  Accounts.Account_Name as  Name,Accounts.Account_Code,"
            MySQL = MySQL & "                Accounts.account_serial,projects_des.des,projects_des.oprid"
            MySQL = MySQL & "         From RptLedger_sub_projects"
            
            MySQL = MySQL & "                LEFT OUTER JOIN ACCOUNTS"
            MySQL = MySQL & "                     ON  RptLedger_sub_projects.Account_Code = ACCOUNTS.Account_Code"
        MySQL = MySQL & "                     Left outer join"
        MySQL = MySQL & "                          projects_des"
       MySQL = MySQL & "                          On RptLedger_sub_projects.project_id = projects_des.project_id"
            MySQL = MySQL & " Where RptLedger_sub_projects.project_id = " & project_id
            If Dcdetails.text <> "" And val(Dcdetails.BoundText) <> 0 Then
                MySQL = MySQL & " and projects_des.oprid= " & val(Dcdetails.BoundText)
            End If
            
            
        End If
        
        'DCAccounts
        If Trim(cmbAccount.BoundText) <> "" And Trim(cmbAccount.text) <> "" Then
            MySQL = MySQL & "   and  ACCOUNTS.Account_Code IN (SELECT Code"
            MySQL = MySQL & "                     FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(cmbAccount.BoundText) & "', '" & Trim(cmbAccount.BoundText) & "', 1))"
            MySQL = MySQL & "  OR (ACCOUNTS.Account_Code = '" & Trim(cmbAccount.BoundText) & "')"
        Else
           If Account_code <> "" Then
                MySQL = MySQL & "  and  ACCOUNTS.Account_Code = '" & Account_code & "'"
            End If
        End If
        If DCAccounts.text <> "" And DCAccounts.BoundText <> "" Then
            MySQL = MySQL & "   and ACCOUNTS.Account_Code = '" & Trim(DCAccounts.BoundText) & "'"
        End If
        Account_code = get_project_customer_account(project_id, "Account_Code")
        If SystemOptions.PaymentIntoAccouStat = False Then
            If Account_code <> "" Then
                MySQL = MySQL & " and RptLedger_sub_projects.Account_Code<>'" & Account_code & "'"
            End If
        End If

        If SystemOptions.Revenueowed = True Then
            MySQL = MySQL & "and NoteType<>5000 "
        Else
            If SystemOptions.PaymentIntoAccouStat = False Then
                MySQL = MySQL & "and NoteType<>4 "
            End If
        End If
            
               
            If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
                MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
            End If
            
            If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
                MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
            End If
            If chk(1).value <> vbChecked And chk(0).value <> vbChecked Then
                MySQL = MySQL + " Group By"
                MySQL = MySQL + "                ACCOUNTS.Account_Name,"
                MySQL = MySQL + "                RptLedger_sub_projects.project_id,RptLedger_sub_projects.Net,"
                MySQL = MySQL + "                RptLedger_sub_projects.Project_name,"
                MySQL = MySQL + "                RptLedger_sub_projects.Credit_Or_Debit,"
                MySQL = MySQL + "                RptLedger_sub_projects.End_user_name,"
                MySQL = MySQL + "                Accounts.account_serial,ACCOUNTS.Account_Code,"
                MySQL = MySQL + "                projects_des.des,projects_des.oprid"
            End If
            StrFileName = App.path & "\Reports\" & "GL _with_projectsWithExpen.rpt"
    Else
        MySQL = "Select RptLedger_sub_projects.*,TblDataTypeExchange.Name as DataTypeExchangeName,TblDataTypeExchange.Id as DataTypeExchangeCode From "
        If chk(0).value = vbChecked Or chk(1).value = vbChecked Then
            MySQL = MySQL & " RptLedger_sub_projects Inner join ExpensesType On ExpensesType.Account_Code =RptLedger_sub_projects.Account_Code  "
        Else
            MySQL = MySQL & " RptLedger_sub_projects Left Outer join ExpensesType On ExpensesType.Account_Code =RptLedger_sub_projects.Account_Code  "
        End If
        MySQL = MySQL & " Left Outer Join TblDataTypeExchange On TblDataTypeExchange.Id =ExpensesType.DataTypeExchangeCode "
        
        MySQL = MySQL & " Where project_id = " & project_id
    Account_code = get_project_customer_account(project_id, "Account_Code")
    If SystemOptions.PaymentIntoAccouStat = False Then
    If Account_code <> "" Then
    MySQL = MySQL & " and RptLedger_sub_projects.Account_Code<>'" & Account_code & "'"
    End If
    End If
    
    
         'DCAccounts
        If Trim(cmbAccount.BoundText) <> "" And Trim(cmbAccount.text) <> "" Then
            MySQL = MySQL & "   and  ACCOUNTS.Account_Code IN (SELECT Code"
            MySQL = MySQL & "                     FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(cmbAccount.BoundText) & "', '" & Trim(cmbAccount.BoundText) & "', 1))"
            MySQL = MySQL & "  OR (ACCOUNTS.Account_Code = '" & Trim(cmbAccount.BoundText) & "')"
        End If
    If Trim(cmbDataTypeExchange.text) <> "" And val(cmbDataTypeExchange.BoundText) <> 0 Then
        MySQL = MySQL & "  and ExpensesType.DataTypeExchangeCode=" & val(cmbDataTypeExchange.BoundText)
    
    End If
    
        If project_id = 0 Then
            Exit Function
            MySQL = "Select * From RptLedger_sub_projects where 1=1 "
        End If
    If SystemOptions.Revenueowed = True Then
        MySQL = MySQL & "and NoteType<>5000 "
    Else
        If SystemOptions.PaymentIntoAccouStat = False Then
         MySQL = MySQL & "and NoteType<>4 "
        End If
    
    End If
     
    
        If Pand <> 0 Then
            MySQL = MySQL + " and pandid=" & Pand & ""
            Dim sql As String
            Dim rsvalue As New ADODB.Recordset
            Dim opr_expected_value As Double
            sql = "select total from projects_des  where oprid=" & Pand & ""
            rsvalue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
            If rsvalue.RecordCount > 0 Then
                opr_expected_value = IIf(IsNull(rsvalue("total").value), 0, rsvalue("total").value)
            Else
                rsvalue.Close
                sql = "select total from terms_operations  where ProjectDes_ID=" & Pand & ""
                rsvalue.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
                If rsvalue.RecordCount > 0 Then
                    opr_expected_value = IIf(IsNull(rsvalue("total").value), 0, rsvalue("total").value)
                End If
            End If
     
        End If
    
        If operid <> 0 Then
            MySQL = MySQL + " and operid=" & operid & ""
            
         End If
        
        
        
        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
        End If
    
        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
        End If
        
        If Grouping = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL _with_projects.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL _with_projectse.rpt"
    End If
    
Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL _with_projects1.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL _with_projectse1.rpt"
    End If

End If

    End If


    If Dir(StrFileName) = "" Then

        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If Account_code <> "" Then
'MySQL = MySQL & " and Account_Code='" & Account_Code & "'"
End If



    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then

        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "No data to view"
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
   
        StrReportTitle = "Projects " '& StrAccountName

 

   StrReportTitle = "Projects " + dcprojects + CHR(13)
        If Dcdetails.text <> "" Then
           
                StrReportTitle = StrReportTitle + " Projects " + Me.Dcdetails.text + CHR(13)
        
        End If
        
        
        
        If DcbProcess1.text <> "" Then
           
                StrReportTitle = StrReportTitle + " Projects  " + Me.DcbProcess1.text + CHR(13)
        
        End If
        
        
        If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " From" & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " To " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function


Function ShowGl(Optional NoteType As Long = 0, Optional branch_id As Integer = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
DCNotesTypes.BoundText = 0
    MySQL = "Select * From RptLedger_Sub where 1=1  "
 
    If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        MySQL = MySQL + " and  RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    End If

    If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    End If

    If NoteType <> 0 Then
        MySQL = MySQL + " and    ( NoteType = " & NoteType & ") "
    End If
    
    If branch_id <> 0 Then
         MySQL = MySQL + " and    ( branch_id = " & branch_id & ") "
    End If
    

    Dim X As Integer

    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("?? ???? ????? ?? ??? ?? ????", vbExclamation + vbYesNo)
    Else
        X = MsgBox("Print Each Voucher In Seprate Page ", vbExclamation + vbYesNo)
    End If

    If X = vbNo Then

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
        Msg = "?????? ?????? ?????"
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
            StrReportTitle = StrReportTitle + " ????? ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub cmdDeleteRows_Click()
    Grid3.rows = 1

End Sub

Private Sub cmdClear_Click()
Grid3.rows = 1
End Sub

Private Sub CmdLoadTree_Click()
    ModTree.LoadTreeAccount Me.TrvAccounts
    Me.TrvAccounts.Nodes("r").EnsureVisible
    Me.TrvAccounts.Nodes("r").Expanded = True
    Me.TrvAccounts.Nodes("r").Selected = True

End Sub

Private Sub cmdSelectAll_Click()
 Dim Row As Integer
    If Grid3.rows > 1 Then
        For Row = 1 To Grid3.rows - 1
          Grid3.TextMatrix(Row, Grid3.ColIndex("Sel")) = -1
        Next
    End If
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim Row As Integer
    If Grid3.rows > 1 Then
        For Row = 1 To Grid3.rows - 1
          Grid3.TextMatrix(Row, Grid3.ColIndex("Sel")) = 0
        Next
    End If
End Sub

Private Sub Command1_Click()
    'MsgBox getprofitValue(Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, Val(Me.DCActivity.BoundText), Val(Me.dcBranch.BoundText))

    MsgBox GetOpeningBalanceDateForType2(DTPickerAccFrom.value)

End Sub

Private Sub Command2_Click()
LblAccountName.Caption = ""
TxtAccountCode.text = ""
StrAccountCode = ""
 StrAccountName = ""
 
End Sub

Private Sub Command7_Click()
Translatefrm Me
End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 22915
    End If
End Sub

Private Sub Dcdetails_Click(Area As Integer)
     Dim Dcombos As ClsDataCombos
 Dim project_id As Integer
       Set Dcombos = New ClsDataCombos
  If dcprojects.BoundText <> "" Then
     
         If Me.Dcdetails.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt DcbProcess1, val(dcprojects.BoundText), , Dcdetails.BoundText, 2
         End If
       
    End If
End Sub

Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FixedAssetsSearch.RetrunType = 6
        FixedAssetsSearch.show vbModal
  
    End If

End Sub

Private Sub DCProjects_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 9
             FrmProjectSearch.show
           
        End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub TxtAccountCode2_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim Msg As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

   
            If KeyCode = vbKeyReturn Then
                If Trim(Me.TxtAccountCode2.text) = "" Then Exit Sub
                If chk(0).value = vbChecked Then
                    StrSQL = "Select Account_Code From ACCOUNTS Where Account_Serial='" & Trim(Me.TxtAccountCode2.text) & "' and AccountTab = 3 and last_account = 0 AND [Level] >=3"
                ElseIf chk(1).value = vbChecked Then
                    StrSQL = " Select Account_Code From ACCOUNTS Where Account_Serial='" & Trim(Me.TxtAccountCode2.text) & "' and AccountTab = 2 and last_account = 0 AND [Level] >=3"
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    cmbAccount.BoundText = rs("Account_Code").value
                Else
                    Msg = "?????? ???? ???? ???? ?????..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DCEmployee.BoundText = EmpID
    End If

End Sub

Private Sub DcEmployee_Click(Area As Integer)

    If val(DCEmployee.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DCEmployee.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
 
End Sub

Private Sub DCEmployee_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 2
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
  
    End If

End Sub
Function fillterms1(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

  
        fill_combo Me.Dcdetails, My_SQL
        
    Dcdetails.ReFill
End Function
Private Sub dcprojects_Click(Area As Integer)
    Dim StrSQL  As String
 
        On Error Resume Next
   Dim fullcode As String
    If dcprojects.BoundText <> "" Then
   GetProjectsDetail val(dcprojects.BoundText), , fullcode
       Text2.text = fullcode
           fillterms1 (val(dcprojects.BoundText))
    End If
    
End Sub

'End Sub



Sub EmployeeBefnet()
If IsNull(DTPickerAccFrom.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Please Select Date"
Else
 MsgBox "Please Select Date"
End If
Exit Sub
End If
If IsNull(DTPickerAccTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Please Select Date"
Else
 MsgBox "Please Select Date"
End If
Exit Sub
End If
Dim sql As String
sql = "SELECT     Emp_ID, Emp_Name, Emp_Name1, Emp_Name2, Emp_Name3, Emp_Name4, Nationality, dean, DepartmentID, BranchId, Emp_Namee, Emp_Namee1, Emp_Namee2, "
sql = sql & "                      Emp_Namee3, Emp_Namee4, Fullcode, dbo.GetBalance('" & SQLDate(DTPickerAccFrom.value) & "', '" & SQLDate(DTPickerAccTo.value) & "', Account_code1,  1) AS Salar, Dateexppoket, Account_code1,"
sql = sql & "                       dbo.GetSalEmployee(Emp_ID, '" & SQLDate(DTPickerAccFrom.value) & "', '" & SQLDate(DTPickerAccTo.value) & "') AS Pay"
sql = sql & "  From dbo.TblEmployee"
sql = sql & "  Where 1=1"
If val(DCEmployee.BoundText) <> 0 And DCEmployee.text <> "" Then
sql = sql & " and Emp_id =" & val(DCEmployee.BoundText) & " "
End If
If val(DcboEmpDepartments.BoundText) <> 0 And DcboEmpDepartments.text <> "" Then
sql = sql & " and DepartmentID =" & val(DcboEmpDepartments.BoundText) & " "
End If
If val(dcBranch.BoundText) <> 0 And dcBranch.text <> "" Then
sql = sql & " and BranchID =" & val(dcBranch.BoundText) & " "
End If
print_report sql
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
  
MySQL = NoteSerial
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmployeeBefint.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmployeeBefintE.rpt"
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
         Msg = "·« ÌÊÃœ »Ì«‰« "
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    Dim str As String
    str = ""
    If SystemOptions.UserInterface = ArabicInterface Then
    If val(dcBranch.BoundText) <> 0 And dcBranch.text <> "" Then
     str = str & "  " & dcBranch.text
    End If
    If val(DcboEmpDepartments.BoundText) <> 0 And DcboEmpDepartments.text <> "" Then
    str = str & "  "
    str = str & "  " & DcboEmpDepartments.text
    End If
    Else
      If val(dcBranch.BoundText) <> 0 And dcBranch.text <> "" Then
   
    str = str & "  " & dcBranch.text
    End If
    If val(DcboEmpDepartments.BoundText) <> 0 And DcboEmpDepartments.text <> "" Then
    str = str & "  "
    
    str = str & "  " & DcboEmpDepartments.text
    End If
    End If
    xReport.ParameterFields(4).AddCurrentValue str
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

End Function
Private Sub Form_Load()
 '   Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
 'Me.Height = 8715
 'Me.Width = 13635
     Me.Height = 10000
    Me.Width = 17600
    Dim Dcombos As ClsDataCombos
    Me.left = (mdifrmmain.Width - Me.Width) / 2 - 1200
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
           Dim StrSQL As String
    
      ' OptAccount(15).Visible = False
       
    If SystemOptions.DateOpt = 1 Then
        FrameDateH.Visible = True
    
    End If
    
        If SystemOptions.AllowBigAccount = True Then
        'C1Elastic1.Visible = True
        OptAccount(35).Visible = True
        OptAccount(5).Visible = True
         OptAccount(18).Visible = True
          OptAccount(25).Visible = True
           
    If SystemOptions.ShowOldAccountReports = True Then
   OptAccount(5).Visible = True
   OptAccount(18).Visible = True
   OptAccount(25).Visible = True
   Else
    OptAccount(5).Visible = False
   OptAccount(18).Visible = False
   OptAccount(25).Visible = False
    End If
           
           
           
           OptAccount(3).Visible = True
            OptAccount(28).Visible = True
             OptAccount(4).Visible = True
              OptAccount(29).Visible = True
              OptAccount(15).Visible = True
        Else
        '        C1Elastic1.Visible = False
OptAccount(35).Visible = False
     
        OptAccount(5).Visible = False
         OptAccount(18).Visible = False
          OptAccount(25).Visible = False
           OptAccount(3).Visible = False
            OptAccount(28).Visible = False
             OptAccount(4).Visible = False
              OptAccount(29).Visible = False
              
              
    
    End If
    
 If SystemOptions.CanProjectAccountOnly = True Then
    For i = 0 To OptAccount.count - 1
'        OptAccount(i).Visible = False
'        If i = 10 Or i = 26 Or i = 42 Then
'            OptAccount(i).Visible = True
'        Else
'            OptAccount(i).Visible = True
'        End If
        Dim isControlExists As Boolean
        On Error Resume Next
        isControlExists = Not (OptAccount(i) Is Nothing)
        On Error GoTo xx ' «” ∆‰«› „⁄«·Ã… «·√Œÿ«¡ »⁄œ «· Õﬁﬁ

        If isControlExists Then
            OptAccount(i).Visible = False
            If i = 10 Or i = 26 Or i = 42 Then
                OptAccount(i).Visible = True
            Else
                OptAccount(i).Visible = False
            End If
        End If
xx:
    Next
    OptAccount(10).value = True
 End If
    DCNotesTypes.BoundText = 0

        StrSQL = "SELECT * From TblDataTypeExchange "
        fill_combo cmbDataTypeExchange, StrSQL
    

    ScreenNameArabic = "?????? ????????"
    ScreenNameEnglish = "Accounting Report"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    If mdifrmmain.TransporterMain.Visible = False Then
      '  OptAccount(15).Visible = False
        OptAccount(16).Visible = False
        OptAccount(17).Visible = False
    End If

    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    
    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)   order by account_name"
    fill_combo Me.DcCostCenter, StrSQL

    'StrSQL = "  SELECT id ,Project_name FROM projects"
    
    
                'StrSQL = " SELECT     ID, Project_Name"
                 If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_name )) as Project_name"
         Else
         StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_namee )) as Project_name"
         End If
         
            StrSQL = StrSQL & "            From dbo.projects"
            
            
            
        If SystemOptions.UserInterface = ArabicInterface Then
                 
                StrSQL = StrSQL & " where Project_name<>N'""' and not (Project_name is null)"
   Else
     
                StrSQL = StrSQL & " where Project_nameE<>N'""' and not (Project_nameE is null)"
End If


    fill_combo Me.dcprojects, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select branch_id,branch_name from TblBranchesData   "
    Else
        StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
    End If





    fill_combo dcBranch, StrSQL


  If SystemOptions.usertype <> UserAdminAll Then
  ' dcBranch.Enabled = False
   dcBranch.BoundText = Current_branch

  Else
  '   dcBranch.Enabled = True
   End If
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT NotesType,LTRIM(NotesTypeName) NotesTypeName From TblNotesTypes order by NotesTypeName "
    Else
        StrSQL = "SELECT NotesType,LTRIM(NotesTypeNamee) NotesTypeNamee From TblNotesTypes  order by NotesTypeNamee"
    End If

    fill_combo DCNotesTypes, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT CombositAccountid,Remark From TblCombositAccount order by Remark "
    Else
        StrSQL = "SELECT CombositAccountid,Remark From TblCombositAccount order by Remark "
    End If

    fill_combo DCCompositeAccount, StrSQL
    
    
     Set Dcombos = New ClsDataCombos
   

  '  Dcombos.GetAccountingCodes Me.cmbAccount
    


        StrSQL = "  SELECT Account_Code,Account_Name FROM ACCOUNTS  WHERE AccountTab = 3 and last_account = 0 AND [Level] >=3"
        StrSQL = StrSQL & "  Union All"
        StrSQL = StrSQL & "  SELECT Account_Code,Account_Serial, Account_Name         FROM ACCOUNTS  WHERE AccountTab = 2 and last_account = 0 AND [Level] >=3"
'a.Account_Serial = '4102' AND
        
        fill_combo cmbAccount, StrSQL
    


    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " select Emp_ID,Emp_Name from TblEmployee  order by Emp_Name"
    Else
        StrSQL = " select Emp_ID,Emp_Namee from TblEmployee  order by Emp_Namee "
    End If

    fill_combo DCEmployee, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select id,name from tblActivitesType   "
    Else
        StrSQL = "  select id,namee from tblActivitesType   "
    End If

    fill_combo DCActivity, StrSQL

    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCars Me.DCCar
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetFixedAssets Me.DcFixedAssets
    Dcombos.GetAccountingCodes Me.DCAccounts, True
    Dcombos.GetSection Me.DCRegionID
    Dcombos.GetIqar Me.DcbAqar
    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo
    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    Me.DTPickerAccFrom = FirstPeriodDateInthisYear
    Me.DTPickerAccTo = Date
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 If SystemOptions.CanProjectAccountOnly = True Then
    OptAccount(10).value = True
 Else
    OptAccount(0).value = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub ISButton1_Click()
    txt_mod_flag.text = "S"
 
    Account_search.show
    Account_search.case_id = 1
End Sub


Public Function Set_account_code(code As String, _
                                 Name As String, _
                                 Optional account_serial As String)
    StrAccountCode = code
    StrAccountName = Name
    Me.LblAccountName.Caption = Name
    TxtAccountCode.text = account_serial

End Function
Function print_report3(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer
  
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
 
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim BrcnActivety As String
 

   FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  
         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
   
            If HideZeroBalance = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
          Dim BranshesReg As String
      
         If val(DCRegionID.BoundText) <> 0 Then
         BranshesReg = BranchRegion(DCRegionID.BoundText)
         End If
         If val(DCActivity.BoundText) <> 0 Then
         BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
         End If


  updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg

  sql = " SELECT    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
  
 If val(DcbAqar.BoundText) = 0 Then
 sql = " SELECT    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
 
 Else
 
 
'  sql = " SELECT    unitno=("
'  sql = sql & "                     SELECT     TOP 1  PERCENT dbo.TblAqarDetai.unitno"
'  sql = sql & "                     FROM         dbo.TblContract INNER JOIN"
'  sql = sql & "                                          dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
'  sql = sql & "                                          dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
'  sql = sql & "                    Where (dbo.TblCustemers.Account_code = a.Account_code)"
'  sql = sql & "                    ORDER BY dbo.TblContract.ContNo DESC"
'  sql = sql & "                    )"
'  sql = sql & "                    ,"
'  sql = sql & "                        ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
 
  sql = " SELECT    unitno=(   SELECT     TOP 1  PERCENT dbo.TblAqarDetai.unitno                     FROM         dbo.TblContract INNER JOIN                                          dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN                                          dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID                    Where (dbo.TblCustemers.Account_code = a.Account_code)                    ORDER BY dbo.TblContract.ContNo DESC                    )   "
  sql = sql & "      , unitType=("
  sql = sql & "     SELECT     TOP 1   dbo.TblAkarUnit.name"
  sql = sql & "    FROM         dbo.TblContract INNER JOIN"
  sql = sql & "                          dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
  sql = sql & "                          dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID INNER JOIN"
  sql = sql & "                          dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id"
  sql = sql & "     Where (dbo.TblCustemers.Account_code = a.Account_code)"
  sql = sql & "     ORDER BY dbo.TblContract.ContNo DESC"
  sql = sql & "    ),"
  sql = sql & "     customerNo=("
  sql = sql & "     SELECT       dbo.TblCustemers. Cus_mobile"
  sql = sql & "    From dbo.TblCustemers"
    sql = sql & "    Where (dbo.TblCustemers.Account_code = a.Account_code)"
   sql = sql & "    )"
  sql = sql & "    ,     LegalIssue=("
  sql = sql & "     SELECT     TOP 1 dbo.TblContract.LegalIssue"
  sql = sql & "    FROM         dbo.TblContract INNER JOIN"
  sql = sql & "                          dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
  sql = sql & "      Where (dbo.TblCustemers.Account_code = a.Account_code)"
   sql = sql & "     ORDER BY dbo.TblContract.ContNo DESC"
  sql = sql & "    )"
  sql = sql & "    ,"
  sql = sql & "         aqarname=("
  sql = sql & "     SELECT     TOP 1  dbo.TblAqar.aqarname"
   sql = sql & "    FROM         dbo.TblContract INNER JOIN"
  sql = sql & "                          dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id INNER JOIN"
  sql = sql & "                          dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID INNER JOIN"
  sql = sql & "                          dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid"
   sql = sql & "     Where (dbo.TblCustemers.Account_code = a.Account_code)"
  sql = sql & "     ORDER BY dbo.TblContract.ContNo DESC"
  sql = sql & "    )"
  sql = sql & "                    ,    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
End If

  
  
  sql = sql & "                         (SELECT     SUM(DEV_Value1)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") AND d.Account_Code = A.Account_Code  and(d.Posted IS NULL)"
 If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
  End If
  'sql = sql & AqarFilter_d

  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BranshesReg & ")"
  End If
  
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d.branch_id =" & val(dcBranch.BoundText) & ""
  End If
 sql = sql & "  ) x),"
  sql = sql & "                    CreditBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & "  AND d1.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") AND d1.Account_Code = A.Account_Code and(d1.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
  End If
 ' sql = sql & AqarFilter_d1

  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
  End If
 If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                     OpeningBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  sql = sql & "                                                   WHERE     (  do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
 ' sql = sql & AqarFilter_do

  If val(dcBranch.BoundText) <> 0 Then
 sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
 End If
sql = sql & "  )) x),"
  sql = sql & "    OpeningBalancebeformdateMinus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
'  sql = sql & AqarFilter_do
'
  sql = sql & " ) x),"
  sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL) "
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  AqarFilter_v1 = ""
  sql = sql & AqarFilter_v1

  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x)"
  sql = sql & " FROM         ACCOUNTS A"
  sql = sql & " WHERE     A.last_account = 1   "
  

  
  sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
  sql = sql & "    Where 1 = 1"
  If val(DcbAqar.BoundText) = 0 Then
  StrAccountCode = (Me.TrvAccounts.SelectedItem.key)
        If mId(StrAccountCode, Len(StrAccountCode), 1) = "G" Then
                    StrAccountCode = mId(StrAccountCode, 1, Len(StrAccountCode) - 1)
                    
                    End If
     End If
    If StrAccountCode <> "" Then
            
                    
 sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
  End If
  
  
       If val(DcbAqar.BoundText) <> 0 Then
' sql = sql & "  and   A.Account_Code in ("
'
' sql = sql & "  SELECT     dbo.TblCustemers.Account_Code"
' sql = sql & "  FROM         dbo.TblContract INNER JOIN"
' sql = sql & "                       dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
' sql = sql & "  Where (dbo.TblContract.Iqar = " & DcbAqar.BoundText & ") ) "

        If val(DcbAqar.BoundText) <> 0 Then
            sql = sql & " AND (ISNULL(DOUBLE_ENTREY_VOUCHERS.Aqarid,0) = " & val(DcbAqar.BoundText) & _
                        " OR ISNULL(DOUBLE_ENTREY_VOUCHERS.iqarid,0) = " & val(DcbAqar.BoundText) & ") "
        End If

  End If
  
  
    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
   sql = sql & "   )"
  sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    
    
      If StrAccountCode <> "" Then
 sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
  End If
  

    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
  
  
'       If val(DcbAqar.BoundText) <> 0 Then
' sql = sql & "  and   A.Account_Code in ("
'
' sql = sql & "  SELECT     dbo.TblCustemers.Account_Code"
' sql = sql & "  FROM         dbo.TblContract INNER JOIN"
' sql = sql & "                       dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID"
' sql = sql & "  Where (dbo.TblContract.Iqar = " & DcbAqar.BoundText & ") ) "
'
'  End If
'
'  If val(DcbAqar.BoundText) <> 0 Then
'    sql = sql & " AND (ISNULL(DOUBLE_ENTREY_VOUCHERS.Aqarid,0) = " & val(DcbAqar.BoundText) & _
'                " OR ISNULL(DOUBLE_ENTREY_VOUCHERS.iqarid,0) = " & val(DcbAqar.BoundText) & ") "
'End If

  
   sql = sql & "   ))"
   
  
    sql = sql & "order by Account_Serial "
    
    
    Dim AqarFilter_d As String, AqarFilter_d1 As String, AqarFilter_do As String, AqarFilterExistsV As String

AqarFilter_d = "": AqarFilter_d1 = "": AqarFilter_do = "": AqarFilterExistsV = ""

If val(DcbAqar.BoundText) <> 0 Then
    AqarFilter_d = " AND (ISNULL(d.Aqarid,0)=" & val(DcbAqar.BoundText) & _
                  " OR ISNULL(d.iqarid,0)=" & val(DcbAqar.BoundText) & ") "
    AqarFilter_d1 = " AND (ISNULL(d1.Aqarid,0)=" & val(DcbAqar.BoundText) & _
                   " OR ISNULL(d1.iqarid,0)=" & val(DcbAqar.BoundText) & ") "
    AqarFilter_do = " AND (ISNULL(do.Aqarid,0)=" & val(DcbAqar.BoundText) & _
                   " OR ISNULL(do.iqarid,0)=" & val(DcbAqar.BoundText) & ") "
    AqarFilterExistsV = " AND (ISNULL(DOUBLE_ENTREY_VOUCHERS.Aqarid,0)=" & val(DcbAqar.BoundText) & _
                      " OR ISNULL(DOUBLE_ENTREY_VOUCHERS.iqarid,0)=" & val(DcbAqar.BoundText) & ") "
End If

Dim AqarFilter_Cont As String
AqarFilter_Cont = ""

If val(DcbAqar.BoundText) <> 0 Then
    AqarFilter_Cont = " AND dbo.TblAqarDetai.Aqarid=" & val(DcbAqar.BoundText)
End If

Dim sHead As String, sDebit As String, sCredit As String, sOpen As String, sOB1 As String, sOB2 As String
Dim sFrom As String, sExists As String

If val(DcbAqar.BoundText) = 0 Then
    sHead = " SELECT ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng, "
Else

    sHead = " SELECT " & _
            " unitno=(" & _
            "   SELECT TOP 1 PERCENT dbo.TblAqarDetai.unitno " & _
            "   FROM dbo.TblContract " & _
            "   INNER JOIN dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id " & _
            "   INNER JOIN dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID " & _
            "   WHERE dbo.TblCustemers.Account_code = A.Account_code " & _
            AqarFilter_Cont & _
            "   ORDER BY dbo.TblContract.ContNo DESC" & _
            " ),"

    sHead = sHead & _
            " unitType=(" & _
            "   SELECT TOP 1 dbo.TblAkarUnit.name " & _
            "   FROM dbo.TblContract " & _
            "   INNER JOIN dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id " & _
            "   INNER JOIN dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID " & _
            "   INNER JOIN dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id " & _
            "   WHERE dbo.TblCustemers.Account_code = A.Account_code " & _
            AqarFilter_Cont & _
            "   ORDER BY dbo.TblContract.ContNo DESC" & _
            " ),"

    sHead = sHead & _
            " customerNo=(" & _
            "   SELECT dbo.TblCustemers.Cus_mobile " & _
            "   FROM dbo.TblCustemers " & _
            "   WHERE dbo.TblCustemers.Account_code = A.Account_code" & _
            " ),"

    sHead = sHead & _
            " LegalIssue=(" & _
            "   SELECT TOP 1 dbo.TblContract.LegalIssue " & _
            "   FROM dbo.TblContract " & _
            "   INNER JOIN dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID " & _
            "   INNER JOIN dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id " & _
            "   WHERE dbo.TblCustemers.Account_code = A.Account_code " & _
            AqarFilter_Cont & _
            "   ORDER BY dbo.TblContract.ContNo DESC" & _
            " ),"

    sHead = sHead & _
            " aqarname=(" & _
            "   SELECT TOP 1 dbo.TblAqar.aqarname " & _
            "   FROM dbo.TblContract " & _
            "   INNER JOIN dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id " & _
            "   INNER JOIN dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID " & _
            "   INNER JOIN dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid " & _
            "   WHERE dbo.TblCustemers.Account_code = A.Account_code " & _
            AqarFilter_Cont & _
            "   ORDER BY dbo.TblContract.ContNo DESC" & _
            " ),"

    sHead = sHead & _
            " ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng, "

End If

sDebit = " debitBalance=(" & _
         " SELECT SUM(DEV_Value1) FROM (" & _
         "  SELECT Account_Code," & _
         "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
         "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
         "  FROM dbo.DOUBLE_ENTREY_VOUCHERS d" & _
         "  WHERE d.Credit_Or_Debit=0" & _
         "   AND d.RecordDate>=" & SQLDate(Me.DTPickerAccFrom.value, True) & _
         "   AND d.RecordDate<=" & SQLDate(Me.DTPickerAccTo.value, True) & _
         "   AND d.Account_Code=A.Account_Code" & _
         "   AND d.Posted IS NULL "

If val(DCActivity.BoundText) <> 0 Then sDebit = sDebit & " AND d.branch_id IN (" & BrcnActivety & ")"
If val(DCRegionID.BoundText) <> 0 Then sDebit = sDebit & " AND d.branch_id IN (" & BranshesReg & ")"
If val(dcBranch.BoundText) <> 0 Then sDebit = sDebit & " AND d.branch_id=" & val(dcBranch.BoundText)

sDebit = sDebit & AqarFilter_d & " ) x ), "
sCredit = " CreditBalance=(" & _
          " SELECT SUM(DEV_Value2) FROM (" & _
          "  SELECT Account_Code," & _
          "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
          "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
          "  FROM dbo.DOUBLE_ENTREY_VOUCHERS d1" & _
          "  WHERE d1.Credit_Or_Debit=1" & _
          "   AND d1.RecordDate>=" & SQLDate(Me.DTPickerAccFrom.value, True) & _
          "   AND d1.RecordDate<=" & SQLDate(Me.DTPickerAccTo.value, True) & _
          "   AND d1.Account_Code=A.Account_Code" & _
          "   AND d1.Posted IS NULL "

If val(DCActivity.BoundText) <> 0 Then sCredit = sCredit & " AND d1.branch_id IN (" & BrcnActivety & ")"
If val(DCRegionID.BoundText) <> 0 Then sCredit = sCredit & " AND d1.branch_id IN (" & BranshesReg & ")"
If val(dcBranch.BoundText) <> 0 Then sCredit = sCredit & " AND d1.branch_id=" & val(dcBranch.BoundText)

sCredit = sCredit & AqarFilter_d1 & " ) x ), "
'sOpen = " OpeningBalance=(" & _
'        " SELECT SUM(DEV_Value1)+SUM(DEV_Value2) FROM (" & _
'        "  SELECT Account_Code," & _
'        "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
'        "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
'        "  FROM dbo.DOUBLE_ENTREY_VOUCHERS1 v1" & _
'        "  WHERE v1.Account_Code=A.Account_Code" & _
'        "   AND v1.Posted IS NULL "
'
'If val(DCActivity.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id IN (" & BrcnActivety & ")"
'If val(DCRegionID.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id IN (" & BranshesReg & ")"
'If val(dcBranch.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id=" & val(dcBranch.BoundText)
'
'sOpen = sOpen & " ) x ), "
If val(DcbAqar.BoundText) = 0 Then
    ' OpeningBalance „‰ vouchers1 (⁄«œÌ)
    sOpen = " OpeningBalance=(" & _
            " SELECT SUM(DEV_Value1)+SUM(DEV_Value2) FROM (" & _
            "  SELECT Account_Code," & _
            "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
            "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
            "  FROM dbo.DOUBLE_ENTREY_VOUCHERS1 v1" & _
            "  WHERE v1.Account_Code=A.Account_Code" & _
            "   AND v1.Posted IS NULL "

    If val(DCActivity.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id IN (" & BrcnActivety & ")"
    If val(DCRegionID.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id IN (" & BranshesReg & ")"
    If val(dcBranch.BoundText) <> 0 Then sOpen = sOpen & " AND v1.branch_id=" & val(dcBranch.BoundText)

    sOpen = sOpen & " ) x ), "
Else
    ' ·„« ÌﬂÊ‰ ›ÌÂ ⁄ﬁ«—: „›Ì‘ OpeningBalance „‰ vouchers1
    sOpen = " OpeningBalance=0, "
End If

sOB1 = " OpeningBalancebeformdateMinus1=(" & _
       " SELECT SUM(DEV_Value1)+SUM(DEV_Value2) FROM (" & _
       "  SELECT Account_Code," & _
       "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
       "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
       "  FROM dbo.DOUBLE_ENTREY_VOUCHERS do" & _
       "  WHERE do.RecordDate>=" & SQLDate(openingBalanceDate, True) & _
       "   AND do.RecordDate<=" & SQLDate(FromdateMinus1, True) & _
       "   AND do.Account_Code=A.Account_Code" & _
       "   AND do.Posted IS NULL "

If val(DCActivity.BoundText) <> 0 Then sOB1 = sOB1 & " AND do.branch_id IN (" & BrcnActivety & ")"
If val(DCRegionID.BoundText) <> 0 Then sOB1 = sOB1 & " AND do.branch_id IN (" & BranshesReg & ")"
If val(dcBranch.BoundText) <> 0 Then sOB1 = sOB1 & " AND do.branch_id=" & val(dcBranch.BoundText)

sOB1 = sOB1 & AqarFilter_do & " ) x ), "
sOB2 = " OpeningBalancebeformStartCurrentyearTOFromDAteminus1=(" & _
       " SELECT SUM(DEV_Value1)+SUM(DEV_Value2) FROM (" & _
       "  SELECT Account_Code," & _
       "   DEV_Value1=CASE WHEN Credit_Or_Debit=0 THEN Value ELSE 0 END," & _
       "   DEV_Value2=CASE WHEN Credit_Or_Debit=1 THEN Value*-1 ELSE 0 END" & _
       "  FROM dbo.DOUBLE_ENTREY_VOUCHERS do" & _
       "  WHERE do.RecordDate>=" & SQLDate(StartCurrentDate, True) & _
       "   AND do.RecordDate<" & SQLDate(Me.DTPickerAccFrom.value, True) & _
       "   AND do.Account_Code=A.Account_Code" & _
       "   AND do.Posted IS NULL "

If val(DCActivity.BoundText) <> 0 Then sOB2 = sOB2 & " AND do.branch_id IN (" & BrcnActivety & ")"
If val(DCRegionID.BoundText) <> 0 Then sOB2 = sOB2 & " AND do.branch_id IN (" & BranshesReg & ")"
If val(dcBranch.BoundText) <> 0 Then sOB2 = sOB2 & " AND do.branch_id=" & val(dcBranch.BoundText)

sOB2 = sOB2 & AqarFilter_do & " ) x ) "
'sFrom = " FROM ACCOUNTS A WHERE A.last_account=1 "
sFrom = " FROM ACCOUNTS A WHERE A.last_account=1 "

If val(DcbAqar.BoundText) <> 0 Then
    sFrom = sFrom & _
            " AND EXISTS ( " & _
            "   SELECT 1 " & _
            "   FROM dbo.TblContract c " & _
            "   INNER JOIN dbo.TblCustemers cu ON c.CusID = cu.CusID " & _
            "   INNER JOIN dbo.TblAqarDetai ad ON c.UnitNo = ad.Id " & _
            "   WHERE cu.Account_code = A.Account_code " & _
            "     AND ad.Aqarid = " & val(DcbAqar.BoundText) & _
            " ) "
End If



sExists = " AND ( " & _
          " A.Account_Code IN (SELECT Account_Code FROM DOUBLE_ENTREY_VOUCHERS WHERE 1=1 "

If StrAccountCode <> "" Then sExists = sExists & " AND Account_Code LIKE '" & StrAccountCode & "a%'"
If val(DCActivity.BoundText) <> 0 Then sExists = sExists & " AND branch_id IN (" & BrcnActivety & ")"
If val(DCRegionID.BoundText) <> 0 Then sExists = sExists & " AND branch_id IN (" & BranshesReg & ")"
If val(dcBranch.BoundText) <> 0 Then sExists = sExists & " AND branch_id=" & val(dcBranch.BoundText)

' ›· — «·⁄ﬁ«— ⁄·Ï vouchers ›ﬁÿ
sExists = sExists & AqarFilterExistsV & " ) "

' ·Ê „›Ì‘ ⁄ﬁ«— »”° ”«⁄ Â« ‰÷Ì› OR vouchers1
If val(DcbAqar.BoundText) = 0 Then
    sExists = sExists & " OR A.Account_Code IN (SELECT Account_Code FROM DOUBLE_ENTREY_VOUCHERS1 WHERE 1=1 "

    If StrAccountCode <> "" Then sExists = sExists & " AND Account_Code LIKE '" & StrAccountCode & "a%'"
    If val(DCActivity.BoundText) <> 0 Then sExists = sExists & " AND branch_id IN (" & BrcnActivety & ")"
    If val(DCRegionID.BoundText) <> 0 Then sExists = sExists & " AND branch_id IN (" & BranshesReg & ")"
    If val(dcBranch.BoundText) <> 0 Then sExists = sExists & " AND branch_id=" & val(dcBranch.BoundText)

    sExists = sExists & " ) "
End If

sExists = sExists & " ) "


sql = sHead & sDebit & sCredit & sOpen & sOB1 & sOB2 & sFrom & sExists & " ORDER BY Account_Serial "





       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSa.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaE.rpt"
        End If
        
        If val(DcbAqar.BoundText) > 0 Then
        Dim X As Integer
        X = MsgBox(" ", vbInformation + vbYesNo)
                         If X = vbYes Then
                        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaAkar.rpt"
'                         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSa.rpt"
                    Else
                        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaAkar2.rpt"
'                        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaE.rpt"
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
       Msg = "·« ÌÊÃœ »Ì«‰« "
     Else
     Msg = "No Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim desc As String
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    desc = ""
       If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Account" & ": " & LblAccountName.Caption & CHR(13)
   Else
   desc = desc & "Account" & ": " & LblAccountName.Caption & CHR(13)
   End If
   
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
   
   
     If val(DcbAqar.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & " ⁄ﬁ«—" & ": " & DcbAqar.text & CHR(13)
   Else
   desc = desc & " Akar" & ": " & DcbAqar.text & CHR(13)
   End If
   End If
   
  
   
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    xReport.ParameterFields(7).AddCurrentValue desc
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
 
Function print_report41(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer
  
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
 
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
  Dim HideLastAccount  As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim ShowOnlyLevelAcc As Integer
   Dim BrcnActivety As String
   FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  
  
           
   If SystemOptions.UserInterface = ArabicInterface Then
                 X = val(InputBox("Specify Level"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            
            account_level = val(X)
             
 
            
            
         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
   
            If HideZeroBalance = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
                     If SystemOptions.UserInterface = ArabicInterface Then
               HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            Else
                HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            End If
                  If HideLastAccount = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
   
   
            
            
                        

                  If ShowOnlyLevelAcc = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
            
            
          Dim BranshesReg As String
      
         If val(DCRegionID.BoundText) <> 0 Then
         BranshesReg = BranchRegion(DCRegionID.BoundText)
         End If
         If val(DCActivity.BoundText) <> 0 Then
         BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
         End If


  updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True

   '=========================================================
  ' FAST Income Statement SQL (Set-Based) - same output cols
  '=========================================================
 '=========================================================
' FAST Income Statement SQL (Set-Based) - VB6 Ready
' Output columns:
' debitBalance, CreditBalance, OpeningBalance,
' OpeningBalancebeformdateMinus1, OpeningBalancebeformStartCurrentyearTOFromDAteminus1
'=========================================================

Dim FromDate As Date, ToDate As Date
Dim OpenFrom As Date, OpenTo As Date
Dim StartYear As Date
Dim OnlyThisLevel As Integer
Dim LevelVal As Integer

FromDate = Me.DTPickerAccFrom.value
ToDate = Me.DTPickerAccTo.value

' openingBalanceDate Ê StartCurrentDate «‰  √’·« » Õ”»Â„ ›Êﬁ
OpenFrom = openingBalanceDate
OpenTo = DateAdd("d", -1, FromDate)
StartYear = StartCurrentDate   ' √Ê DATEFROMPARTS(YEAR(FromDate),1,1) ·Ê ⁄«Ì“Â ﬂœÂ

' „” ÊÏ «·Õ”«»
LevelVal = 5
If account_level <> 0 Then LevelVal = account_level

' ShowOnlyLevelAcc = vbYes => OnlyThisLevel = 1 (Ì⁄‰Ì = Level)
' €Ì— ﬂœÂ => <= Level
If ShowOnlyLevelAcc = 0 Then
    OnlyThisLevel = 1
Else
    OnlyThisLevel = 0
End If
 OnlyThisLevel = 1
sql = ""
sql = sql & "DECLARE @FromDate date = " & SQLDate(FromDate, True) & ";" & vbCrLf
sql = sql & "DECLARE @ToDate   date = " & SQLDate(ToDate, True) & ";" & vbCrLf
sql = sql & "DECLARE @OpenFrom date = " & SQLDate(OpenFrom, True) & ";" & vbCrLf
sql = sql & "DECLARE @OpenTo   date = " & SQLDate(OpenTo, True) & ";" & vbCrLf
sql = sql & "DECLARE @StartYear date = " & SQLDate(StartYear, True) & ";" & vbCrLf
sql = sql & "DECLARE @Level int = " & LevelVal & ";" & vbCrLf
sql = sql & "DECLARE @OnlyThisLevel bit = " & OnlyThisLevel & ";" & vbCrLf
sql = sql & vbCrLf

sql = sql & ";WITH A AS (" & vbCrLf
sql = sql & "    SELECT" & vbCrLf
sql = sql & "        A.last_account," & vbCrLf
sql = sql & "        A.ProfitBalance," & vbCrLf
sql = sql & "        A.Parent_Account_Code," & vbCrLf
sql = sql & "        A.AccountTypes," & vbCrLf
sql = sql & "        A.Account_Code," & vbCrLf
sql = sql & "        A.Account_Serial," & vbCrLf
sql = sql & "        A.Account_Name," & vbCrLf
sql = sql & "        A.Account_NameEng" & vbCrLf
sql = sql & "    FROM dbo.ACCOUNTS A" & vbCrLf
sql = sql & "    WHERE" & vbCrLf
sql = sql & "        A.AccountTypes = 2" & vbCrLf

' HideLastAccount („Â„: MsgBox »Ì—Ã⁄ vbYes/vbNo)
If HideLastAccount = vbYes Then
    sql = sql & "        AND A.last_account = 0" & vbCrLf
End If

If (TxtAccountCode.text) <> "" Then
    sql = sql & "        AND A.Account_Serial = '" & Replace(TxtAccountCode.text, "'", "''") & "'" & vbCrLf
End If

' ›· — «·„” ÊÏ
sql = sql & "        AND ( " & vbCrLf
sql = sql & "            (@OnlyThisLevel = 1 AND (LEN(A.account_code) - LEN(REPLACE(A.account_code,'a',''))) = @Level)" & vbCrLf
sql = sql & "            OR" & vbCrLf
sql = sql & "            (@OnlyThisLevel = 0 AND (LEN(A.account_code) - LEN(REPLACE(A.account_code,'a',''))) <= @Level)" & vbCrLf
sql = sql & "        )" & vbCrLf

' ‘—ÿ ÊÃÊœ Õ—ﬂ« /√—’œ…/√Ê Parent
sql = sql & "        AND (" & vbCrLf
sql = sql & "               EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS d WHERE d.Account_Code = A.Account_Code" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "                      AND d.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "                      AND d.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "                      AND d.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "               )" & vbCrLf

sql = sql & "            OR EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS1 d1 WHERE d1.Account_Code = A.Account_Code" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "                      AND d1.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "                      AND d1.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "                      AND d1.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "               )" & vbCrLf

sql = sql & "            OR EXISTS (SELECT 1 FROM dbo.TblyearsData y WHERE y.Account_Code = A.Account_Code)" & vbCrLf
sql = sql & "            OR A.last_account = 0" & vbCrLf
sql = sql & "        )" & vbCrLf
sql = sql & ")" & vbCrLf

'========================
' Aggregates („—… Ê«Õœ…)
'========================
sql = sql & ", D_2025 AS (" & vbCrLf
sql = sql & "    SELECT d.Account_Code," & vbCrLf
sql = sql & "           Debit  = SUM(CASE WHEN d.Credit_Or_Debit = 0 THEN d.Value ELSE 0 END)," & vbCrLf
sql = sql & "           Credit = SUM(CASE WHEN d.Credit_Or_Debit = 1 THEN d.Value ELSE 0 END)" & vbCrLf
sql = sql & "    FROM dbo.DOUBLE_ENTREY_VOUCHERS d" & vbCrLf
sql = sql & "    WHERE d.Posted IS NULL" & vbCrLf
sql = sql & "      AND d.RecordDate >= @FromDate AND d.RecordDate <= @ToDate" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "      AND d.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "    GROUP BY d.Account_Code" & vbCrLf
sql = sql & ")" & vbCrLf

sql = sql & ", D_Open_From_To AS (" & vbCrLf
sql = sql & "    SELECT d.Account_Code," & vbCrLf
sql = sql & "           Net = SUM(CASE WHEN d.Credit_Or_Debit = 0 THEN d.Value" & vbCrLf
sql = sql & "                        WHEN d.Credit_Or_Debit = 1 THEN -d.Value ELSE 0 END)" & vbCrLf
sql = sql & "    FROM dbo.DOUBLE_ENTREY_VOUCHERS d" & vbCrLf
sql = sql & "    WHERE d.Posted IS NULL" & vbCrLf
sql = sql & "      AND d.RecordDate >= @OpenFrom AND d.RecordDate <= @OpenTo" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "      AND d.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "    GROUP BY d.Account_Code" & vbCrLf
sql = sql & ")" & vbCrLf

sql = sql & ", D1_Opening AS (" & vbCrLf
sql = sql & "    SELECT d1.Account_Code," & vbCrLf
sql = sql & "           Net = SUM(CASE WHEN d1.Credit_Or_Debit = 0 THEN d1.Value" & vbCrLf
sql = sql & "                        WHEN d1.Credit_Or_Debit = 1 THEN -d1.Value ELSE 0 END)" & vbCrLf
sql = sql & "    FROM dbo.DOUBLE_ENTREY_VOUCHERS1 d1" & vbCrLf
sql = sql & "    WHERE d1.Posted IS NULL" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "      AND d1.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "      AND d1.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "      AND d1.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "    GROUP BY d1.Account_Code" & vbCrLf
sql = sql & ")" & vbCrLf

sql = sql & ", D_FromStartYear_To_BeforeFrom AS (" & vbCrLf
sql = sql & "    SELECT d.Account_Code," & vbCrLf
sql = sql & "           Net = SUM(CASE WHEN d.Credit_Or_Debit = 0 THEN d.Value" & vbCrLf
sql = sql & "                        WHEN d.Credit_Or_Debit = 1 THEN -d.Value ELSE 0 END)" & vbCrLf
sql = sql & "    FROM dbo.DOUBLE_ENTREY_VOUCHERS d" & vbCrLf
sql = sql & "    WHERE d.Posted IS NULL" & vbCrLf
sql = sql & "      AND d.RecordDate >= @StartYear AND d.RecordDate < @FromDate" & vbCrLf
If val(DCActivity.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BrcnActivety & ")" & vbCrLf
If val(DCRegionID.BoundText) <> 0 Then sql = sql & "      AND d.branch_id in (" & BranshesReg & ")" & vbCrLf
If val(dcBranch.BoundText) <> 0 Then sql = sql & "      AND d.branch_id = " & val(dcBranch.BoundText) & vbCrLf
sql = sql & "    GROUP BY d.Account_Code" & vbCrLf
sql = sql & ")" & vbCrLf

'========================
' Final Select (‰›” «·√⁄„œ…)
'========================
sql = sql & "SELECT" & vbCrLf
sql = sql & "    a.last_account," & vbCrLf
sql = sql & "    a.ProfitBalance," & vbCrLf
sql = sql & "    a.Parent_Account_Code," & vbCrLf
sql = sql & "    a.AccountTypes," & vbCrLf
sql = sql & "    a.Account_Code," & vbCrLf
sql = sql & "    a.Account_Serial," & vbCrLf
sql = sql & "    a.Account_Name," & vbCrLf
sql = sql & "    a.Account_NameEng," & vbCrLf

sql = sql & "    debitBalance = COALESCE(SUM(CASE" & vbCrLf
sql = sql & "        WHEN a.last_account = 1 AND d.Account_Code = a.Account_Code THEN d.Debit" & vbCrLf
sql = sql & "        WHEN a.last_account = 0 AND d.Account_Code LIKE a.Account_Code + 'a%' THEN d.Debit" & vbCrLf
sql = sql & "        ELSE 0 END),0)," & vbCrLf

sql = sql & "    CreditBalance = COALESCE(SUM(CASE" & vbCrLf
sql = sql & "        WHEN a.last_account = 1 AND d.Account_Code = a.Account_Code THEN d.Credit" & vbCrLf
sql = sql & "        WHEN a.last_account = 0 AND d.Account_Code LIKE a.Account_Code + 'a%' THEN d.Credit" & vbCrLf
sql = sql & "        ELSE 0 END),0)," & vbCrLf

sql = sql & "    OpeningBalance = COALESCE(SUM(CASE" & vbCrLf
sql = sql & "        WHEN a.last_account = 1 AND o1.Account_Code = a.Account_Code THEN o1.Net" & vbCrLf
sql = sql & "        WHEN a.last_account = 0 AND o1.Account_Code LIKE a.Account_Code + 'a%' THEN o1.Net" & vbCrLf
sql = sql & "        ELSE 0 END),0)," & vbCrLf

sql = sql & "    OpeningBalancebeformdateMinus1 = COALESCE(SUM(CASE" & vbCrLf
sql = sql & "        WHEN a.last_account = 1 AND ob.Account_Code = a.Account_Code THEN ob.Net" & vbCrLf
sql = sql & "        WHEN a.last_account = 0 AND ob.Account_Code LIKE a.Account_Code + 'a%' THEN ob.Net" & vbCrLf
sql = sql & "        ELSE 0 END),0)," & vbCrLf

sql = sql & "    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 = COALESCE(SUM(CASE" & vbCrLf
sql = sql & "        WHEN a.last_account = 1 AND cy.Account_Code = a.Account_Code THEN cy.Net" & vbCrLf
sql = sql & "        WHEN a.last_account = 0 AND cy.Account_Code LIKE a.Account_Code + 'a%' THEN cy.Net" & vbCrLf
sql = sql & "        ELSE 0 END),0" & vbCrLf
sql = sql & "    )" & vbCrLf

sql = sql & "FROM A a" & vbCrLf

sql = sql & "LEFT JOIN D_2025 d" & vbCrLf
sql = sql & "  ON (a.last_account = 1 AND d.Account_Code = a.Account_Code)" & vbCrLf
sql = sql & "  OR (a.last_account = 0 AND d.Account_Code LIKE a.Account_Code + 'a%')" & vbCrLf

sql = sql & "LEFT JOIN D1_Opening o1" & vbCrLf
sql = sql & "  ON (a.last_account = 1 AND o1.Account_Code = a.Account_Code)" & vbCrLf
sql = sql & "  OR (a.last_account = 0 AND o1.Account_Code LIKE a.Account_Code + 'a%')" & vbCrLf

sql = sql & "LEFT JOIN D_Open_From_To ob" & vbCrLf
sql = sql & "  ON (a.last_account = 1 AND ob.Account_Code = a.Account_Code)" & vbCrLf
sql = sql & "  OR (a.last_account = 0 AND ob.Account_Code LIKE a.Account_Code + 'a%')" & vbCrLf

sql = sql & "LEFT JOIN D_FromStartYear_To_BeforeFrom cy" & vbCrLf
sql = sql & "  ON (a.last_account = 1 AND cy.Account_Code = a.Account_Code)" & vbCrLf
sql = sql & "  OR (a.last_account = 0 AND cy.Account_Code LIKE a.Account_Code + 'a%')" & vbCrLf

sql = sql & "GROUP BY" & vbCrLf
sql = sql & "    a.last_account, a.ProfitBalance, a.Parent_Account_Code, a.AccountTypes," & vbCrLf
sql = sql & "    a.Account_Code, a.Account_Serial, a.Account_Name, a.Account_NameEng" & vbCrLf
sql = sql & "ORDER BY a.Account_Serial;" & vbCrLf

 
  
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Incomstatement2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Incomstatement2.rpt"
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
       Msg = "·« ÌÊÃœ »Ì«‰« "
     Else
     Msg = "No Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim desc As String
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    xReport.ParameterFields(7).AddCurrentValue desc
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

 
 
'
Function print_report41Old(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer
  
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
 
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
  Dim HideLastAccount  As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim ShowOnlyLevelAcc As Integer
   Dim BrcnActivety As String
   FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  
  
           
   If SystemOptions.UserInterface = ArabicInterface Then
                 X = val(InputBox("Specify Level"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            
            account_level = val(X)
             
 
            
            
         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
   
            If HideZeroBalance = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
                     If SystemOptions.UserInterface = ArabicInterface Then
               HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            Else
                HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            End If
                  If HideLastAccount = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
   
   
            
            
                        

                  If ShowOnlyLevelAcc = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
            
            
          Dim BranshesReg As String
      
         If val(DCRegionID.BoundText) <> 0 Then
         BranshesReg = BranchRegion(DCRegionID.BoundText)
         End If
         If val(DCActivity.BoundText) <> 0 Then
         BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
         End If


  updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True

  sql = " SELECT   last_account, ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name,Account_NameEng, debitBalance ="
  sql = sql & "                         (SELECT     SUM(DEV_Value1)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
 'sql = sql & "            AND d.Account_Code   like  A.Account_Code+'%'"
sql = sql & "  AND ( ( last_account= 0 and   d.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   d.Account_Code   =  A.Account_Code ) )"
  sql = sql & "           and(d.Posted IS NULL)"
 If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
  End If
  
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BranshesReg & ")"
  End If
  
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d.branch_id =" & val(dcBranch.BoundText) & ""
  End If
 sql = sql & "  ) x),"
   sql = sql & "                    CreditBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & "  AND d1.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
 'sql = sql & "                  AND d1.Account_Code   like  A.Account_Code+'%'"
    sql = sql & "     AND ( ( last_account= 0 and   d1.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   d1.Account_Code   =  A.Account_Code ) )"
    
 
  sql = sql & "                 and(d1.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
  End If
 If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                     OpeningBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  sql = sql & "                                                   WHERE     ( "
  'do.Account_Code   like  A.Account_Code+'%'
 sql = sql & "       ( ( last_account= 0 and   do.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   do.Account_Code   =  A.Account_Code ) )"
  sql = sql & "      and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
 sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
 End If
sql = sql & "  )) x),"
  sql = sql & "    OpeningBalancebeformdateMinus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") "
 ' AND do.Account_Code   like  A.Account_Code+'%'
sql = sql & " and      ( ( last_account= 0 and   do.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   do.Account_Code   =  A.Account_Code ) )"

 sql = sql & "       and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"
'  AND do.Account_Code  like  A.Account_Code+'%'
sql = sql & "  and      ( ( last_account= 0 and   do.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   do.Account_Code   =  A.Account_Code ) )"

 sql = sql & "    and(do.Posted IS NULL) "
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x)"
  sql = sql & " FROM         ACCOUNTS A"
  If HideLastAccount = True Then
    sql = sql & " WHERE     A.last_account = 0   "
  Else
  sql = sql & " WHERE  1=1   "
  End If
  
  If (TxtAccountCode.text) <> "" Then
  sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
  End If
  '*****************************************************
  
  sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
  sql = sql & "    Where 1 = 1"
 
    If val(DCActivity.BoundText) <> 0 Then
 sql = sql & " and branch_id in (" & BrcnActivety & ")"
 End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
 End If
   sql = sql & "   )"
 
  sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
 sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & "   )"
  
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData  Where 1 = 1)"
 ' ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
' ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

   sql = sql & " or A.Account_Code in( SELECT     Account_Code  FROM         dbo.ACCOUNTS WHERE     last_account = 0)"
 
 

  sql = sql & "   )"
'************************************************************************************************************
      If account_level <> 0 Then
            If ShowOnlyLevelAcc = vbYes Then
                     sql = sql & " and len(account_code) - len(replace(account_code,'a',''))  = " & account_level
              Else
                   sql = sql & " and len(account_code) - len(replace(account_code,'a',''))  <= " & account_level
              End If
    End If

sql = sql & " and AccountTypes=2"

    sql = sql & "order by Account_Serial "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Incomstatement2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Incomstatement2.rpt"
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
       Msg = "·« ÌÊÃœ »Ì«‰« "
     Else
     Msg = "No Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim desc As String
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    xReport.ParameterFields(7).AddCurrentValue desc
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

Function print_report40Del(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim AccountTypes As Integer
  
    Dim OpeningBalancebeformdateMinus1 As Double
    Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
    Dim NewOpinning As Double
    Dim OpeningBalance As Double
    Dim ProfitBalance As Double
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
 
    Dim i As Integer
    Dim BranchID As String
    Dim HideZeroBalance As Integer
    Dim HideLastAccount  As Integer
    Dim openingBalanceDate As Date
    Dim FromdateMinus1 As Date
    Dim StartCurrentDate As Date
    Dim ShowOnlyLevelAcc As Integer
    Dim BrcnActivety As String
    Dim BranshesReg As String
    Dim X As Variant
    Dim account_level As Long
    Dim desc As String

    FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
           
    If SystemOptions.UserInterface = ArabicInterface Then
        X = val(InputBox("Specify Level"))
    Else
        X = val(InputBox("Specify Level"))
    End If
    account_level = val(X)
             
    If SystemOptions.UserInterface = ArabicInterface Then
        HideZeroBalance = MsgBox("Hide Zero Account", vbInformation + vbYesNoCancel)
    Else
        HideZeroBalance = MsgBox("Hide Zero Account", vbInformation + vbYesNoCancel)
    End If
    If HideZeroBalance = 2 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
    If SystemOptions.UserInterface = ArabicInterface Then
        HideLastAccount = MsgBox("Hide Last Account", vbInformation + vbYesNoCancel)
    Else
        HideLastAccount = MsgBox("Hide Last Account", vbInformation + vbYesNoCancel)
    End If
    If HideLastAccount = 2 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
                        
    If SystemOptions.UserInterface = ArabicInterface Then
        ShowOnlyLevelAcc = MsgBox("Show only this level ?", vbInformation + vbYesNoCancel)
    Else
        ShowOnlyLevelAcc = MsgBox("Show only this level ?", vbInformation + vbYesNoCancel)
    End If
    If ShowOnlyLevelAcc = 2 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
            
    If val(DCRegionID.BoundText) <> 0 Then
        BranshesReg = BranchRegion(DCRegionID.BoundText)
    End If
    If val(DCActivity.BoundText) <> 0 Then
        BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
    End If

    updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True

    '================= ≈⁄œ«œ „ €Ì—«  ‘«∆⁄… =================
'  Dim sql As String
Dim FromDate$, ToDate$, OpenFrom$, fromMinus1$, startCurrent$
FromDate = SQLDate(Me.DTPickerAccFrom.value, True)
ToDate = SQLDate(Me.DTPickerAccTo.value, True)
OpenFrom = SQLDate(openingBalanceDate, True)
fromMinus1 = SQLDate(FromdateMinus1, True)
startCurrent = SQLDate(StartCurrentDate, True)

' ›·« — «·›—Ê⁄/«·‰‘«ÿ/«·≈ﬁ·Ì„
Dim flt_d$, flt_do1$
If val(DCActivity.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id IN (" & BrcnActivety & ")"
    flt_do1 = flt_do1 & " AND do1.branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id IN (" & BranshesReg & ")"
    flt_do1 = flt_do1 & " AND do1.branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id = " & val(dcBranch.BoundText)
    flt_do1 = flt_do1 & " AND do1.branch_id = " & val(dcBranch.BoundText)
End If

'==========================================================
'            ?? SQL „ıÕ”¯Û‰: ﬂÊ»Ì-»Ì”  ??
'==========================================================

sql = ""

' ≈”ﬁ«ÿ «·Ãœ«Ê· «·„ƒﬁ … ·Ê „ÊÃÊœ…
AddSQL sql, "IF OBJECT_ID('tempdb..#H')  IS NOT NULL DROP TABLE #H;"
AddSQL sql, "IF OBJECT_ID('tempdb..#VS') IS NOT NULL DROP TABLE #VS;"
AddSQL sql, "IF OBJECT_ID('tempdb..#V1S') IS NOT NULL DROP TABLE #V1S;"

' ====== 1)  ÕœÌœ «·Õ”«»«  «·„ÿ·Ê»… Õ”» «·⁄„ﬁ ›Ì #H ======
AddSQL sql, "WITH Acc AS ("
AddSQL sql, "  SELECT A.Account_ID,"
AddSQL sql, "         LTRIM(RTRIM(A.Account_Code)) AS Account_Code,"
AddSQL sql, "         NULLIF(LTRIM(RTRIM(A.Parent_Account_Code)), N'') AS Parent_Account_Code,"
AddSQL sql, "         A.last_account,"
AddSQL sql, "         A.Account_Serial,"
AddSQL sql, "         A.AccountTypes,"
AddSQL sql, "         A.ProfitBalance,"
AddSQL sql, "         A.Account_Name,"
AddSQL sql, "         A.Account_NameEng"
AddSQL sql, "  FROM dbo.ACCOUNTS A WITH (NOLOCK)"
AddSQL sql, "), Roots AS ("
AddSQL sql, "  SELECT R.Account_ID, R.Account_Code, R.Parent_Account_Code, R.last_account,"
AddSQL sql, "         R.Account_Serial, R.AccountTypes, R.ProfitBalance, R.Account_Name, R.Account_NameEng,"
AddSQL sql, "         1 AS ParentDepth"
AddSQL sql, "  FROM Acc R"
AddSQL sql, "  LEFT JOIN Acc P ON P.Account_Code = R.Parent_Account_Code"
AddSQL sql, "  WHERE R.Parent_Account_Code IS NULL OR R.Parent_Account_Code = N'r' OR P.Account_Code IS NULL"
AddSQL sql, "), Hierarchy AS ("
AddSQL sql, "  SELECT * FROM Roots"
AddSQL sql, "  UNION ALL"
AddSQL sql, "  SELECT C.Account_ID, C.Account_Code, C.Parent_Account_Code, C.last_account,"
AddSQL sql, "         C.Account_Serial, C.AccountTypes, C.ProfitBalance, C.Account_Name, C.Account_NameEng,"
AddSQL sql, "         H.ParentDepth + 1"
AddSQL sql, "  FROM Acc C"
AddSQL sql, "  JOIN Hierarchy H ON C.Parent_Account_Code = H.Account_Code"
AddSQL sql, ")"
AddSQL sql, "SELECT * INTO #H FROM Hierarchy H WHERE 1=1"

' ·«ÕŸ ≈œ—«Ã ‘—Êÿ «·‹VB «·ﬁ«»·… ·· €ÌÌ—
If account_level <> 0 Then
    If ShowOnlyLevelAcc = vbYes Then
        AddSQL sql, "  AND H.ParentDepth = " & CStr(account_level)
    Else
        AddSQL sql, "  AND H.ParentDepth <= " & CStr(account_level)
    End If
End If
If (TxtAccountCode.text) <> "" Then
    AddSQL sql, "  AND H.Account_Serial = '" & Replace(TxtAccountCode.text, "'", "''") & "'"
End If
AddSQL sql, ";"
AddSQL sql, "CREATE UNIQUE CLUSTERED INDEX IX_H_AccountID ON #H(Account_ID);"
AddSQL sql, "CREATE NONCLUSTERED INDEX IX_H_Code ON #H(Account_Code);"
AddSQL sql, "CREATE NONCLUSTERED INDEX IX_H_Last ON #H(last_account);"

' ====== 2)  Ã„Ì⁄ «·Õ—ﬂ«  ›Ì #VS Ê #V1S ======
AddSQL sql, "SELECT d.Account_Code,"
AddSQL sql, "       SUM(CASE WHEN d.Credit_Or_Debit=0 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS debitBalance,"
AddSQL sql, "       SUM(CASE WHEN d.Credit_Or_Debit=1 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS CreditBalance,"
AddSQL sql, "       SUM(CASE WHEN d.RecordDate >= " & OpenFrom & " AND d.RecordDate <= " & fromMinus1 & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformdateMinus1,"
AddSQL sql, "       SUM(CASE WHEN d.RecordDate >= " & startCurrent & " AND d.RecordDate < " & FromDate & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
AddSQL sql, "INTO #VS"
AddSQL sql, "FROM dbo.DOUBLE_ENTREY_VOUCHERS d WITH (NOLOCK)"
AddSQL sql, "WHERE d.Posted IS NULL" & flt_d
AddSQL sql, "GROUP BY d.Account_Code;"
AddSQL sql, "CREATE CLUSTERED INDEX IX_VS_Code ON #VS(Account_Code);"

AddSQL sql, "SELECT do1.Account_Code,"
AddSQL sql, "       SUM(CASE WHEN do1.Credit_Or_Debit=0 THEN do1.Value ELSE -do1.Value END) AS OpeningBalance"
AddSQL sql, "INTO #V1S"
AddSQL sql, "FROM dbo.DOUBLE_ENTREY_VOUCHERS1 do1 WITH (NOLOCK)"
AddSQL sql, "WHERE do1.Posted IS NULL" & flt_do1
AddSQL sql, "GROUP BY do1.Account_Code;"
AddSQL sql, "CREATE CLUSTERED INDEX IX_V1S_Code ON #V1S(Account_Code);"

' ====== 3) OpeningBalanceCorrected ======
AddSQL sql, ";WITH OBC AS ("
AddSQL sql, "  SELECT A.Account_ID, SUM(ISNULL(V.OpeningBalance,0)) AS OpeningBalance"
AddSQL sql, "  FROM dbo.ACCOUNTS A WITH (NOLOCK)"
AddSQL sql, "  LEFT JOIN #V1S V"
AddSQL sql, "    ON V.Account_Code = A.Account_Code"
AddSQL sql, "    OR (V.Account_Code >= A.Account_Code + 'a' AND V.Account_Code < A.Account_Code + 'b')"
AddSQL sql, "  GROUP BY A.Account_ID"
AddSQL sql, ")"

' ====== 4) «·‰ ÌÃ… «·‰Â«∆Ì… ======
AddSQL sql, "SELECT H.last_account, H.ProfitBalance, H.Parent_Account_Code, H.AccountTypes, H.Account_Code, H.Account_Serial, H.Account_Name, H.Account_NameEng,"
AddSQL sql, "       SUM(F.debitBalance) AS debitBalance,"
AddSQL sql, "       SUM(F.CreditBalance) AS CreditBalance,"
AddSQL sql, "       SUM(F.OpeningBalance) AS OpeningBalance,"
AddSQL sql, "       SUM(F.OpeningBalancebeformdateMinus1) AS OpeningBalancebeformdateMinus1,"
AddSQL sql, "       SUM(F.OpeningBalancebeformStartCurrentyearTOFromDAteminus1) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
AddSQL sql, "FROM #H H"
AddSQL sql, "LEFT JOIN ("
AddSQL sql, "    SELECT A.Account_ID,"
AddSQL sql, "           ISNULL(VS.debitBalance,0) AS debitBalance,"
AddSQL sql, "           ISNULL(VS.CreditBalance,0) AS CreditBalance,"
AddSQL sql, "           ISNULL(OBC.OpeningBalance,0) AS OpeningBalance,"
AddSQL sql, "           ISNULL(VS.OpeningBalancebeformdateMinus1,0) AS OpeningBalancebeformdateMinus1,"
AddSQL sql, "           ISNULL(VS.OpeningBalancebeformStartCurrentyearTOFromDAteminus1,0) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
AddSQL sql, "    FROM dbo.ACCOUNTS A WITH (NOLOCK)"
AddSQL sql, "    LEFT JOIN #VS VS  ON VS.Account_Code = A.Account_Code"
AddSQL sql, "    LEFT JOIN OBC OBC ON OBC.Account_ID = A.Account_ID"
AddSQL sql, "    WHERE A.last_account = 1"
AddSQL sql, "    UNION ALL"
AddSQL sql, "    SELECT A.Account_ID,"
AddSQL sql, "           SUM(ISNULL(VS.debitBalance,0)) AS debitBalance,"
AddSQL sql, "           SUM(ISNULL(VS.CreditBalance,0)) AS CreditBalance,"
AddSQL sql, "           MAX(ISNULL(OBC.OpeningBalance,0)) AS OpeningBalance,"
AddSQL sql, "           SUM(ISNULL(VS.OpeningBalancebeformdateMinus1,0)) AS OpeningBalancebeformdateMinus1,"
AddSQL sql, "           SUM(ISNULL(VS.OpeningBalancebeformStartCurrentyearTOFromDAteminus1,0)) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
AddSQL sql, "    FROM dbo.ACCOUNTS A WITH (NOLOCK)"
AddSQL sql, "    LEFT JOIN #VS VS ON VS.Account_Code >= A.Account_Code + 'a' AND VS.Account_Code < A.Account_Code + 'b'"
AddSQL sql, "    LEFT JOIN OBC OBC ON OBC.Account_ID = A.Account_ID"
AddSQL sql, "    WHERE A.last_account = 0"
AddSQL sql, "    GROUP BY A.Account_ID"
AddSQL sql, ") F ON F.Account_ID = H.Account_ID"
AddSQL sql, "WHERE 1=1"

If HideLastAccount = 6 Then AddSQL sql, "  AND H.last_account = 0"
If (TxtAccountCode.text) <> "" Then AddSQL sql, "  AND H.Account_Serial = '" & Replace(TxtAccountCode.text, "'", "''") & "'"
If HideZeroBalance = 6 Then _
    AddSQL sql, "  AND (F.debitBalance <> 0 OR F.CreditBalance <> 0 OR F.OpeningBalance <> 0 OR F.OpeningBalancebeformdateMinus1 <> 0 OR F.OpeningBalancebeformStartCurrentyearTOFromDAteminus1 <> 0 OR H.last_account = 0)"

AddSQL sql, "GROUP BY H.last_account, H.ProfitBalance, H.Parent_Account_Code, H.AccountTypes, H.Account_Code, H.Account_Serial, H.Account_Name, H.Account_NameEng"
AddSQL sql, "ORDER BY H.Account_Serial"
AddSQL sql, "OPTION (MAXRECURSION 32767, RECOMPILE);"


' IMPORTANT: “Ê¯œ «·„Â·… ‘ÊÌ… ·√‰ √Ê·  ‘€Ì· ÂÌ»‰¯Ì «·≈Õ’«∆Ì« 
Cn.CommandTimeout = 1300
On Error Resume Next
Cn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
On Error GoTo 0

With RsData
    .CursorLocation = adUseServer
    .Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
End With


    '================= «Œ Ì«— „·› «· ﬁ—Ì— =================
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSa2.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSaE.rpt"
    End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '================= › Õ «·‹Recordset =================
   ' Cn.CommandTimeout = 120
    On Error Resume Next
    Cn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
    On Error GoTo 0

    With RsData
        .CursorLocation = adUseServer
        .Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End With

    If RsData.BOF Or RsData.EOF Then
        MsgBox IIf(SystemOptions.UserInterface = ArabicInterface, "·« ÌÊÃœ »Ì«‰« ", "No Data"), vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close: Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    '================= »«—«„Ì — «· ﬁ—Ì— =================
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        StrReportTitle = ""
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName
        StrReportTitle = ""
    End If

    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·‰‘«ÿ: " & DCActivity.text & CHR(13)
        Else
            desc = desc & "Activity: " & DCActivity.text & CHR(13)
        End If
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·≈ﬁ·Ì„: " & DCRegionID.text & CHR(13)
        Else
            desc = desc & "Region: " & DCRegionID.text & CHR(13)
        End If
    End If
    If val(dcBranch.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·›—⁄: " & dcBranch.text & CHR(13)
        Else
            desc = desc & "Branch: " & dcBranch.text & CHR(13)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
        xReport.ParameterFields(6).AddCurrentValue 1
    Else
        xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
        xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
        xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    xReport.ParameterFields(7).AddCurrentValue desc

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

ErrTrap:
End Function

'

'
Function print_report40(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer
  
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
 
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
  Dim HideLastAccount  As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim ShowOnlyLevelAcc As Integer
   Dim BrcnActivety As String
   FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  
  
           
   If SystemOptions.UserInterface = ArabicInterface Then
               X = val(InputBox("Specify Level"))
            Else
                X = val(InputBox("Specify Level"))
            End If
        
            
            account_level = val(X)
             
 
            
            
         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
   
            If HideZeroBalance = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
                     If SystemOptions.UserInterface = ArabicInterface Then
                HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            Else
                HideLastAccount = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            End If
                  If HideLastAccount = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
   
   
            
            
                        
                     If SystemOptions.UserInterface = ArabicInterface Then
                ShowOnlyLevelAcc = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            Else
                ShowOnlyLevelAcc = MsgBox("Hide LAst Account  ", vbInformation + vbYesNoCancel)
            End If
              If ShowOnlyLevelAcc = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
            
            
          Dim BranshesReg As String
      
         If val(DCRegionID.BoundText) <> 0 Then
         BranshesReg = BranchRegion(DCRegionID.BoundText)
         End If
         If val(DCActivity.BoundText) <> 0 Then
         BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
         End If


  updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True

  sql = " SELECT   last_account, ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name,Account_NameEng, debitBalance ="
  sql = sql & "                         (SELECT     SUM(DEV_Value1)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
 'sql = sql & "            AND d.Account_Code   like  A.Account_Code+'%'"
sql = sql & "  AND ( ( last_account= 0 and   d.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   d.Account_Code   =  A.Account_Code ) )"
  sql = sql & "           and(d.Posted IS NULL)"
 If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
  End If
  
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BranshesReg & ")"
  End If
  
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d.branch_id =" & val(dcBranch.BoundText) & ""
  End If
 sql = sql & "  ) x),"
   sql = sql & "                    CreditBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & "  AND d1.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
 'sql = sql & "                  AND d1.Account_Code   like  A.Account_Code+'%'"
    sql = sql & "     AND ( ( last_account= 0 and   d1.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   d1.Account_Code   =  A.Account_Code ) )"
    
 
  sql = sql & "                 and(d1.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
  End If
 If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                     OpeningBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  sql = sql & "                                                   WHERE     ( "
  'do.Account_Code   like  A.Account_Code+'%'
 sql = sql & "       ( (   do.Account_Code   like  A.Account_Code+'a%' )  or ( do.Account_Code   =  A.Account_Code ) )"
  sql = sql & "      and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
 sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
 End If
sql = sql & "  )) x),"
  sql = sql & "    OpeningBalancebeformdateMinus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") "
 ' AND do.Account_Code   like  A.Account_Code+'%'
sql = sql & " and      ( ( last_account= 0 and   do.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   do.Account_Code   =  A.Account_Code ) )"

 sql = sql & "       and(do.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"
'  AND do.Account_Code  like  A.Account_Code+'%'
sql = sql & "  and      ( ( last_account= 0 and   do.Account_Code   like  A.Account_Code+'a%' )  or ( last_account= 1 and   do.Account_Code   =  A.Account_Code ) )"

 sql = sql & "    and(do.Posted IS NULL) "
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x)"
  sql = sql & " FROM         ACCOUNTS A"
  If HideLastAccount = True Then
    sql = sql & " WHERE     A.last_account = 0   "
  Else
  sql = sql & " WHERE  1=1   "
  End If
  
  If (TxtAccountCode.text) <> "" Then
  sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
  End If
  '*****************************************************
  
  sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
  sql = sql & "    Where 1 = 1"
 
    If val(DCActivity.BoundText) <> 0 Then
 sql = sql & " and branch_id in (" & BrcnActivety & ")"
 End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
 End If
   sql = sql & "   )"
 
  sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
 sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & "   )"
  
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData  Where 1 = 1)"
 ' ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
' ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

   sql = sql & " or A.Account_Code in( SELECT     Account_Code  FROM         dbo.ACCOUNTS WHERE     last_account = 0)"
 
 

  sql = sql & "   )"
'************************************************************************************************************
      If account_level <> 0 Then
            If ShowOnlyLevelAcc = vbYes Then
                     sql = sql & " and ((len(account_code) - len(replace(account_code,'a',''))  = " & account_level
              Else
                   sql = sql & " and  ((len(account_code) - len(replace(account_code,'a',''))  <= " & account_level
              End If
              
              sql = sql & " and len(account_code) - len(replace(account_code,'a',''))  <= " & account_level & ") OR  Account_Name LIKE '%?????%') "
              
    End If

    sql = sql & "order by Account_Serial "
    
    
    
'    updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True
'================= ≈⁄œ«œ „ €Ì—«  ‘«∆⁄… =================
Dim FromDate$, ToDate$, OpenFrom$, fromMinus1$, startCurrent$
FromDate = SQLDate(Me.DTPickerAccFrom.value, True)
ToDate = SQLDate(Me.DTPickerAccTo.value, True)
OpenFrom = SQLDate(openingBalanceDate, True)
fromMinus1 = SQLDate(FromdateMinus1, True)
startCurrent = SQLDate(StartCurrentDate, True)

Dim lvlClause$
If account_level <> 0 Then
    If ShowOnlyLevelAcc = vbYes Then
        lvlClause = " AND ( (LEN(account_code) - LEN(REPLACE(account_code,'a','')) = " & account_level & ") OR Account_Name LIKE N'%?????%')"
    Else
        lvlClause = " AND ( (LEN(account_code) - LEN(REPLACE(account_code,'a','')) <= " & account_level & ") OR Account_Name LIKE N'%?????%')"
    End If
End If

' ›·« — «·›—Ê⁄ (»‰›” „‰ÿﬁﬂ)
Dim flt_d$, flt_d1$, flt_do1$, flt_e$, flt_e1$
If val(DCActivity.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id IN (" & BrcnActivety & ")"
    flt_d1 = flt_d1 & " AND d1.branch_id IN (" & BrcnActivety & ")"
    flt_do1 = flt_do1 & " AND do1.branch_id IN (" & BrcnActivety & ")"
    flt_e = flt_e & " AND e.branch_id IN (" & BrcnActivety & ")"
    flt_e1 = flt_e1 & " AND e1.branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id IN (" & BranshesReg & ")"
    flt_d1 = flt_d1 & " AND d1.branch_id IN (" & BranshesReg & ")"
    flt_do1 = flt_do1 & " AND do1.branch_id IN (" & BranshesReg & ")"
    flt_e = flt_e & " AND e.branch_id IN (" & BranshesReg & ")"
    flt_e1 = flt_e1 & " AND e1.branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    flt_d = flt_d & " AND d.branch_id = " & val(dcBranch.BoundText)
    flt_d1 = flt_d1 & " AND d1.branch_id = " & val(dcBranch.BoundText)
    flt_do1 = flt_do1 & " AND do1.branch_id = " & val(dcBranch.BoundText)
    flt_e = flt_e & " AND e.branch_id = " & val(dcBranch.BoundText)
    flt_e1 = flt_e1 & " AND e1.branch_id = " & val(dcBranch.BoundText)
End If

'================= »‰«¡ «·«” ⁄·«„ =================
sql = ""
sql = sql & "SELECT  last_account, ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng,"
sql = sql & "        debitBalance = SUM(debitBalance),"
sql = sql & "        CreditBalance = SUM(CreditBalance),"
sql = sql & "        OpeningBalance = SUM(OpeningBalance),"
sql = sql & "        OpeningBalancebeformdateMinus1 = SUM(OpeningBalancebeformdateMinus1),"
sql = sql & "        OpeningBalancebeformStartCurrentyearTOFromDAteminus1 = SUM(OpeningBalancebeformStartCurrentyearTOFromDAteminus1)"
sql = sql & " FROM ("

'---------- Ã“¡ «·Õ”«»«  «·‰Â«∆Ì… (last_account = 1) ----------
sql = sql & " SELECT A.last_account, A.ProfitBalance, A.Parent_Account_Code, A.AccountTypes, A.Account_Code, A.Account_Serial, A.Account_Name, A.Account_NameEng,"
sql = sql & "        SUM(CASE WHEN d.Credit_Or_Debit = 0 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS debitBalance,"
sql = sql & "        SUM(CASE WHEN d.Credit_Or_Debit = 1 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS CreditBalance,"
sql = sql & "        SUM(CASE WHEN do1.Credit_Or_Debit = 0 THEN do1.Value WHEN do1.Credit_Or_Debit = 1 THEN -do1.Value ELSE 0 END) AS OpeningBalance,"
sql = sql & "        SUM(CASE WHEN d.RecordDate >= " & OpenFrom & " AND d.RecordDate <= " & fromMinus1 & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformdateMinus1,"
sql = sql & "        SUM(CASE WHEN d.RecordDate >= " & startCurrent & " AND d.RecordDate < " & FromDate & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
sql = sql & " FROM ACCOUNTS A WITH (NOLOCK)"
sql = sql & " LEFT JOIN dbo.DOUBLE_ENTREY_VOUCHERS  AS d   WITH (NOLOCK)"
sql = sql & "   ON d.Posted IS NULL AND d.Account_Code = A.Account_Code" & flt_d
sql = sql & " LEFT JOIN dbo.DOUBLE_ENTREY_VOUCHERS1 AS do1 WITH (NOLOCK)"
sql = sql & "   ON do1.Posted IS NULL AND do1.Account_Code = A.Account_Code" & flt_do1
sql = sql & " WHERE A.last_account = 1"
If HideLastAccount = True Then sql = sql & " AND 1=0"
If (TxtAccountCode.text) <> "" Then sql = sql & " AND A.Account_Serial = '" & TxtAccountCode.text & "'"
sql = sql & lvlClause
sql = sql & " AND ( EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS  AS e  WITH (NOLOCK) WHERE e.Account_Code = A.Account_Code" & flt_e & ")"
sql = sql & "   OR EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS1 AS e1 WITH (NOLOCK) WHERE e1.Account_Code = A.Account_Code" & flt_e1 & ")"
sql = sql & "   OR EXISTS (SELECT 1 FROM dbo.TblyearsData        AS t  WITH (NOLOCK) WHERE t.Account_Code  = A.Account_Code) )"
sql = sql & " GROUP BY A.last_account, A.ProfitBalance, A.Parent_Account_Code, A.AccountTypes, A.Account_Code, A.Account_Serial, A.Account_Name, A.Account_NameEng"

sql = sql & " UNION ALL "

'---------- Ã“¡ Õ”«»«  «·√» (last_account = 0) ----------
sql = sql & " SELECT A.last_account, A.ProfitBalance, A.Parent_Account_Code, A.AccountTypes, A.Account_Code, A.Account_Serial, A.Account_Name, A.Account_NameEng,"
sql = sql & "        SUM(CASE WHEN d.Credit_Or_Debit = 0 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS debitBalance,"
sql = sql & "        SUM(CASE WHEN d.Credit_Or_Debit = 1 AND d.RecordDate >= " & FromDate & " AND d.RecordDate <= " & ToDate & " THEN d.Value ELSE 0 END) AS CreditBalance,"
sql = sql & "        SUM(CASE WHEN do1.Credit_Or_Debit = 0 THEN do1.Value WHEN do1.Credit_Or_Debit = 1 THEN -do1.Value ELSE 0 END) AS OpeningBalance,"
sql = sql & "        SUM(CASE WHEN d.RecordDate >= " & OpenFrom & " AND d.RecordDate <= " & fromMinus1 & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformdateMinus1,"
sql = sql & "        SUM(CASE WHEN d.RecordDate >= " & startCurrent & " AND d.RecordDate < " & FromDate & " THEN (CASE WHEN d.Credit_Or_Debit=0 THEN d.Value ELSE -d.Value END) ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1"
sql = sql & " FROM ACCOUNTS A WITH (NOLOCK)"
sql = sql & " LEFT JOIN dbo.DOUBLE_ENTREY_VOUCHERS  AS d   WITH (NOLOCK)"
sql = sql & "   ON d.Posted IS NULL AND d.Account_Code >= A.Account_Code + 'a' AND d.Account_Code < A.Account_Code + 'b'" & flt_d
sql = sql & " LEFT JOIN dbo.DOUBLE_ENTREY_VOUCHERS1 AS do1 WITH (NOLOCK)"
sql = sql & "   ON do1.Posted IS NULL AND do1.Account_Code >= A.Account_Code + 'a' AND do1.Account_Code < A.Account_Code + 'b'" & flt_do1
sql = sql & " WHERE A.last_account = 0"
If HideLastAccount = True Then sql = sql & " AND A.last_account = 0"
If (TxtAccountCode.text) <> "" Then sql = sql & " AND A.Account_Serial = '" & TxtAccountCode.text & "'"
sql = sql & lvlClause
' ‰›” „‰ÿﬁ  ÷„Ì‰ «·Õ”«»«  (ÊÃÊœ Õ—ﬂ…/«›  «ÕÌ/”‰Ê«  √Ê ·√‰Â √»)
sql = sql & " AND ( EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS  AS e  WITH (NOLOCK) WHERE e.Account_Code >= A.Account_Code + 'a' AND e.Account_Code < A.Account_Code + 'b'" & flt_e & ")"
sql = sql & "   OR EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS1 AS e1 WITH (NOLOCK) WHERE e1.Account_Code >= A.Account_Code + 'a' AND e1.Account_Code < A.Account_Code + 'b'" & flt_e1 & ")"
sql = sql & "   OR EXISTS (SELECT 1 FROM dbo.TblyearsData        AS t  WITH (NOLOCK) WHERE t.Account_Code  = A.Account_Code)"
sql = sql & "   OR A.last_account = 0 )"
sql = sql & " GROUP BY A.last_account, A.ProfitBalance, A.Parent_Account_Code, A.AccountTypes, A.Account_Code, A.Account_Serial, A.Account_Name, A.Account_NameEng"

sql = sql & " ) AS Z"
sql = sql & " GROUP BY last_account, ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng"
sql = sql & " ORDER BY Account_Serial"
sql = sql & " OPTION (RECOMPILE, FAST 1000)"
'================= ‰Â«Ì… »‰«¡ «·«” ⁄·«„ =================



' «»œ√ » ⁄—Ì› „ €Ì— ‰’Ì ›«—€ ··«” ⁄·«„
 sql = ""

' »‰«¡ «·«” ⁄·«„ »«” Œœ«„ CTEs · ÕﬁÌﬁ √ﬁ’Ï ”—⁄… Êœﬁ…
sql = sql & " -- «·ŒÿÊ… 1:  Ã„Ì⁄ „”»ﬁ ·Ãœ«Ê· «·Õ—ﬂ«  «·÷Œ„… „—… Ê«Õœ… ›ﬁÿ" & vbCrLf
sql = sql & " WITH VouchersSummary AS (" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "         d.Account_Code," & vbCrLf
sql = sql & "         SUM(IIF(d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ", d.Value, 0)) AS debitBalance," & vbCrLf
sql = sql & "         SUM(IIF(d.Credit_Or_Debit = 1 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ", d.Value, 0)) AS CreditBalance," & vbCrLf
sql = sql & "         SUM(IIF(d.RecordDate >= " & SQLDate(openingBalanceDate, True) & " AND d.RecordDate <= " & SQLDate(FromdateMinus1, True) & ", IIF(d.Credit_Or_Debit = 0, d.Value, -d.Value), 0)) AS OpeningBalancebeformdateMinus1," & vbCrLf
sql = sql & "         SUM(IIF(d.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND d.RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ", IIF(d.Credit_Or_Debit = 0, d.Value, -d.Value), 0)) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1" & vbCrLf
sql = sql & "     FROM dbo.DOUBLE_ENTREY_VOUCHERS AS d WITH (NOLOCK)" & vbCrLf
sql = sql & "     WHERE d.Posted IS NULL"
' ≈÷«›… ›·« — «·›—Ê⁄ «·œÌ‰«„ÌﬂÌ… Â‰«
If val(DCActivity.BoundText) <> 0 Then
    sql = sql & " AND d.branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    sql = sql & " AND d.branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    sql = sql & " AND d.branch_id = " & val(dcBranch.BoundText)
End If
sql = sql & "     GROUP BY d.Account_Code" & vbCrLf
sql = sql & " )," & vbCrLf

sql = sql & " Vouchers1Summary AS (" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "         do1.Account_Code," & vbCrLf
sql = sql & "         SUM(IIF(do1.Credit_Or_Debit = 0, do1.Value, -do1.Value)) AS OpeningBalance" & vbCrLf
sql = sql & "     FROM dbo.DOUBLE_ENTREY_VOUCHERS1 AS do1 WITH (NOLOCK)" & vbCrLf
sql = sql & "     WHERE do1.Posted IS NULL"
' ≈÷«›… ›·« — «·›—Ê⁄ «·œÌ‰«„ÌﬂÌ… Â‰« √Ì÷«
If val(DCActivity.BoundText) <> 0 Then
    sql = sql & " AND do1.branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    sql = sql & " AND do1.branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    sql = sql & " AND do1.branch_id = " & val(dcBranch.BoundText)
End If
sql = sql & "     GROUP BY do1.Account_Code" & vbCrLf
sql = sql & " )," & vbCrLf

sql = sql & " -- «·ŒÿÊ… 2: ‰⁄«·Ã „‰ÿﬁ OpeningBalance «·Œ«’ Ê«·„Œ ·› ﬂ„« ÂÊ ›Ì «·ﬂÊœ «·√’·Ì" & vbCrLf
sql = sql & " OpeningBalanceCorrected AS (" & vbCrLf
sql = sql & "     SELECT" & vbCrLf
sql = sql & "         A.Account_ID," & vbCrLf
sql = sql & "         SUM(ISNULL(v1s.OpeningBalance, 0)) AS OpeningBalance" & vbCrLf
sql = sql & "     FROM dbo.ACCOUNTS A WITH (NOLOCK)" & vbCrLf
sql = sql & "     LEFT JOIN Vouchers1Summary v1s ON v1s.Account_Code LIKE A.Account_Code + 'a%' OR v1s.Account_Code = A.Account_Code" & vbCrLf
sql = sql & "     GROUP BY A.Account_ID" & vbCrLf
sql = sql & " )" & vbCrLf

sql = sql & " -- «·ŒÿÊ… 3: ‰Ã„⁄ ﬂ· ‘Ì¡ „⁄«" & vbCrLf
sql = sql & " SELECT" & vbCrLf
sql = sql & "     A.last_account, A.ProfitBalance, A.Parent_Account_Code, A.AccountTypes, A.Account_Code, A.Account_Serial, A.Account_Name, A.Account_NameEng," & vbCrLf
sql = sql & "     -1 * F.CreditBalance AS CreditBalance," & vbCrLf
sql = sql & "     A.[Level] AS [Level]," & vbCrLf
sql = sql & "     F.debitBalance," & vbCrLf
sql = sql & "     F.OpeningBalance," & vbCrLf
sql = sql & "     F.OpeningBalancebeformdateMinus1," & vbCrLf
sql = sql & "     F.OpeningBalancebeformStartCurrentyearTOFromDAteminus1" & vbCrLf
sql = sql & " FROM dbo.ACCOUNTS A WITH (NOLOCK)" & vbCrLf
sql = sql & " LEFT JOIN (" & vbCrLf

' -- »‰«¡ «·Ã“¡ «·Œ«’ »‹ UNION ALL --
' «·Ã“¡ «·√Ê·: «·Õ”«»«  «·‰Â«∆Ì… (last_account = 1)
sql = sql & "     SELECT" & vbCrLf
sql = sql & "         A.Account_ID," & vbCrLf
sql = sql & "         ISNULL(vs.debitBalance, 0) AS debitBalance," & vbCrLf
sql = sql & "         ISNULL(vs.CreditBalance, 0) AS CreditBalance," & vbCrLf
sql = sql & "         ISNULL(obc.OpeningBalance, 0) AS OpeningBalance," & vbCrLf
sql = sql & "         ISNULL(vs.OpeningBalancebeformdateMinus1, 0) AS OpeningBalancebeformdateMinus1," & vbCrLf
sql = sql & "         ISNULL(vs.OpeningBalancebeformStartCurrentyearTOFromDAteminus1, 0) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1" & vbCrLf
sql = sql & "     FROM dbo.ACCOUNTS A WITH (NOLOCK)" & vbCrLf
sql = sql & "     LEFT JOIN VouchersSummary vs ON A.Account_Code = vs.Account_Code" & vbCrLf
sql = sql & "     LEFT JOIN OpeningBalanceCorrected obc ON A.Account_ID = obc.Account_ID" & vbCrLf
sql = sql & "     WHERE A.last_account = 1"

sql = sql & "     UNION ALL" & vbCrLf

' «·Ã“¡ «·À«‰Ì: «·Õ”«»«  «·—∆Ì”Ì… (last_account = 0)
sql = sql & "     SELECT" & vbCrLf
sql = sql & "         A.Account_ID," & vbCrLf
sql = sql & "         SUM(ISNULL(vs.debitBalance, 0)) AS debitBalance," & vbCrLf
sql = sql & "         SUM(ISNULL(vs.CreditBalance, 0)) AS CreditBalance," & vbCrLf
sql = sql & "         MAX(ISNULL(obc.OpeningBalance, 0)) AS OpeningBalance," & vbCrLf
sql = sql & "         SUM(ISNULL(vs.OpeningBalancebeformdateMinus1, 0)) AS OpeningBalancebeformdateMinus1," & vbCrLf
sql = sql & "         SUM(ISNULL(vs.OpeningBalancebeformStartCurrentyearTOFromDAteminus1, 0)) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1" & vbCrLf
sql = sql & "     FROM dbo.ACCOUNTS A WITH (NOLOCK)" & vbCrLf
sql = sql & "     LEFT JOIN VouchersSummary vs ON vs.Account_Code LIKE A.Account_Code + 'a%'" & vbCrLf
sql = sql & "     LEFT JOIN OpeningBalanceCorrected obc ON A.Account_ID = obc.Account_ID" & vbCrLf
sql = sql & "     WHERE A.last_account = 0" & vbCrLf
sql = sql & "     GROUP BY A.Account_ID" & vbCrLf
sql = sql & ") AS F ON A.Account_ID = F.Account_ID" & vbCrLf

' -- »‰«¡ Ã„·… WHERE «·‰Â«∆Ì… --
sql = sql & " WHERE 1=1 "
If HideLastAccount = True Then
    sql = sql & " AND A.last_account = 0"
End If
If (TxtAccountCode.text) <> "" Then
    sql = sql & " AND A.Account_Serial = '" & TxtAccountCode.text & "'"
End If
'If account_level <> 0 Then
'    If ShowOnlyLevelAcc = vbYes Then
'        sql = sql & " AND ((LEN(A.account_code) - LEN(REPLACE(A.account_code, 'a', '')) = " & account_level & ")"
'    Else
'        sql = sql & " AND ((LEN(A.account_code) - LEN(REPLACE(A.account_code, 'a', '')) <= " & account_level & ")"
'    End If
'    sql = sql & " OR A.Account_Name LIKE '%?????%')"
'Else
'     sql = sql & " AND ((LEN(A.account_code) - LEN(REPLACE(A.account_code, 'a', '')) <= " & account_level & " ) OR A.Account_Name LIKE '%?????%')"
'End If

If account_level <> 0 Then
    If ShowOnlyLevelAcc = vbYes Then
        sql = sql & " AND (A.[Level] = " & account_level & " "
    Else
        sql = sql & " AND (A.[Level] <= " & account_level & " "
    End If
    sql = sql & " OR A.Account_Name LIKE N'%?????%')"
Else
    sql = sql & " AND (A.[Level] <= " & account_level & " OR A.Account_Name LIKE N'%?????%')"
End If


' ‘—ÿ «·ÊÃÊœ ·· √ﬂœ „‰ √‰ «·Õ”«» ·Â Õ—ﬂ… (»«” Œœ«„ EXISTS «·√”—⁄)
sql = sql & " AND (" & vbCrLf
sql = sql & "     F.debitBalance <> 0 OR F.CreditBalance <> 0 OR F.OpeningBalance <> 0 OR F.OpeningBalancebeformdateMinus1 <> 0 OR F.OpeningBalancebeformStartCurrentyearTOFromDAteminus1 <> 0" & vbCrLf
sql = sql & "     OR EXISTS (SELECT 1 FROM dbo.TblyearsData t WHERE t.Account_Code = A.Account_Code)" & vbCrLf
sql = sql & "     OR A.last_account = 0" & vbCrLf
sql = sql & ")" & vbCrLf

' -- Ã„·… ORDER BY «·‰Â«∆Ì… --
sql = sql & " ORDER BY A.Account_Serial"
'================= «Œ Ì«— „·› «· ﬁ—Ì— =================
If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSa2.rpt"
Else
    StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSaE.rpt"
End If
If Dir(StrFileName) = "" Then
    Screen.MousePointer = vbDefault
    Exit Function
End If

'================= › Õ «·‹Recordset »”—⁄… =================
Cn.CommandTimeout = 120
On Error Resume Next
Cn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
On Error GoTo 0

'Dim RsData As New ADODB.Recordset
With RsData
    .CursorLocation = adUseServer
    .Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
End With

If RsData.BOF Or RsData.EOF Then
    MsgBox IIf(SystemOptions.UserInterface = ArabicInterface, "·« ÌÊÃœ »Ì«‰« ", "No Data"), vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    RsData.Close: Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
End If

Screen.MousePointer = vbArrowHourglass
Set xReport = xApp.OpenReport(StrFileName)
xReport.Database.SetDataSource RsData
'Set RsData.ActiveConnection = Nothing

If RsData.BOF Or RsData.EOF Then
    MsgBox IIf(SystemOptions.UserInterface = ArabicInterface, "?? ???? ??????", "No Data"), vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    RsData.Close: Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
End If

Screen.MousePointer = vbArrowHourglass
Set xReport = xApp.OpenReport(StrFileName)
xReport.Database.SetDataSource RsData
'Set RsData.ActiveConnection = Nothing
'================= ????? =================

    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "?????? " & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "???????" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "?????" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    
    
    
    
    xReport.ParameterFields(7).AddCurrentValue desc
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


Function print_report40New(Optional NoteSerial As String)
On Error GoTo ErrTrap

    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim HideZeroBalance As Integer
    Dim HideLastAccount As Integer
    Dim ShowOnlyLevelAcc As Integer
    Dim openingBalanceDate As Date
    Dim FromdateMinus1 As Date
    Dim StartCurrentDate As Date
    Dim account_level As Long
    Dim X As String
    Dim BranshesReg As String
    Dim BrcnActivety As String
    Dim BranchList As String
    Dim desc As String
    
    FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate

    '================= Level =================
    X = InputBox("Specify Level")
    account_level = val(X)

    '================= Hide Zero =================
    HideZeroBalance = MsgBox("Hide Zero Account", vbInformation + vbYesNoCancel)
    If HideZeroBalance = 2 Then Exit Function

    '================= Hide Last =================
    HideLastAccount = MsgBox("Hide LAst Account", vbInformation + vbYesNoCancel)
    If HideLastAccount = 2 Then Exit Function

    '================= Only Level =================
    ShowOnlyLevelAcc = MsgBox("Show Only This Level ?", vbInformation + vbYesNoCancel)
    If ShowOnlyLevelAcc = 2 Then Exit Function

    '================= Branch Filters =================
    If val(DCRegionID.BoundText) <> 0 Then
        BranshesReg = BranchRegion(DCRegionID.BoundText) ' returns "1,2,3" or "1,2,3" style
    End If
    If val(DCActivity.BoundText) <> 0 Then
        BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
    End If

    ' BranchList priority: Activity then Region (“Ì „‰ÿﬁﬂ ≈‰Â„ »Ì —„Ê« ⁄·Ï ‰›” «·›· —)
    BranchList = ""
    If val(DCActivity.BoundText) <> 0 Then BranchList = BrcnActivety
    If val(DCRegionID.BoundText) <> 0 Then
        If BranchList <> "" Then
            ' ·Ê «·« ‰Ì‰ „ÊÃÊœÌ‰ «‰  ﬂ‰  » ⁄„· AND IN (...) AND IN (...)
            ' ÊœÂ „⁄‰«Â «· ﬁ«ÿ⁄° »” ›Ì VB6 ’⁄» ‰⁄„· Intersection »”—⁄…
            ' ›Â‰« Â‰Œ «— «·√‘œ  ÕœÌœ«: ·Ê ›—⁄ „Õœœ ÂÌ”Êœ
            ' «·√›÷·: ·Ê „Õ «Ã «· ﬁ«ÿ⁄ ﬁÊ·¯Ì Ê‰⁄„·Â SQL-side.
        Else
            BranchList = BranshesReg
        End If
    End If

    '=================  ÕœÌÀ ProfitAccount (“Ì „« ⁄‰œﬂ) =================
    updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg, True

    '================= Build EXEC Stored Procedure =================
    sql = "EXEC dbo.usp_TrialBalanceLevels " & _
          " @FromDate=" & SQLDate(Me.DTPickerAccFrom.value, True) & _
          ",@ToDate=" & SQLDate(Me.DTPickerAccTo.value, True) & _
          ",@OpeningBalanceDate=" & SQLDate(openingBalanceDate, True) & _
          ",@StartCurrentDate=" & SQLDate(StartCurrentDate, True)

    ' BranchID
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & ",@BranchID=" & val(dcBranch.BoundText)
    Else
        sql = sql & ",@BranchID=NULL"
    End If

    ' BranchList CSV
    If Trim$(BranchList) <> "" Then
        sql = sql & ",@BranchList=N'" & Replace(BranchList, "'", "''") & "'"
    Else
        sql = sql & ",@BranchList=NULL"
    End If

    ' Account Serial filter
    If Trim$(TxtAccountCode.text) <> "" Then
        sql = sql & ",@AccountSerial=N'" & Replace(TxtAccountCode.text, "'", "''") & "'"
    Else
        sql = sql & ",@AccountSerial=NULL"
    End If

    ' HideZeroBalance: MsgBox Yes=6
    If HideZeroBalance = 6 Then
        sql = sql & ",@HideZeroBalance=1"
    Else
        sql = sql & ",@HideZeroBalance=0"
    End If

    If HideLastAccount = 6 Then
        sql = sql & ",@HideLastAccount=1"
    Else
        sql = sql & ",@HideLastAccount=0"
    End If

    sql = sql & ",@AccountLevel=" & account_level

    If ShowOnlyLevelAcc = 6 Then
        sql = sql & ",@OnlyLevel=1"
    Else
        sql = sql & ",@OnlyLevel=0"
    End If

    '================= Select report file =================
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSa2.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\TrialBalanceNewSaE.rpt"
    End If
    If Dir(StrFileName) = "" Then Exit Function

    '================= Open RS =================
    Cn.CommandTimeout = 120
    On Error Resume Next
    Cn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
    On Error GoTo ErrTrap

    With RsData
        .CursorLocation = adUseServer
        .Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End With

    If RsData.BOF Or RsData.EOF Then
        MsgBox IIf(SystemOptions.UserInterface = ArabicInterface, "·« ÌÊÃœ »Ì«‰« ", "No Data"), vbExclamation, App.Title
        RsData.Close: Set RsData = Nothing
        Exit Function
    End If

    '================= Crystal =================
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        StrReportTitle = ""
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName
        StrReportTitle = ""
    End If

    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·‰‘«ÿ: " & DCActivity.text & CHR(13)
        Else
            desc = desc & "Activity: " & DCActivity.text & CHR(13)
        End If
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·„‰ÿﬁ…: " & DCRegionID.text & CHR(13)
        Else
            desc = desc & "Region: " & DCRegionID.text & CHR(13)
        End If
    End If
    If val(dcBranch.BoundText) <> 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            desc = desc & "«·›—⁄: " & dcBranch.text & CHR(13)
        Else
            desc = desc & "Branch: " & dcBranch.text & CHR(13)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(6).AddCurrentValue IIf(HideZeroBalance = 6, 1, 0)
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    xReport.ParameterFields(7).AddCurrentValue desc

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

    Exit Function

ErrTrap:
    Screen.MousePointer = vbDefault
End Function
'
'
Function print_report2(Optional NoteSerial As String)
    On Error Resume Next
    On Error GoTo ErrTrap
'print_report2Old
print_report33
Exit Function
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    Dim BranshesReg As String
    Dim BrcnActivety As String
    Dim HideZeroBalance As Integer
    Dim FromdateMinus1 As Date
    Dim StartCurrentDate As Date
    Dim openingBalanceDate As Date

    ' Õ”«» «· Ê«—ÌŒ
    FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate

    ' «·Õ’Ê· ⁄·Ï «·›·« — «·Œ«’… »«·›—⁄ Ê«·‰‘«ÿ
    If val(DCRegionID.BoundText) <> 0 Then
        BranshesReg = BranchRegion(DCRegionID.BoundText)
    End If
    If val(DCActivity.BoundText) <> 0 Then
        BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
    End If

    '  ÕœÌÀ Õ”«» «·√—»«Õ
    updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg

    ' »‰«¡ «·«” ⁄·«„ »«” Œœ«„ CTEs

' »‰«¡ «·«” ⁄·«„ »«” Œœ«„ CTEs
' »‰«¡ «·«” ⁄·«„ »«” Œœ«„ CTEs
' «·Ã“¡ «·√Ê·: »‰«¡ «·›· —
Dim accountSerialFilter As String
If Not mIsFirstYear Then
    accountSerialFilter = " AND do.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
    If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
        accountSerialFilter = accountSerialFilter & " AND do.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
    End If
End If

' «·Ã“¡ «·À«‰Ì: œ„Ã «·›· — ›Ì «·«” ⁄·«„« 


Dim FromDate As String
Dim ToDate As String


'  ⁄ÌÌ‰ «·»—«„ —«  «·Œ«’… »«· Ê«—ÌŒ
FromDate = SQLDate(Me.DTPickerAccFrom.value, True)
ToDate = SQLDate(Me.DTPickerAccTo.value, True)
'openingBalanceDate = SQLDate(DateSerial(year(Me.DTPickerAccFrom.value), 1, 1), True)

' «·Ã“¡ «·√Ê·: VoucherData
sql = "WITH VoucherData AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) AS DebitBalance, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value * -1 ELSE 0 END) AS CreditBalance " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= " & FromDate & ") AND (RecordDate <= " & ToDate & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' «·Ã“¡ «·À«‰Ì: OpeningBalanceData
sql = sql & "OpeningBalanceData AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalance " & _
      " FROM DOUBLE_ENTREY_VOUCHERS1 " & _
      " WHERE Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' «·Ã“¡ «·À«·À: OpeningBalanceBeforeDateMinus1
sql = sql & "OpeningBalanceBeforeDateMinus1 AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformdateMinus1 " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= " & openingBalanceDate & ") AND (RecordDate <= " & FromDate & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' «·Ã“¡ «·—«»⁄: OpeningBalanceBeforeStartCurrentYear
sql = sql & "OpeningBalanceBeforeStartCurrentYear AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= " & openingBalanceDate & ") " & _
      "   AND (RecordDate < " & FromDate & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code) "

' «·Ã“¡ «·Œ«„”: «·«” ⁄·«„ «·—∆Ì”Ì
sql = sql & "SELECT A.Account_Code, A.Account_Name, A.Account_NameEng, A.Account_Serial, " & _
      "       A.AccountTypes, A.Parent_Account_Code, A.ProfitBalance, A.last_account, " & _
      "       VD.DebitBalance, VD.CreditBalance, OBD.OpeningBalance, " & _
      "       OBMD1.OpeningBalancebeformdateMinus1, " & _
      "       OBSCY.OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
      "FROM ACCOUNTS AS A " & _
      "LEFT OUTER JOIN VoucherData AS VD ON A.Account_Code = VD.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceData AS OBD ON A.Account_Code = OBD.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceBeforeDateMinus1 AS OBMD1 ON A.Account_Code = OBMD1.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceBeforeStartCurrentYear AS OBSCY ON A.Account_Code = OBSCY.Account_Code " & _
      "WHERE (A.last_account = 1) "

' ≈÷«›… «·›·« — «·œÌ‰«„ÌﬂÌ…
sql = sql & " OR A.Account_Code IN (SELECT Account_Code FROM DOUBLE_ENTREY_VOUCHERS1 WHERE 1 = 1 "
If val(DCActivity.BoundText) <> 0 Then
    sql = sql & " AND branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    sql = sql & " AND branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    sql = sql & " AND branch_id = " & val(dcBranch.BoundText)
End If
sql = sql & ") "

sql = sql & " OR A.Account_Code IN (SELECT Account_Code FROM TblyearsData WHERE 1 = 1 "
If val(DCActivity.BoundText) <> 0 Then
    sql = sql & " AND branch_id IN (" & BrcnActivety & ")"
End If
If val(DCRegionID.BoundText) <> 0 Then
    sql = sql & " AND branch_id IN (" & BranshesReg & ")"
End If
If val(dcBranch.BoundText) <> 0 Then
    sql = sql & " AND branch_id = " & val(dcBranch.BoundText)
End If
sql = sql & ") "

sql = sql & " ORDER BY A.Account_Serial"




'Dim toDate As String

' ????? ?????????? ?????? ?????????
FromDate = SQLDate(Me.DTPickerAccFrom.value, True)
ToDate = SQLDate(Me.DTPickerAccTo.value, True)

' ????? ?????: VoucherData
sql = "WITH VoucherData AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) AS DebitBalance, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value * -1 ELSE 0 END) AS CreditBalance " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= " & FromDate & ") " & _
      "   AND (RecordDate <= " & ToDate & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' ????? ??????: OpeningBalanceData
sql = sql & "OpeningBalanceData AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalance " & _
      " FROM DOUBLE_ENTREY_VOUCHERS1 " & _
      " WHERE Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' ????? ??????: OpeningBalanceBeforeDateMinus1
sql = sql & "OpeningBalanceBeforeDateMinus1 AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformdateMinus1 " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= '01-Jan-2023') AND (RecordDate <= " & SQLDate(DateAdd("d", -1, Me.DTPickerAccFrom.value), True) & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code), "

' ????? ??????: OpeningBalanceBeforeStartCurrentYear
sql = sql & "OpeningBalanceBeforeStartCurrentYear AS (" & _
      " SELECT Account_Code, " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
      "        SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
      " FROM DOUBLE_ENTREY_VOUCHERS " & _
      " WHERE (RecordDate >= " & SQLDate(DateSerial(year(Me.DTPickerAccFrom.value), 1, 1), True) & ") " & _
      "   AND (RecordDate < " & FromDate & ") " & _
      "   AND Account_Code IN (" & _
      "       SELECT a.Account_Code FROM ACCOUNTS a " & _
      "       WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 " & _
      "          OR SUBSTRING(a.Account_Serial, 1, 1) < 3)" & _
      " GROUP BY Account_Code) "

' ????? ??????: ????????? ???????
sql = sql & "SELECT A.Account_Code, A.Account_Name, A.Account_NameEng, A.Account_Serial, " & _
      "       A.AccountTypes, A.Parent_Account_Code, A.ProfitBalance, A.last_account, " & _
      "       VD.DebitBalance, VD.CreditBalance, OBD.OpeningBalance, " & _
      "       OBMD1.OpeningBalancebeformdateMinus1, " & _
      "       OBSCY.OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
      "FROM ACCOUNTS AS A " & _
      "LEFT OUTER JOIN VoucherData AS VD ON A.Account_Code = VD.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceData AS OBD ON A.Account_Code = OBD.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceBeforeDateMinus1 AS OBMD1 ON A.Account_Code = OBMD1.Account_Code " & _
      "LEFT OUTER JOIN OpeningBalanceBeforeStartCurrentYear AS OBSCY ON A.Account_Code = OBSCY.Account_Code " & _
      "WHERE (A.last_account = 1) "
    If (TxtAccountCode.text) <> "" Then
        sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
    End If
  
    sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
    sql = sql & "    Where 1 = 1"
    
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
   sql = sql & "   )"
   
    sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
  sql = sql & "   )"
  
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData"
  sql = sql & "    Where 1 = 1"
  ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
  ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

  sql = sql & "   ))"
      
    sql = sql & "ORDER BY A.Account_Serial"

    '  ‰›Ì– «·«” ⁄·«„ Ê⁄—÷ «· ﬁ—Ì—
    Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = IIf(SystemOptions.UserInterface = ArabicInterface, "·« ÌÊÃœ »Ì«‰« ", "No Data")
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '  Õ„Ì· «· ﬁ—Ì—
    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & IIf(SystemOptions.UserInterface = ArabicInterface, "TrialBalanceNewSa.rpt", "TrialBalanceNewFAHDE.rpt")
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Dim cCompanyInfo As New ClsCompanyInfo
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    '  ⁄ÌÌ‰ „⁄·„«  «· ﬁ—Ì—
    xReport.ParameterFields(1).AddCurrentValue IIf(SystemOptions.UserInterface = ArabicInterface, cCompanyInfo.ArabCompanyName, cCompanyInfo.EngCompanyName)
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    xReport.ParameterFields(7).AddCurrentValue ""
 If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    ' ⁄—÷ «· ﬁ—Ì—
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

ErrTrap:
    Exit Function
End Function
Function print_report33(Optional NoteSerial As String)
'   print_report2Old
    On Error Resume Next
    On Error GoTo ErrTrap

    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    Dim BranshesReg As String
    Dim BrcnActivety As String
    Dim HideZeroBalance As Integer
    Dim FromdateMinus1 As Date
    Dim StartCurrentDate As Date
    Dim openingBalanceDate As Date
    HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
    ' ????? ????????
    FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate

    ' ????? ??????? ????? ??? ?????? ????????
    If val(DCRegionID.BoundText) <> 0 Then
        BranshesReg = BranchRegion(DCRegionID.BoundText)
    End If
    If val(DCActivity.BoundText) <> 0 Then
        BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
    End If

    ' ????? ??? ????????
    updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg

    ' ????????? ???????
       
       
Dim part1 As String
Dim Part2 As String
Dim part3 As String
Dim part4 As String
Dim part5 As String
Dim employeePart As String


' «·Ã“¡ «·√Ê·
Dim s As String
s = "Select * from TblyearsData  where IsNull(IsFirstYear,0) = 1 and YEAR(datesatrt)  = " & year(DTPickerAccFrom.value)
Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If rsDummy.EOF Then
    mIsFirstYear = False
Else
    mIsFirstYear = True
End If
Dim accountSerialFilter As String
If Not mIsFirstYear Then
   ' accountSerialFilter = " AND do.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
    If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
        accountSerialFilter = accountSerialFilter & " AND do.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) <3) "
    End If
End If

Dim AqarFilterV As String
Dim AqarFilterV1 As String

AqarFilterV = ""
AqarFilterV1 = ""

If val(DcbAqar.BoundText) <> 0 Then
    ' ··‹ DOUBLE_ENTREY_VOUCHERS
    AqarFilterV = " AND (ISNULL(DOUBLE_ENTREY_VOUCHERS.Aqarid,0) = " & val(DcbAqar.BoundText) & _
                  " OR ISNULL(DOUBLE_ENTREY_VOUCHERS.iqarid,0) = " & val(DcbAqar.BoundText) & ") "

    ' ··‹ DOUBLE_ENTREY_VOUCHERS1 (·Ê ›ÌÂ ‰›” «·ÕﬁÊ·)
    AqarFilterV1 = " and DOUBLE_ENTREY_VOUCHERS1.Notes_id = -99"
End If



part1 = "WITH VoucherData AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) AS DebitBalance, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value * -1 ELSE 0 END) AS CreditBalance " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and (RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & ") AND (RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
          If val(DCActivity.BoundText) <> 0 Then
                part1 = part1 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            part1 = part1 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            part1 = part1 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        
        part1 = part1 & AqarFilterV

        part1 = part1 & "   "
        part1 = part1 & "    GROUP BY Account_Code), "

' «·Ã“¡ «·À«‰Ì
Part2 = " OpeningBalanceData AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalance " & _
        " FROM DOUBLE_ENTREY_VOUCHERS1 where (Posted IS NULL) and  1 = 1 "
        
        If val(DCActivity.BoundText) <> 0 Then
                Part2 = Part2 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            Part2 = Part2 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            Part2 = Part2 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        If Not mIsFirstYear Then
            'Part2 = Part2 & "   AND DOUBLE_ENTREY_VOUCHERS1.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
            Part2 = Part2 & "   AND DOUBLE_ENTREY_VOUCHERS1.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE  SUBSTRING(a.Account_Serial, 1, 1) <3) "
        End If
        Part2 = Part2 & "   "
        Part2 = Part2 & "    GROUP BY Account_Code), "


' «·Ã“¡ «·À«·À
part3 = " OpeningBalanceBeforeDateMinus1 AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformdateMinus1 " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and  (RecordDate >= " & SQLDate(openingBalanceDate, True) & ") AND (RecordDate <= " & SQLDate(FromdateMinus1, True) & ") "
        
        If val(DCActivity.BoundText) <> 0 Then
                part3 = part3 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            part3 = part3 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            part3 = part3 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        part3 = part3 & AqarFilterV

        If Not mIsFirstYear Then
            If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
                'part3 = part3 & "   AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
                part3 = part3 & "   AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE   SUBSTRING(a.Account_Serial, 1, 1) <3) "
            End If
        End If
        part3 = part3 & "   "
        part3 = part3 & "    GROUP BY Account_Code), "


' «·Ã“¡ «·—«»⁄
part4 = " OpeningBalanceBeforeStartCurrentYear AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and  ( RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ") "
        '" WHERE (Posted IS NULL) and  (RecordDate >= " & SQLDate(StartCurrentDate, True) & ") AND (RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ") "
        
                  If val(DCActivity.BoundText) <> 0 Then
                part4 = part4 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            part4 = part4 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            part4 = part4 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        part4 = part4 & AqarFilterV

        If Not mIsFirstYear Then
            If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
                'part4 = part4 & "   AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
                part4 = part4 & "   AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE   SUBSTRING(a.Account_Serial, 1, 1) <3) "
            End If
        End If
        part4 = part4 & "   "
        part4 = part4 & "    GROUP BY Account_Code) "


' «·Ã“¡ «·Œ«„” («·«” ⁄·«„ «·—∆Ì”Ì)
part5 = " SELECT A.Account_Code, A.Account_Name, A.Account_NameEng, A.Account_Serial, " & _
        " A.AccountTypes, A.Parent_Account_Code, A.ProfitBalance, A.last_account, " & _
        " VD.DebitBalance, VD.CreditBalance, OBD.OpeningBalance, " & _
        " OBMD1.OpeningBalancebeformdateMinus1, OBSCY.OpeningBalancebeformStartCurrentyearTOFromDAteminus1, " & _
        " EmployeeData.Emp_Code, EmployeeData.Emp_Name, EmployeeData.GroupName " & _
        " FROM ACCOUNTS AS A " & _
        " LEFT OUTER JOIN VoucherData AS VD ON A.Account_Code = VD.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceData AS OBD ON A.Account_Code = OBD.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceBeforeDateMinus1 AS OBMD1 ON A.Account_Code = OBMD1.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceBeforeStartCurrentYear AS OBSCY ON A.Account_Code = OBSCY.Account_Code "

' «·Ã“¡ «·Œ«’ »Ãœ«Ê· «·„ÊŸ›Ì‰
employeeUnion = " SELECT Account_code1 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code3 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID "

employeeUnion = employeeUnion & _
                " UNION ALL " & _
                " SELECT Account_code4 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code5 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code2 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID "

employeeJoin = " LEFT OUTER JOIN (" & _
               " SELECT DISTINCT Account_Code, Emp_Code, Emp_Name, GroupName, " & _
               " ROW_NUMBER() OVER (PARTITION BY Account_Code ORDER BY Emp_Code) AS RowNum " & _
               " FROM (" & employeeUnion & _
               " ) AS EmployeeAccounts " & _
               " ) EmployeeData ON A.Account_Code = EmployeeData.Account_Code AND EmployeeData.RowNum = 1 "

' «· Ã„Ì⁄ «·‰Â«∆Ì
sql = part1 & Part2 & part3 & part4 & part5 & employeeJoin & _
      " WHERE (A.last_account = 1) "
   If (TxtAccountCode.text) <> "" Then
        sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
    End If
  
    sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
    sql = sql & "    Where 1 = 1"
    
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
   sql = sql & "   )"
   
    sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
  sql = sql & "   )"
  
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData"
  sql = sql & "    Where 1 = 1"
  ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
  ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

  sql = sql & "   ))"
            
      sql = sql & "    ORDER BY A.Account_Serial"

    ' ????? ???????
    
    
    
    ' ===================================================================================
'  „  ⁄œÌ· Â–« «·”ﬂ—Ì»  · ’ÕÌÕ ‘—ÿ ›· —… «·Õ”«»«  »‰«¡ ⁄·Ï ÿ·»ﬂ
' ===================================================================================

'  ÕœÌœ ≈–« ﬂ«‰  «·”‰… «·„«·Ì… ÂÌ «·√Ê·Ï √„ ·«
'Dim s As String
s = "Select * from TblyearsData where IsNull(IsFirstYear,0) = 1 and YEAR(datesatrt) = " & year(DTPickerAccFrom.value)
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If rsDummy.EOF Then
    mIsFirstYear = False
Else
    mIsFirstYear = True
End If

'  „ Õ–› „ €Ì— accountSerialFilter ·√‰Â ·„ Ìﬂ‰ „” Œœ„« Ê»”»» «·√Œÿ«¡ «·„‰ÿﬁÌ… ›ÌÂ

' «·Ã“¡ «·√Ê·: »Ì«‰«  «·Õ—ﬂ«  Œ·«· «·› —… «·„Õœœ… (»œÊ‰  €ÌÌ—)
part1 = "WITH VoucherData AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) AS DebitBalance, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value * -1 ELSE 0 END) AS CreditBalance " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and (RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & ") AND (RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") "
        If val(DCActivity.BoundText) <> 0 Then
                part1 = part1 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            part1 = part1 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            part1 = part1 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        part1 = part1 & AqarFilterV

        part1 = part1 & " GROUP BY Account_Code), "

' «·Ã“¡ «·À«‰Ì: «·√—’œ… «·«›  «ÕÌ… „‰ ÃœÊ· «·√—’œ…
' ***  „  ’ÕÌÕ «·‘—ÿ Â‰« ***
Part2 = " OpeningBalanceData AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalance " & _
        " FROM DOUBLE_ENTREY_VOUCHERS1 where (Posted IS NULL) and 1 = 1 "
        
        If val(DCActivity.BoundText) <> 0 Then
                Part2 = Part2 & " and branch_id in (" & BrcnActivety & ")"
        End If
        Part2 = Part2 & AqarFilterV1

        If val(DCRegionID.BoundText) <> 0 Then
            Part2 = Part2 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            Part2 = Part2 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If
        
        ' -->> »œ«Ì… «· ⁄œÌ·:  ÿ»Ìﬁ «·‘—ÿ «·’ÕÌÕ ﬂ„« ›Ì «·”ﬂ—Ì»  «·À«‰Ì
'        If Not mIsFirstYear Then
'            Part2 = Part2 & " AND DOUBLE_ENTREY_VOUCHERS1.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
'        End If
        
        If Not mIsFirstYear Then
        'Wael
            'Part2 = Part2 & " AND DOUBLE_ENTREY_VOUCHERS1.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) IN ('1','2','3')) "
                Part2 = Part2 & " AND EXISTS ( " & _
               "   SELECT 1 FROM ACCOUNTS a " & _
               "   WHERE a.Account_Code = DOUBLE_ENTREY_VOUCHERS1.Account_Code " & _
               "     AND (a.Account_Serial LIKE '1%' OR a.Account_Serial LIKE '2%' OR a.AccountTypes = 1) " & _
               ") "

        End If
        ' -->> ‰Â«Ì… «· ⁄œÌ·
        
        Part2 = Part2 & " GROUP BY Account_Code), "


' «·Ã“¡ «·À«·À: —’Ìœ «·Õ—ﬂ«  „‰ »œ«Ì… «·”‰… Õ Ï  «—ÌŒ »œ«Ì… «· ﬁ—Ì— (‰«ﬁ’ ÌÊ„ Ê«Õœ)
' ***  „  ’ÕÌÕ «·‘—ÿ Â‰« ***
part3 = " OpeningBalanceBeforeDateMinus1 AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformdateMinus1 " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and (RecordDate >= " & SQLDate(openingBalanceDate, True) & ") AND (RecordDate <= " & SQLDate(FromdateMinus1, True) & ") "
        
        If val(DCActivity.BoundText) <> 0 Then
                part3 = part3 & " and branch_id in (" & BrcnActivety & ")"
        End If
        If val(DCRegionID.BoundText) <> 0 Then
            part3 = part3 & " and branch_id in (" & BranshesReg & ")"
        End If
        part3 = part3 & AqarFilterV

        If val(dcBranch.BoundText) <> 0 Then
            part3 = part3 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If

        ' -->> »œ«Ì… «· ⁄œÌ·:  ÿ»Ìﬁ «·‘—ÿ «·’ÕÌÕ Ê≈“«·… «·«⁄ „«œ ⁄·Ï  «—ÌŒ 1 Ì‰«Ì—
        If Not mIsFirstYear Then
            'part3 = part3 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
        End If
        
        If Not mIsFirstYear Then
            'part3 = part3 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) IN ('1','2','3')) "
            'Wael
            'part3 = part3 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) IN ('1','2')) "
            part3 = part3 & " AND EXISTS ( " & _
               "   SELECT 1 FROM ACCOUNTS a " & _
               "   WHERE a.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code " & _
               "     AND (a.Account_Serial LIKE '1%' OR a.Account_Serial LIKE '2%' OR a.AccountTypes = 1) " & _
               ") "

        End If
        ' -->> ‰Â«Ì… «· ⁄œÌ·
        
        part3 = part3 & " GROUP BY Account_Code), "


' «·Ã“¡ «·—«»⁄: —’Ìœ «·Õ—ﬂ«  „‰ »œ«Ì… «·”‰… «·Õ«·Ì… Õ Ï „« ﬁ»·  «—ÌŒ »œ«Ì… «· ﬁ—Ì—
' ***  „  ’ÕÌÕ «·‘—ÿ Â‰« ***
part4 = " OpeningBalanceBeforeStartCurrentYear AS (" & _
        " SELECT Account_Code, " & _
        " SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) - " & _
        " SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS OpeningBalancebeformStartCurrentyearTOFromDAteminus1 " & _
        " FROM DOUBLE_ENTREY_VOUCHERS " & _
        " WHERE (Posted IS NULL) and ( RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ") "
        
        If val(DCActivity.BoundText) <> 0 Then
            part4 = part4 & " and branch_id in (" & BrcnActivety & ")"
        End If
        part4 = part4 & AqarFilterV

        If val(DCRegionID.BoundText) <> 0 Then
            part4 = part4 & " and branch_id in (" & BranshesReg & ")"
        End If
        If val(dcBranch.BoundText) <> 0 Then
            part4 = part4 & " and branch_id =" & val(dcBranch.BoundText) & ""
        End If

        ' -->> »œ«Ì… «· ⁄œÌ·:  ÿ»Ìﬁ «·‘—ÿ «·’ÕÌÕ Ê≈“«·… «·«⁄ „«œ ⁄·Ï  «—ÌŒ 1 Ì‰«Ì—
        If Not mIsFirstYear Then
            'part4 = part4 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) > 4 OR SUBSTRING(a.Account_Serial, 1, 1) < 3) "
        End If
        If Not mIsFirstYear Then
            'part4 = part4 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) IN ('1','2','3')) "
            'Wael
            'part4 = part4 & " AND DOUBLE_ENTREY_VOUCHERS.Account_Code IN (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial, 1, 1) IN ('1','2')) "
            
            part4 = part4 & " AND EXISTS ( " & _
               "   SELECT 1 FROM ACCOUNTS a " & _
               "   WHERE a.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code " & _
               "     AND (a.Account_Serial LIKE '1%' OR a.Account_Serial LIKE '2%' OR a.AccountTypes = 1) " & _
               ") "

        End If
        ' -->> ‰Â«Ì… «· ⁄œÌ·
        
        part4 = part4 & " GROUP BY Account_Code) "


' «·Ã“¡ «·Œ«„” («·«” ⁄·«„ «·—∆Ì”Ì) Ê Ã„Ì⁄ »«ﬁÌ «·√Ã“«¡ (»œÊ‰  €ÌÌ—)
part5 = " SELECT A.Account_Code, A.Account_Name, A.Account_NameEng, A.Account_Serial, " & _
        " A.AccountTypes, A.Parent_Account_Code, A.ProfitBalance, A.last_account, " & _
        " VD.DebitBalance, VD.CreditBalance, OBD.OpeningBalance, " & _
        " OBMD1.OpeningBalancebeformdateMinus1, OBSCY.OpeningBalancebeformStartCurrentyearTOFromDAteminus1, " & _
        " EmployeeData.Emp_Code, EmployeeData.Emp_Name, EmployeeData.GroupName " & _
        " FROM ACCOUNTS AS A " & _
        " LEFT OUTER JOIN VoucherData AS VD ON A.Account_Code = VD.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceData AS OBD ON A.Account_Code = OBD.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceBeforeDateMinus1 AS OBMD1 ON A.Account_Code = OBMD1.Account_Code " & _
        " LEFT OUTER JOIN OpeningBalanceBeforeStartCurrentYear AS OBSCY ON A.Account_Code = OBSCY.Account_Code "

employeeUnion = " SELECT Account_code1 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code3 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID "

employeeUnion = employeeUnion & _
                " UNION ALL " & _
                " SELECT Account_code4 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code5 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID " & _
                " UNION ALL " & _
                " SELECT Account_code2 AS Account_Code, Emp_Code, Emp_Name, GroupName " & _
                " FROM TblEmployee " & _
                " LEFT JOIN EmpGroupDep ON TblEmployee.GroupID = EmpGroupDep.GroupID "

employeeJoin = " LEFT OUTER JOIN (" & _
               " SELECT DISTINCT Account_Code, Emp_Code, Emp_Name, GroupName, " & _
               " ROW_NUMBER() OVER (PARTITION BY Account_Code ORDER BY Emp_Code) AS RowNum " & _
               " FROM (" & employeeUnion & _
               " ) AS EmployeeAccounts " & _
               " ) EmployeeData ON A.Account_Code = EmployeeData.Account_Code AND EmployeeData.RowNum = 1 "

' «· Ã„Ì⁄ «·‰Â«∆Ì (»œÊ‰  €ÌÌ—)
sql = part1 & Part2 & part3 & part4 & part5 & employeeJoin & _
      " WHERE (A.last_account = 1) "
    If (TxtAccountCode.text) <> "" Then
        sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
    End If
   
    sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
    sql = sql & "   Where 1 = 1"
    
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
    sql = sql & AqarFilterV
   sql = sql & "   )"
   
    sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "   Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BrcnActivety & ")"
    End If
    sql = sql & AqarFilterV1

    If val(DCRegionID.BoundText) <> 0 Then
        sql = sql & " and branch_id in (" & BranshesReg & ")"
    End If
    

    If val(dcBranch.BoundText) <> 0 Then
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
    End If
  sql = sql & "   )"
 
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData"
  sql = sql & "   Where 1 = 1"
  ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
  ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

  sql = sql & "   ))"
               
      sql = sql & "   ORDER BY A.Account_Serial"
    Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = IIf(SystemOptions.UserInterface = ArabicInterface, "·« ÌÊÃœ œ« «", "No Data")
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ' ????? ??? ???????
    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & IIf(SystemOptions.UserInterface = ArabicInterface, "TrialBalanceNewSa.rpt", "TrialBalanceNewFAHDE.rpt")

    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Dim cCompanyInfo As New ClsCompanyInfo
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    ' ????? ???????? ???????
    xReport.ParameterFields(1).AddCurrentValue IIf(SystemOptions.UserInterface = ArabicInterface, cCompanyInfo.ArabCompanyName, cCompanyInfo.EngCompanyName)
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
     If HideZeroBalance = 6 Then
        xReport.ParameterFields(6).AddCurrentValue 1
    Else
        xReport.ParameterFields(6).AddCurrentValue 0
    End If
    
    xReport.ParameterFields(7).AddCurrentValue ""






'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        StrReportTitle = "" '& StrAccountName
'        Else
'         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
'        StrReportTitle = ""
'    End If
    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "???????" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "?????" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
'    If HideZeroBalance = 6 Then
'    xReport.ParameterFields(6).AddCurrentValue 1
'    Else
'    xReport.ParameterFields(6).AddCurrentValue 0
'    End If
  
    'xReport.ParameterFields(7).AddCurrentValue desc
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    ' ??? ???????
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

ErrTrap:
    Exit Function
End Function

Function print_report2Old(Optional NoteSerial As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer
  Dim mIsFirstYear As Boolean
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
 
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim BrcnActivety As String
   FromdateMinus1 = DateAdd("d", -1, DTPickerAccFrom.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  HideZeroBalance = 7
'         If SystemOptions.UserInterface = ArabicInterface Then
'                HideZeroBalance = MsgBox("?? ???? ????? ?????? ????? ??? ?? ?? ", vbInformation + vbYesNoCancel)
'            Else
'                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
'            End If
'
'            If HideZeroBalance = 2 Then
'                Screen.MousePointer = vbDefault
'                Exit Function
'            End If
          Dim BranshesReg As String
      
         If val(DCRegionID.BoundText) <> 0 Then
         BranshesReg = BranchRegion(DCRegionID.BoundText)
         End If
         If val(DCActivity.BoundText) <> 0 Then
         BrcnActivety = BrcnhActivityType(DCActivity.BoundText)
         End If



Dim s As String
s = "Select * from TblyearsData  where IsNull(IsFirstYear,0) = 1 and YEAR(datesatrt)  = " & year(DTPickerAccFrom.value)
Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If rsDummy.EOF Then
    mIsFirstYear = False
Else
    mIsFirstYear = True
End If


  updateprofitAccount val(DCActivity.BoundText), val(dcBranch.BoundText), Me.DTPickerAccTo.value, BranshesReg

  sql = " SELECT   TblEmployee.Emp_Code,TblEmployee.Emp_Name,EmpGroupDep.GroupName,last_account,   ProfitBalance, Parent_Account_Code, AccountTypes, a.Account_Code, Account_Serial, Account_Name,Account_NameEng, debitBalance ="
  sql = sql & "                         (SELECT     SUM(DEV_Value1)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & " AND d.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") AND d.Account_Code = A.Account_Code  and(d.Posted IS NULL)"
 If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
  End If
  'sql = sql & " and  d.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d.branch_id in (" & BranshesReg & ")"
  End If
  
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d.branch_id =" & val(dcBranch.BoundText) & ""
  End If
 sql = sql & "  ) x),"
   sql = sql & "                    CreditBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(Me.DTPickerAccFrom.value, True) & "  AND d1.RecordDate <= " & SQLDate(Me.DTPickerAccTo.value, True) & ") AND d1.Account_Code = A.Account_Code and(d1.Posted IS NULL)"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
  End If
 If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and d1.branch_id =" & val(dcBranch.BoundText) & ""
  End If
 ' sql = sql & " and  d1.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  sql = sql & " ) x),"
  
  
  sql = sql & "                     OpeningBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  sql = sql & "                                                   WHERE     (  do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
    If Not mIsFirstYear Then
        If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
       ' sql = sql & " and do.Account_Code  in  (SELECT a.Account_Code FROM ACCOUNTS a WHERE (SUBSTRING(a.Account_Serial , 1, 1) > 4  or SUBSTRING(a.Account_Serial , 1, 1) <3) )                                                  "
        End If
        sql = sql & " and do.Account_Code  in  (SELECT a.Account_Code FROM ACCOUNTS a WHERE (SUBSTRING(a.Account_Serial , 1, 1) > 4  or SUBSTRING(a.Account_Serial , 1, 1) <3) )                                                  "
    End If
    'sql = sql & " and  do.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
 sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
 End If
sql = sql & "  )) x),"
  sql = sql & "    OpeningBalancebeformdateMinus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     do.Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   INNER JOIN ACCOUNTS aa ON do.Account_Code = Aa.Account_Code"
  sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
  
  If Not mIsFirstYear Then
    sql = sql & " and do.Account_Code   In (SELECT a.Account_Code FROM ACCOUNTS a WHERE (SUBSTRING(a.Account_Serial , 1, 1)  > 4 or SUBSTRING(a.Account_Serial , 1, 1)  < 3   )  )   "
    If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
    '   sql = sql & " and do.Account_Code   In (SELECT a.Account_Code FROM ACCOUNTS a WHERE (SUBSTRING(a.Account_Serial , 1, 1)  > 4 or SUBSTRING(a.Account_Serial , 1, 1)  < 3   )  )   "
    Else
        'sql = sql & "  AND  yeAR(do.RecordDate) >= CASE  (SUBSTRING(aa.Account_Serial , 1, 1)  ) WHEN 4   THEN " & year(DTPickerAccFrom.value) & "  WHEN 3 THEN " & year(DTPickerAccFrom.value) & " ELSE  1900 END"
        'sql = sql & " and do.Account_Code   In (SELECT a.Account_Code FROM ACCOUNTS a WHERE (SUBSTRING(a.Account_Serial , 1, 1)  > 4 or SUBSTRING(a.Account_Serial , 1, 1)  < 3   )  )   "

   End If
  End If
  'sql = sql & " and  do.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     do.Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   INNER JOIN ACCOUNTS aa ON do.Account_Code = Aa.Account_Code"
  sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(Me.DTPickerAccFrom.value, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL) "
  
    If Not mIsFirstYear Then
        sql = sql & " and do.Account_Code   In (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial , 1, 1) >4 or SUBSTRING(a.Account_Serial , 1, 1) <3 )                                                  "
        If Month(DTPickerAccFrom.value) = 1 And day(DTPickerAccFrom.value) = 1 Then
            sql = sql & " and do.Account_Code   In (SELECT a.Account_Code FROM ACCOUNTS a WHERE SUBSTRING(a.Account_Serial , 1, 1) >4 or SUBSTRING(a.Account_Serial , 1, 1) <3 )                                                  "
        Else
       '     sql = sql & "  AND  yeAR(do.RecordDate) >= CASE  (SUBSTRING(aa.Account_Serial , 1, 1)  ) WHEN 4   THEN " & year(DTPickerAccFrom.value) & "  WHEN 3 THEN " & year(DTPickerAccFrom.value) & " ELSE  1900 END"
        End If
  End If

  If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and do.branch_id =" & val(dcBranch.BoundText) & ""
  End If
  
  'sql = sql & " and  do.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  
  sql = sql & " ) x)"
  sql = sql & " FROM         ACCOUNTS A"
  sql = sql & " Left outer join TblEmployee"
    sql = sql & " ON a.Account_Code in (TblEmployee.Account_code1,TblEmployee.Account_code,TblEmployee.Account_code3,TblEmployee.Account_Code4,TblEmployee.Account_code5, TblEmployee.Account_Code2)"
    sql = sql & " LEFT OUTER JOIN                      dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID"
  sql = sql & " WHERE     A.last_account = 1   "
  If (TxtAccountCode.text) <> "" Then
  sql = sql & " and A.Account_Serial ='" & TxtAccountCode.text & "'"
  End If
  
 ' sql = sql & " and  A.Account_Code not in ( Select tt.account_code from accounts tt where tt.Parent_Account_Code in(Select accountcode from AccountSetting))"
  sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
  sql = sql & "    Where 1 = 1"
  
    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
   sql = sql & "   )"
   
  sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"
    If val(DCActivity.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(DCRegionID.BoundText) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(dcBranch.BoundText) <> 0 Then
  sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
  End If
  sql = sql & "   )"
  
  sql = sql & " or A.Account_Code in(select Account_Code from  TblyearsData"
 sql = sql & "    Where 1 = 1"
 ' sql = sql & " and (OpeneingbalancesDate = " & SQLDate(StartCurrentDate, True) & ")"
 ' AND OpeneingbalancesDate<= " & SQLDate(Me.DTPickerAccFrom.value, True) & ")"

   sql = sql & "   ))"
   
  'sql = sql & " and (AccountTypes = 1 OR AccountTypes = 0)"
    sql = sql & "order by Account_Serial "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSa.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewFAHDE.rpt"
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
       Msg = "·« ÌÊÃœ »Ì«‰« "
     Else
     Msg = "No Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim desc As String
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Dim cCompanyInfo As New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    desc = ""
    If val(DCActivity.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   Else
   desc = desc & "Region" & ": " & DCActivity.text & CHR(13)
   End If
   End If
   
   If val(DCRegionID.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "???????" & ": " & DCRegionID.text & CHR(13)
   Else
   desc = desc & "Activity" & ": " & DCRegionID.text & CHR(13)
   End If
   End If
  If val(dcBranch.BoundText) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "?????" & ": " & dcBranch.text & CHR(13)
   Else
   desc = desc & "Branch" & ": " & dcBranch.text & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
    xReport.ParameterFields(7).AddCurrentValue desc
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


Private Sub OptAccount_Click(Index As Integer)
'    Frame5.Visible = False
'    Frame5.Visible = False
'    Frame2.Visible = False
   ' Frame1.Visible = False
   ' Frame3.Visible = False
   ' DCEmployee.text = ""
   ' Frame7.Visible = False
   ' Frame6.Visible = False
   ' DCCompositeAccount.Enabled = False
   ' DCCompositeAccount.text = ""
Me.Ele(1).Visible = True
    Select Case Index
'
        Case 0, 36
            Me.Ele(1).Visible = True
            Frame3.Visible = True
            Frame6.Visible = True

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

        Case 20
            Me.Ele(1).Visible = True
            Frame2.Visible = True
            Frame1.Visible = False
            
        Case 11
            Me.Ele(1).Visible = True
            Frame2.Visible = True
            Frame1.Visible = False

        Case 12
            Me.Ele(1).Visible = True
            Frame2.Visible = True
            Frame1.Visible = False

        Case 13
            Me.Ele(1).Visible = True
'            Frame5.Visible = True

        Case 15
            Me.Ele(1).Visible = True
            Frame7.Visible = True
     
        Case 16
            Me.Ele(1).Visible = True
            Frame7.Visible = True
     
        Case 17
            Me.Ele(1).Visible = True
            Frame7.Visible = True
     
        Case 18
            Me.Ele(1).Visible = True

        Case 19
            Me.Ele(1).Visible = True
     
            DCCompositeAccount.Enabled = True
 
            StrAccountCode = ""
            TxtAccountCode.text = ""
            StrAccountName = ""
     
    End Select

End Sub

Private Sub TrvAccounts_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
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
    On Error Resume Next
    Me.LblAccountName.Caption = Me.TrvAccounts.SelectedItem.text
    txt_mod_flag.text = "N"
    StrAccountCode = Me.TrvAccounts.SelectedItem.key
    TxtAccountCode.text = Get_Account_Serial(StrAccountCode)
    StrAccountName = Me.TrvAccounts.SelectedItem.text
    
    If Grid3.rows > 1 Then
        If Grid3.TextMatrix(Grid3.rows - 1, Grid3.ColIndex("Account_Code")) <> "" Then
            Grid3.rows = Grid3.rows + 1
        End If
        Grid3.TextMatrix(Grid3.rows - 1, Grid3.ColIndex("Account_Serial")) = Get_Account_Serial(StrAccountCode)
        Grid3.TextMatrix(Grid3.rows - 1, Grid3.ColIndex("Account_Code")) = (StrAccountCode)
        Grid3.TextMatrix(Grid3.rows - 1, Grid3.ColIndex("Account_Name")) = StrAccountName
        
        
    End If
    
    
            
End Sub

Private Sub TxtAccountCode_Change()
 
'    StrAccountCode = Get_Account_code(TxtAccountCode.text, 1)
'    LblAccountName.Caption = Get_Account_name(, StrAccountCode)
'    StrAccountName = LblAccountName.Caption

End Sub

Private Sub TxtAccountCode_KeyUp(KeyCode As Integer, _
   Shift As Integer)
 
    If KeyCode = vbKeyF3 Then

        txt_mod_flag.text = "S"
 
        Account_search.show
        Account_search.case_id = 1
    
    End If
 
    If KeyCode = vbKeyReturn Then
        '        CmdAccount_Click

        StrAccountCode = Get_Account_code(TxtAccountCode.text)
        LblAccountName.Caption = Get_Account_Name(, StrAccountCode)
    End If

End Sub
Function collectionToArray(c As Collection) As String()
    Dim a() As String: ReDim a(0 To c.count - 1)
    Dim i   As Integer
    For i = 1 To c.count
        a(i - 1) = c.Item(i)
    Next
    collectionToArray = a
End Function



' Ì„‰⁄  ﬂ—«— «·⁄‰«’— œ«Œ· «·‹Collection
Private Sub AddUnique(ByRef c As Collection, ByVal s As String)
    Dim i As Long
    For i = 1 To c.count
        If StrComp(c(i), s, vbTextCompare) = 0 Then Exit Sub
    Next
    c.Add s
End Sub


Private Sub AddSQL(ByRef s As String, ByVal line As String)
    s = s & line & vbCrLf
End Sub

