VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAccountCharts 
   Caption         =   "œ·Ì· «·Õ”«»« "
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   15765
   Icon            =   "FrmAccountCharts.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   15765
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
      Height          =   8460
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15765
      _cx             =   27808
      _cy             =   14923
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
      BackColor       =   4210752
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7875
         Index           =   4
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   15720
         _cx             =   27728
         _cy             =   13891
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7815
            Index           =   6
            Left            =   11820
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   30
            Width           =   3870
            _cx             =   6826
            _cy             =   13785
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
            Begin MSComctlLib.ImageList ImgLstChartTree 
               Left            =   600
               Top             =   1980
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
                     Picture         =   "FrmAccountCharts.frx":038A
                     Key             =   "Expanded_Node"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmAccountCharts.frx":11DC
                     Key             =   "Root"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmAccountCharts.frx":1576
                     Key             =   "Open_Node"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmAccountCharts.frx":1910
                     Key             =   "Closed_Node"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmAccountCharts.frx":1CAA
                     Key             =   "Item"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.TreeView TrvAccounts 
               Height          =   7755
               HelpContextID   =   380
               Left            =   45
               TabIndex        =   5
               Top             =   30
               Width           =   3870
               _ExtentX        =   6826
               _ExtentY        =   13679
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   706
               LabelEdit       =   1
               Style           =   7
               Checkboxes      =   -1  'True
               ImageList       =   "ImgLstChartTree"
               Appearance      =   1
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7815
            Index           =   3
            Left            =   0
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   11775
            _cx             =   20770
            _cy             =   13785
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "»Ì«‰«  «·Õ”«»"
            Align           =   0
            AutoSizeChildren=   7
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
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   0
               PasswordChar    =   "*"
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   7560
               Width           =   1440
            End
            Begin VB.CommandButton Command2 
               Caption         =   " €ÌÌ— „ﬂ«‰ «·Õ”«»"
               Height          =   255
               Left            =   5625
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   7560
               Visible         =   0   'False
               Width           =   2040
            End
            Begin VB.Frame Frame8 
               Height          =   615
               Left            =   285
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   6930
               Visible         =   0   'False
               Width           =   10695
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»‰Êﬂ"
                  Height          =   195
                  Index           =   5
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   270
                  Width           =   975
               End
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Õ”«» Õ—"
                  Height          =   195
                  Index           =   4
                  Left            =   6060
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄„·«¡"
                  Height          =   195
                  Index           =   3
                  Left            =   900
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   270
                  Width           =   975
               End
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„’—Ê›« "
                  Height          =   195
                  Index           =   2
                  Left            =   2910
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄Âœ"
                  Height          =   195
                  Index           =   1
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton optMove 
                  Alignment       =   1  'Right Justify
                  Caption         =   "–„„"
                  Height          =   195
                  Index           =   0
                  Left            =   4830
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   240
                  Width           =   975
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "‰ﬁ·"
                  Height          =   255
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   210
                  Width           =   855
               End
               Begin MSDataListLib.DataCombo DboParentAccount2 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   97
                  Top             =   150
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  _Version        =   393216
                  MatchEntry      =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·Õ”«» «·—∆Ì”Ì   "
                  Height          =   315
                  Left            =   10860
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   150
                  Width           =   1230
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   90
                  Width           =   495
               End
            End
            Begin VB.CheckBox chkIsAll 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬂ·"
               Height          =   195
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   7620
               Width           =   1365
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4230
               Index           =   1
               Left            =   30
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   30
               Width           =   11715
               _cx             =   20664
               _cy             =   7461
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "»Ì«‰«  «·Õ”«»« "
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   1695
                  Left            =   0
                  TabIndex        =   84
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   5835
                  _cx             =   10292
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
                  Begin VB.ListBox ListAllUser 
                     Height          =   1035
                     ItemData        =   "FrmAccountCharts.frx":2044
                     Left            =   3165
                     List            =   "FrmAccountCharts.frx":204B
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   360
                     Width           =   2565
                  End
                  Begin VB.ListBox ListUserSelect 
                     BackColor       =   &H0080FFFF&
                     Height          =   1035
                     ItemData        =   "FrmAccountCharts.frx":205C
                     Left            =   105
                     List            =   "FrmAccountCharts.frx":2063
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   360
                     Width           =   2670
                  End
                  Begin VB.Label Label9 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„” Œœ„Ì‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   255
                     Left            =   2445
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
                     Top             =   0
                     Width           =   1035
                  End
                  Begin VB.Label LblSelect 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   480
                     Width           =   420
                  End
                  Begin VB.Label Label22 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   720
                     Width           =   420
                  End
                  Begin VB.Label Label5 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   960
                     Width           =   420
                  End
                  Begin VB.Label Label6 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   1200
                     Width           =   420
                  End
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   615
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   120
                  Width           =   2865
                  Begin VB.CheckBox Check1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·Â „Ê«“‰‹‹‹…"
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.CheckBox ChKBlock 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ìﬁ«› «· ⁄«„·"
                     ForeColor       =   &H000000C0&
                     Height          =   195
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   240
                     Width           =   1215
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ»Ì⁄Â «·—’Ìœ"
                  ForeColor       =   &H000000C0&
                  Height          =   1815
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   720
                  Width           =   3075
                  Begin VB.Frame Frame5 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   615
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   1080
                     Width           =   3375
                     Begin VB.OptionButton Differenttype 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " Õ–Ì— ›ﬁÿ"
                        Height          =   195
                        Index           =   1
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   63
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   1125
                     End
                     Begin VB.OptionButton Differenttype 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„‰⁄ „‰ « „«„ «·⁄„·Ì…"
                        Height          =   195
                        Index           =   0
                        Left            =   1440
                        RightToLeft     =   -1  'True
                        TabIndex        =   62
                        Top             =   240
                        Width           =   1725
                     End
                  End
                  Begin VB.Frame Frame3 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   615
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   240
                     Width           =   3375
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„œÌ‰"
                        Height          =   195
                        Index           =   0
                        Left            =   1440
                        RightToLeft     =   -1  'True
                        TabIndex        =   60
                        Top             =   240
                        Width           =   1605
                     End
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "œ«∆‰"
                        Height          =   195
                        Index           =   1
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   59
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   1005
                     End
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "›Ì Õ«·… „Œ«·›… ÿ»Ì⁄… «·Õ”«»"
                     ForeColor       =   &H000000C0&
                     Height          =   375
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   840
                     Width           =   2175
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’·«ÕÌ… «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   1335
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   3075
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„” Œœ„"
                     Height          =   195
                     Index           =   2
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   960
                     Value           =   -1  'True
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ã„Ê⁄…"
                     Height          =   195
                     Index           =   1
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   600
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ· «·„” Œœ„Ì‰"
                     Height          =   195
                     Index           =   0
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   240
                     Width           =   1845
                  End
                  Begin MSDataListLib.DataCombo DcUserGroup 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   55
                     Top             =   600
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcUser 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   56
                     Top             =   960
                     Width           =   2295
                     _ExtentX        =   4048
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—ﬂ“ «· ﬂ·›…"
                  ForeColor       =   &H000000C0&
                  Height          =   855
                  Left            =   3060
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   1680
                  Width           =   8505
                  Begin VB.CheckBox Check2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·Â „—ﬂ“  ﬂ·›…"
                     Height          =   255
                     Left            =   8640
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·„—ﬂ“"
                     Enabled         =   0   'False
                     Height          =   615
                     Left            =   5415
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   195
                     Width           =   3015
                     Begin VB.OptionButton Option2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "€Ì— „Õœœ"
                        Height          =   195
                        Left            =   1200
                        RightToLeft     =   -1  'True
                        TabIndex        =   42
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.OptionButton Option1 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„Õœœ"
                        Height          =   195
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   41
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin MSDataListLib.DataCombo DcCostCenter 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   44
                     Top             =   240
                     Width           =   3375
                     _ExtentX        =   5953
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ „—ﬂ“ «· ﬂ·›Â"
                     Height          =   255
                     Left            =   3480
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   240
                     Width           =   1215
                  End
               End
               Begin VB.TextBox TxtAccount_NameE 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7170
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   960
                  Width           =   2940
               End
               Begin MSDataListLib.DataCombo DboParentAccount 
                  Height          =   315
                  Left            =   7125
                  TabIndex        =   31
                  Top             =   1380
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtAccount_Code 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   4950
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   -210
                  Visible         =   0   'False
                  Width           =   690
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1695
                  Index           =   5
                  Left            =   2970
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   4170
                  _cx             =   7355
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "‰Ê⁄ «·Õ”«»"
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
                  Begin VB.ComboBox DcAccountTypes 
                     Height          =   315
                     ItemData        =   "FrmAccountCharts.frx":2077
                     Left            =   120
                     List            =   "FrmAccountCharts.frx":2084
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   900
                     Width           =   3375
                  End
                  Begin VB.ComboBox DcAccountTab 
                     Height          =   315
                     ItemData        =   "FrmAccountCharts.frx":20A1
                     Left            =   120
                     List            =   "FrmAccountCharts.frx":20B4
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   1260
                     Width           =   3375
                  End
                  Begin VB.CheckBox Check3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«»  Ã„Ì⁄Ì"
                     Height          =   255
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   315
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» ‰Â«∆Ï"
                     Height          =   255
                     Index           =   0
                     Left            =   1710
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   315
                     Value           =   -1  'True
                     Width           =   1215
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» —∆Ì”ÌÏ"
                     Height          =   240
                     Index           =   1
                     Left            =   3270
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   315
                     Width           =   1335
                  End
                  Begin MSDataListLib.DataCombo DcActivityType 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   70
                     Top             =   -480
                     Visible         =   0   'False
                     Width           =   3735
                     _ExtentX        =   6588
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Image ImgFavorites 
                     Height          =   270
                     Left            =   120
                     Picture         =   "FrmAccountCharts.frx":20E0
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   405
                  End
                  Begin VB.Label lbllevel 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   255
                     Left            =   3360
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   600
                     Width           =   1455
                  End
                  Begin VB.Label Lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»Ì⁄Â  «·‰‘«ÿ"
                     ForeColor       =   &H00000000&
                     Height          =   270
                     Index           =   10
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   -480
                     Visible         =   0   'False
                     Width           =   1230
                  End
                  Begin VB.Label LblAccType 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»Ì⁄Â «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   270
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   930
                     Width           =   1230
                  End
                  Begin VB.Label LblAccTab 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " »ÊÌ» «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   270
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   1290
                     Width           =   1110
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   1
                     Left            =   4620
                     Picture         =   "FrmAccountCharts.frx":5D48
                     Top             =   315
                     Width           =   240
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   0
                     Left            =   2940
                     Picture         =   "FrmAccountCharts.frx":60D2
                     Top             =   315
                     Width           =   240
                  End
               End
               Begin VB.TextBox TxtAccount_Name 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7170
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   540
                  Width           =   2940
               End
               Begin VB.TextBox TxtAccount_Serial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8535
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   150
                  Width           =   1575
               End
               Begin VB.TextBox TxtAccount_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   -360
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin MSDataListLib.DataCombo DCCURRENCY 
                  Height          =   315
                  Left            =   7170
                  TabIndex        =   34
                  Top             =   150
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ALLButtonS.ALLButton cmdAdd 
                  Height          =   300
                  Left            =   7365
                  TabIndex        =   71
                  Tag             =   "Delete Row"
                  Top             =   12240
                  Visible         =   0   'False
                  Width           =   1800
                  _ExtentX        =   3175
                  _ExtentY        =   529
                  BTYPE           =   3
                  TX              =   " ÕœÌÀ «·»Ì«‰« "
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
                  BCOL            =   8421376
                  BCOLO           =   8421376
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmAccountCharts.frx":645C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                  Height          =   1695
                  Left            =   5925
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   5745
                  _cx             =   10134
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
                  Begin VB.ListBox ListGroupSelected 
                     BackColor       =   &H0080FFFF&
                     Height          =   1035
                     ItemData        =   "FrmAccountCharts.frx":6478
                     Left            =   105
                     List            =   "FrmAccountCharts.frx":647F
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   360
                     Width           =   2565
                  End
                  Begin VB.ListBox ListGroupAll 
                     Height          =   1035
                     ItemData        =   "FrmAccountCharts.frx":6496
                     Left            =   3075
                     List            =   "FrmAccountCharts.frx":649D
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   360
                     Width           =   2565
                  End
                  Begin VB.Label Label10 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·›—Ê⁄"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   255
                     Left            =   2355
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   0
                     Width           =   1035
                  End
                  Begin VB.Label Label7 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   1080
                     Width           =   435
                  End
                  Begin VB.Label Label8 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   840
                     Width           =   435
                  End
                  Begin VB.Label Label13 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   600
                     Width           =   435
                  End
                  Begin VB.Label Label14 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Left            =   2655
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   360
                     Width           =   435
                  End
               End
               Begin VB.Label lblCurrency 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„·…"
                  Height          =   195
                  Left            =   7665
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   240
                  Width           =   435
               End
               Begin VB.Label LblNameE 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Õ”«» «‰Ã·Ì“Ì"
                  Height          =   285
                  Left            =   10425
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   990
                  Width           =   1245
               End
               Begin VB.Label lblParentAcc 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·Õ”«» «·—∆Ì”Ì   "
                  Height          =   315
                  Left            =   10635
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   1380
                  Width           =   1035
               End
               Begin VB.Label LblNameA 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Õ”«» ⁄—»Ì"
                  Height          =   285
                  Left            =   10425
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label LblCode 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·Õ”«»"
                  Height          =   345
                  Left            =   10635
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ”«»"
                  Height          =   345
                  Index           =   1
                  Left            =   4185
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   795
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1530
               Index           =   7
               Left            =   30
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   810
               Width           =   1680
               _cx             =   2963
               _cy             =   2699
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
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   255
                  Index           =   0
                  Left            =   30
                  TabIndex        =   26
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   450
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmAccountCharts.frx":64AF
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   255
                  Index           =   1
                  Left            =   450
                  TabIndex        =   27
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   450
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmAccountCharts.frx":6849
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   285
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   28
                  Top             =   30
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   503
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
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
                  ButtonImage     =   "FrmAccountCharts.frx":6BE3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   2
               Left            =   30
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   6690
               Width           =   11715
               _cx             =   20664
               _cy             =   1349
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
               Begin VB.TextBox TxtModflg 
                  Alignment       =   1  'Right Justify
                  Height          =   135
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   15
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Image Image1 
                  Height          =   240
                  Left            =   11325
                  Picture         =   "FrmAccountCharts.frx":6F7D
                  Top             =   30
                  Width           =   240
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ… Â«„…:-"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   240
                  Index           =   4
                  Left            =   9585
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   30
                  Width           =   1635
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
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
                  Height          =   465
                  Index           =   5
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   300
                  Width           =   11355
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts 
               Height          =   2370
               Left            =   -75
               TabIndex        =   50
               Top             =   4275
               Width           =   11820
               _cx             =   20849
               _cy             =   4180
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
               Rows            =   50
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmAccountCharts.frx":7307
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
            Begin XtremeSuiteControls.PushButton cmdRenew 
               Height          =   270
               Left            =   1800
               TabIndex        =   108
               Top             =   7560
               Visible         =   0   'False
               Width           =   375
               _Version        =   786432
               _ExtentX        =   661
               _ExtentY        =   476
               _StockProps     =   79
               Caption         =   "R"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Currency"
               Height          =   1530
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   810
               Width           =   2595
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„·… «·Õ”«»"
               Height          =   765
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   30
               Width           =   1680
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7905
         Width           =   15720
         _cx             =   27728
         _cy             =   953
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   13905
            TabIndex        =   7
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            Height          =   375
            Index           =   1
            Left            =   12150
            TabIndex        =   8
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
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
            Height          =   375
            Index           =   2
            Left            =   10440
            TabIndex        =   9
            Top             =   90
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   661
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
            Height          =   375
            Index           =   3
            Left            =   9030
            TabIndex        =   10
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
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
            Height          =   375
            Index           =   4
            Left            =   6720
            TabIndex        =   11
            Top             =   90
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   661
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
            Height          =   375
            Index           =   5
            Left            =   5310
            TabIndex        =   12
            Top             =   120
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
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
            Height          =   375
            Index           =   6
            Left            =   570
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   90
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
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
            Height          =   375
            Index           =   7
            Left            =   3570
            TabIndex        =   14
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
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
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
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
            Height          =   375
            Index           =   8
            Left            =   1245
            TabIndex        =   72
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   330
            Left            =   0
            TabIndex        =   76
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
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
            ButtonImage     =   "FrmAccountCharts.frx":74F6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ReNew Serials"
      Visible         =   0   'False
      Begin VB.Menu mnEdit 
         Caption         =   "Ren"
      End
   End
End
Attribute VB_Name = "FrmAccountCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xAccCol As Collection
Dim IntCurrentIndex As Integer
Dim Dcombos As ClsDataCombos
Dim StrTemp As String
Dim FirstPeriodDateInthisYear  As Date
Dim DesStr As String

Private dictParents As Object
Sub SaveBransh_UserAccount(Optional StrNewAccountCode As String)
    Dim i   As Integer
    Dim sql As String
    Dim Rs3 As ADODB.Recordset
    If ListGroupSelected.ListCount >= 0 Then
        sql = "Select * from  TblAccountBranch where 1=-1"
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        For i = 0 To ListGroupSelected.ListCount - 1
            Rs3.AddNew
            Rs3("BranchID").value = ListGroupSelected.ItemData(i)
            Rs3("Account_Code").value = Trim(StrNewAccountCode)
            Rs3.update
        Next i
    End If

    If ListUserSelect.ListCount >= 0 Then
        sql = "Select * from  TblAccountUser where 1=-1"
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        For i = 0 To ListUserSelect.ListCount - 1
            Rs3.AddNew
            Rs3("UserID").value = ListUserSelect.ItemData(i)
            Rs3("Account_Code").value = Trim(StrNewAccountCode)
            Rs3.update
        Next i
    End If
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·Õ”«»  " & TxtAccount_Serial.text & CHR(13) & "   «”„ «·Õ”«»  " & TxtAccount_Name & CHR(13) & "     «·Õ”«»   «·—∆Ì”Ì " & DboParentAccount & CHR(13) & "   ÿ»Ì⁄Â  «·Õ”«»  " & DcAccountTypes & CHR(13) & "    »ÊÌ»  «·Õ”«»  " & DcAccountTab

    If OptAccountType(0).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ‰Ê⁄ «·Õ”«» " & "  Õ”«» ‰Â«∆Ï  "
    ElseIf OptAccountType(1).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ‰Ê⁄ «·Õ”«» " & "  Õ”«» —∆Ì”Ì  "
    End If
                   
    If Check2.value = Checked Then
        LogTextA = LogTextA & CHR(13) & " ·Â „—ﬂ“  ﬂ·›…  "

        If Option2.value = True Then
            LogTextA = LogTextA & "  -  €Ì— „Õœœ"
        ElseIf Option1.value = True Then
            LogTextA = LogTextA & "  -    „Õœœ"
                                                                 
            LogTextA = LogTextA & CHR(13) & "   „—ﬂ“  ﬂ·›…  " & DcCostCenter
                                                                 
        End If
                   
    End If
              
    If DepitOrCredit(0).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ÿ»Ì⁄Â «·Õ”«» " & " „œÌ‰ "
    ElseIf DepitOrCredit(1).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ÿ»Ì⁄Â «·Õ”«» " & " œ«∆‰ "
    End If
                   
    If Differenttype(0).value = True Then
        LogTextA = LogTextA & CHR(13) & " ›Ì Õ«·… „Œ«·›… ÿ»Ì⁄… «·Õ”«» " & " „‰⁄ „‰ « „«„ «·⁄„·Ì… "
    ElseIf Differenttype(1).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ›Ì Õ«·… „Œ«·›… ÿ»Ì⁄… «·Õ”«» " & "  Õ–Ì— ›ﬁÿ "
    End If
                   
    If Authority(0).value = True Then
        LogTextA = LogTextA & CHR(13) & " ’·«ÕÌ… «· ⁄«„·  " & "  ﬂ· «·„” Œœ„Ì‰ "
    ElseIf Authority(1).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ’·«ÕÌ… «· ⁄«„· " & " „Ã„Ê⁄Â „ÕœœÂ " & DcUserGroup
    ElseIf Authority(2).value = True Then
        LogTextA = LogTextA & CHR(13) & "  ’·«ÕÌ… «· ⁄«„· " & "   „” Œœ„ „Õœœ  " & DcUser
                   
    End If
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "  Acc.  Code  " & TxtAccount_Serial.text & CHR(13) & "     Acc.  Name    " & TxtAccount_NameE & CHR(13) & "   Parent   Acc.  " & DboParentAccount & CHR(13) & "   Acc.  Type " & DcAccountTypes & CHR(13) & "     Acc.  Tab  " & DcAccountTab

    If OptAccountType(0).value = True Then
        LogTexte = LogTexte & CHR(13) & "   Acc. Type  " & "  Final   Acc.   "
    ElseIf OptAccountType(1).value = True Then
        LogTexte = LogTexte & CHR(13) & "   Acc.  Type " & "  Parent   Acc.   "
    End If
                   
    If Check2.value = Checked Then
        LogTexte = LogTexte & CHR(13) & " Have Cost Center "

        If Option2.value = True Then
            LogTexte = LogTexte & "  - un Specific "
        ElseIf Option1.value = True Then
            LogTexte = LogTexte & "-Specific"
                                                                 
            LogTexte = LogTexte & CHR(13) & "   Cost Center    " & DcCostCenter
                                                                 
        End If
                   
    End If
              
    If DepitOrCredit(0).value = True Then
        LogTexte = LogTexte & CHR(13) & " Acc. Type " & " Debit "
    ElseIf DepitOrCredit(1).value = True Then
        LogTexte = LogTexte & CHR(13) & "  Acc. Type " & " Credit "
    End If
                   
    If Differenttype(0).value = True Then
        LogTexte = LogTexte & CHR(13) & " In the case of violation of the nature of the account " & "  Not to complete the transaction   "
    ElseIf Differenttype(1).value = True Then
        LogTexte = LogTexte & CHR(13) & "In the case of violation of the nature of the account " & "   Warning only "
    End If
                   
    If Authority(0).value = True Then
        LogTexte = LogTexte & CHR(13) & "  Authority    " & "  All Users "
    ElseIf Authority(1).value = True Then
        LogTexte = LogTexte & CHR(13) & "  Authority    " & "   Specific Group " & DcUserGroup
    ElseIf Authority(2).value = True Then
        LogTexte = LogTexte & CHR(13) & "   Authority   " & "      Specific User  " & DcUser
                   
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModflg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If

End Sub

Function CheCkAutoAccountx(Optional AccountCode As String, Optional strOutput As String) As Boolean
 Dim sql As String
 Dim Rs3 As ADODB.Recordset
 Set Rs3 = New ADODB.Recordset
 sql = "SELECT     Account_Code"
 sql = sql & "   FROM         (SELECT     Account_Code"
 sql = sql & "                       From dbo.ExpensesType"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.TblStore"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentAccount AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code3 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code4 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParetnAccount AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParetnAccount1 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     RootAccount AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentAccountSub AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentAccount1 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     RootAccount1 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code5 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code6 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code7 AS Account_Code"
 sql = sql & "                       From dbo.Tblinvestment"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.TblBuyLanReEst"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.TblStore"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.TblStore"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code3 AS Account_Code"
 sql = sql & "                       From dbo.TblStore"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParetnAccount AS Account_Code"
 sql = sql & "                       From dbo.TblStore"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.TblRevenuesTypes"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Material_account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParetnAccount AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code3 AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code4 AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentExpensesAccount AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentEAssetAccount AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code5 AS Account_Code"
 sql = sql & "                       From dbo.FixedAssetsGroup"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Project_account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Salary_account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     REVENUE_account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     expanses_account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     sub_contractor_Account AS Account_Code"
 sql = sql & "                       From dbo.Projects"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.BanksData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.BanksData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.BanksData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_code3 AS Account_Code"
 sql = sql & "                       From dbo.BanksData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     parent_account AS Account_Code"
 sql = sql & "                       From dbo.BanksData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.TblBoxesData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.TblBoxesData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.TblBoxesData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentAccount AS Account_Code"
 sql = sql & "                       From dbo.TblBoxesData"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     parent_account AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     ParentAccount AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code1 AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code_As_Client AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code_As_Supplier AS Account_Code"
 sql = sql & "                       From dbo.TblCustemers"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_code AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_code1 AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code2 AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                        Union all"
 sql = sql & "                       SELECT     Account_Code3 AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code4 AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code5 AS Account_Code"
 sql = sql & "                       From dbo.TblEmployee"
 sql = sql & "                       Union all"
 sql = sql & "                       SELECT     Account_Code5 AS Account_Code"
 sql = sql & "                       FROM         dbo.TblEmployee) DERIVEDTBL"
 sql = sql & " WHERE     (NOT (DERIVEDTBL.Account_Code IS NULL) AND DERIVEDTBL.Account_Code = N'" & AccountCode & "')"
 Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs3.RecordCount > 0 Then
 CheCkAutoAccountx = True
 Else
 CheCkAutoAccountx = False
 End If
End Function
Function CheCkAutoAccount(Optional AccountCode As String, Optional ByRef strOutput As String) As Boolean
 Dim sql As String
 Dim Rs3 As ADODB.Recordset
 Set Rs3 = New ADODB.Recordset
 Dim i As Integer
 
    
    sql = " SELECT Account_Code FROM ("
    sql = sql & "SELECT Account_Code  From dbo.ExpensesType"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Store Account'                             AS Account_Code From dbo.TblStore"
    sql = sql & " Union all SELECT '?? ' + ParentAccount              +'  -ParentAccount Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code Tblinvestment'                AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code3              +'  -Account_Code3 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code4              +'  -Account_Code4 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + ParetnAccount              +'  -ParetnAccount Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + ParetnAccount1             +'  -ParetnAccount1 Tblinvestment'              AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + RootAccount                +'  -RootAccount Tblinvestment'                 AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + ParentAccountSub           +'  -ParentAccountSub Tblinvestment'            AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + ParentAccount1             +'  -ParentAccount1 Tblinvestment'              AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + RootAccount1               +'  -RootAccount1 Tblinvestment'                AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code5              +'  -Account_Code5 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code6              +'  -Account_Code6 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code7              +'  -Account_Code7 Tblinvestment'               AS Account_Code From dbo.Tblinvestment"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code TblBuyLanReEst'               AS Account_Code From dbo.TblBuyLanReEst"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 TblStore'                    AS Account_Code From dbo.TblStore"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 TblStore'                    AS Account_Code From dbo.TblStore"
    sql = sql & " Union all SELECT '?? ' + Account_Code3              +'  -Account_Code3 TblStore'                    AS Account_Code From dbo.TblStore"
    sql = sql & " Union all SELECT '?? ' + ParetnAccount              +'  -ParetnAccount TblStore'                    AS Account_Code From dbo.TblStore"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code TblRevenuesTypes'             AS Account_Code From dbo.TblRevenuesTypes"
    sql = sql & " Union all SELECT '?? ' + ParetnAccount              +'  -ParetnAccount FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code FixedAssetsGroup'             AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code3              +'  -Account_Code3 FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code4              +'  -Account_Code4 FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + ParentExpensesAccount      +'  -ParentExpensesAccount FixedAssetsGroup'    AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + ParentEAssetAccount        +'  -ParentEAssetAccount FixedAssetsGroup'      AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Account_Code5              +'  -Account_Code5 FixedAssetsGroup'            AS Account_Code From dbo.FixedAssetsGroup"
    sql = sql & " Union all SELECT '?? ' + Material_account           +'  -Material_account Projects'                 AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + Project_account            +'  -Project_account Projects'                  AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + Salary_account             +'  -Salary_account Projects'                   AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + REVENUE_account            +'  -REVENUE_account Projects'                  AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + expanses_account           +'  -expanses_account Projects'                 AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + sub_contractor_Account     +'  -sub_contractor_Account Projects'           AS Account_Code From dbo.Projects"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code BanksData'                    AS Account_Code From dbo.BanksData"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 BanksData'                   AS Account_Code From dbo.BanksData"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 BanksData'                   AS Account_Code From dbo.BanksData"
    sql = sql & " Union all SELECT '?? ' + Account_code3              +'  -Account_code3 BanksData'                   AS Account_Code From dbo.BanksData"
    sql = sql & " Union all SELECT '?? ' + parent_account             +'  -parent_account BanksData'                  AS Account_Code From dbo.BanksData"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code TblBoxesData'                 AS Account_Code From dbo.TblBoxesData"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 TblBoxesData'                AS Account_Code From dbo.TblBoxesData"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 TblBoxesData'                AS Account_Code From dbo.TblBoxesData"
    sql = sql & " Union all SELECT '?? ' + ParentAccount              +'  -ParentAccount TblBoxesData'                AS Account_Code From dbo.TblBoxesData"
    sql = sql & " Union all SELECT '?? ' + parent_account             +'  -parent_account TblCustemers'               AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + ParentAccount              +'  -ParentAccount TblCustemers'                AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 TblCustemers'                AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_Code1              +'  -Account_Code1 TblCustemers'                AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_Code               +'  -Account_Code TblCustemers'                 AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_Code_As_Client     +'  -Account_Code_As_Client TblCustemers'       AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_Code_As_Supplier   +'  -Account_Code_As_Supplier TblCustemers'     AS Account_Code From dbo.TblCustemers"
    sql = sql & " Union all SELECT '?? ' + Account_code               +'  -Account_code TblEmployee'                  AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_code1              +'  -Account_code1 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_Code2              +'  -Account_Code2 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_Code3              +'  -Account_Code3 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_Code4              +'  -Account_Code4 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_Code5              +'  -Account_Code5 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " Union all SELECT '?? ' + Account_Code5              +'  -Account_Code5 TblEmployee'                 AS Account_Code From dbo.TblEmployee"
    sql = sql & " )"
    sql = sql & " DERIVEDTBL WHERE     (NOT (DERIVEDTBL.Account_Code IS NULL) AND DERIVEDTBL.Account_Code like  '?? " & AccountCode & "%')"
    
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    strOutput = ""
    If Rs3.RecordCount > 0 Then
            Rs3.MoveFirst
        For i = 0 To Rs3.RecordCount - 1
            strOutput = strOutput & Rs3("Account_Code").value & CHR(13)
            Rs3.MoveNext
        Next i
        
        CheCkAutoAccount = True
    Else
        CheCkAutoAccount = False

    End If
End Function

Private Sub Cmd_Click(Index As Integer)
    Dim cReport        As ClsAccReports
    Dim StrAccountCode As String

    Select Case Index
        Case 8
            Dcombos.GetAccountingCodes Me.DboParentAccount
            LoadData
    
        Case 0
            Me.TxtModflg.text = "N"
       
            DCCURRENCY.BoundText = 1
            TxtAccount_Serial.text = ""
            Me.DboParentAccount.BoundText = StrTemp
            TxtAccount_Name.text = ""
            TxtAccount_NameE.text = ""
            ListGroupSelected.Clear
            ListUserSelect.Clear
            Check2.value = vbUnchecked
            Option1.value = False
            Option2.value = True

        Case 1
            Ele(5).Enabled = True
            Frame6.Enabled = True
            StrAccountCode = Trim(Me.TxtAccount_Code.text)
            If StrAccountCode = "a1" Or StrAccountCode = "a2" Or StrAccountCode = "a3" Or StrAccountCode = "a4" Or StrAccountCode = "a5" Or StrAccountCode = "a6" Then GoTo ll
            If CheckDelAccount(StrAccountCode) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "  Â–«     «·Õ”«» ·« Ì„ﬂ‰ Õ–›… ", vbInformation
              
                Else
                    MsgBox "  Can't Modify This Account it have Transactions", vbInformation
                End If
ll:
                '  Exit Sub
                Ele(5).Enabled = False
                Frame6.Enabled = False
            End If
            
            If SystemOptions.AllowEditeAccounts = False Then
                If CheCkAutoAccount(TxtAccount_Code.text, DesStr) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Â–« «·Õ”«» «·Ì ·«Ì„ﬂ‰  ⁄œÌ·Â" & CHR(13) & DesStr
                    Else
                        MsgBox "This is  Account auto. It can not be modified" & CHR(13) & DesStr
                    End If
                    Exit Sub
                End If
            End If
 
            Me.TxtModflg.text = "E"
            CuurentLogdata

        Case 2
            SaveData

        Case 3
            Me.TxtModflg.text = "R"

        Case 4
            If CheCkAutoAccount(TxtAccount_Code.text, DesStr) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Â–« «·Õ”«» «·Ì ·«Ì„ﬂ‰ Õ–›Â" & CHR(13) & DesStr
                Else
                    MsgBox "This is  Account auto. It can not be delete" & CHR(13) & DesStr
                End If
                Exit Sub
            End If
            DelAccount

        Case 5
            Account_search.show
            Account_search.case_id = 0
        
        Case 6
            Unload Me

        Case 7
            Set cReport = New ClsAccReports
            cReport.ShowChartAccounts WindowTarget, , IIf(chkIsAll.value, True, False)
            Set cReport = Nothing
    End Select

End Sub

Private Sub cmdAdd_Click()
GetHeaders Me

  
    'Me.Retrive StrNewAccountCode

End Sub

Private Sub CmdN_Click(Index As Integer)
    Dim StrTemp As String
    Dim xx As Object

    Select Case Index

        Case 0

            'Back Move
            If IntCurrentIndex = 0 Then
                IntCurrentIndex = xAccCol.count
            Else
                IntCurrentIndex = IntCurrentIndex - 1
            End If

            If IntCurrentIndex = 0 Then
                'Me.CmdN(0).Enabled = False
            Else

                If IntCurrentIndex <= xAccCol.count Then
                    StrTemp = xAccCol.Item("A" & IntCurrentIndex)
                    Me.Retrive StrTemp, False
                Else
                    IntCurrentIndex = 1
                End If
            End If

        Case 1

            'Forward Move
            If IntCurrentIndex = 0 Then
                IntCurrentIndex = xAccCol.count
            Else
                IntCurrentIndex = IntCurrentIndex + 1
            End If

            If IntCurrentIndex = 0 Then
                'Me.CmdN(0).Enabled = False
            Else

                If IntCurrentIndex <= xAccCol.count Then
                    StrTemp = xAccCol.Item("A" & IntCurrentIndex)
                    Me.Retrive StrTemp, False
                Else
                    IntCurrentIndex = 1
                End If
            End If

        Case 2
            GetUpLevel
    End Select

End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = " SELECT     UserID, UserName"
    sql = sql & "         From dbo.TblUsers"
    sql = sql & " order by UserName"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListAllUser.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            ListAllUser.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
            ListAllUser.ItemData(ListAllUser.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

    'fil

    sql = " SELECT * from  TblBranchesData where not  ActivityTypeId is null "
 
 If SystemOptions.UserInterface = ArabicInterface Then
sql = sql & " order by  branch_name"
Else
sql = sql & " order by  branch_name"
End If
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListGroupAll.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function
Function FillMylistData(Optional AccountCode As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = " SELECT     dbo.TblAccountBranch.BranchID, dbo.TblAccountBranch.Account_Code, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    sql = sql & "    FROM         dbo.TblAccountBranch INNER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblAccountBranch.BranchID = dbo.TblBranchesData.branch_id"
    sql = sql & "     WHERE     (dbo.TblAccountBranch.Account_Code = N'" & AccountCode & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListGroupSelected.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
        If SystemOptions.UserInterface = ArabicInterface Then
            ListGroupSelected.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
          Else
            ListGroupSelected.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
       End If
            ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = rs("BranchID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

    'fil

    sql = " SELECT     dbo.TblAccountUser.UserID, dbo.TblAccountUser.Account_Code, dbo.TblUsers.UserName"
    sql = sql & "    FROM         dbo.TblAccountUser LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblUsers ON dbo.TblAccountUser.UserID = dbo.TblUsers.UserID"
    sql = sql & "  WHERE     (dbo.TblAccountUser.Account_Code = N'" & AccountCode & "')"
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ListUserSelect.Clear
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
                ListUserSelect.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                ListUserSelect.ItemData(ListUserSelect.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function

Private Sub cmdRenew_Click()
    Dim Msg
    Msg = IIf(SystemOptions.UserInterface = ArabicInterface, "Â· «‰  „ «ﬂœ „‰ «⁄«œÂ  ﬂÊÌœ «·Õ”«»«  «·›—⁄ÌÂ", "renew all childs account ? ")
    If MsgBox(Msg, vbYesNo) = vbYes Then
        Dim rschilds         As New ADODB.Recordset
        Dim StrParentAccCode As String
        StrParentAccCode = Trim(Me.TxtAccount_Code.text)
        Dim MySQL
        MySQL = "SELECT * "
        MySQL = MySQL & "FROM ACCOUNTS "
        MySQL = MySQL & "WHERE Parent_Account_Code = '" & StrParentAccCode & "'  AND last_account = 1 ;"
        Dim Count_ACCOUNT_digit As Integer
        Dim NoOfAs              As Integer
     
        NoOfAs = CountAs(StrParentAccCode) + 1
        Count_ACCOUNT_digit = GetAccountsLevel(NoOfAs)
        rschilds.Open MySQL, Cn, adOpenKeyset, adLockOptimistic
        Dim startIndex As Integer
        
        Do While Not rschilds.EOF
            startIndex = startIndex + 1
            Dim newSerial
            newSerial = Get_Account_Serial(StrParentAccCode) & Format(startIndex, String(Count_ACCOUNT_digit, "0"))
            rschilds!Account_Serial = newSerial
            rschilds.update
            rschilds.MoveNext
        Loop
        rschilds.Close
        MsgBox "Done"
        Dcombos.GetAccountingCodes Me.DboParentAccount
        LoadData
        Me.Retrive StrParentAccCode
    End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim rsDummy2 As New ADODB.Recordset
Dim StrNewAccountCode As String

s = "Delete tmpAccount"
Cn.Execute s

    s = " SELECT a.last_account,"
s = s & "         a.Account_Code,"
s = s & "                 a.Account_Name,"
s = s & "                 a2.Account_Name      AS parantName,"
s = s & "                 a.Parent_Account_Code"
s = s & "          FROM   ACCOUNTS             AS a"
s = s & "                 INNER JOIN ACCOUNTS  AS a2"
s = s & "                      ON  a.Parent_Account_Code = a2.Account_Code"
s = s & "          WHERE  a.Account_Code IN (SELECT Code"
s = s & "                                    FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(TxtAccount_Code) & "', '" & Trim(TxtAccount_Code) & "', 1))"
s = s & " OR (a.Account_Code = '" & Trim(TxtAccount_Code) & "')"
s = s & "          Order By"
s = s & "                 a.Parent_Account_Code,"
s = s & "                 a.last_account"
 
 


Dim rsD As New ADODB.Recordset
Dim ss As String
Dim mIsLast As Boolean
Dim mParent As String
mParent = Trim(DboParentAccount2.BoundText)
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
Do While Not rsDummy.EOF
    
    ss = "Select * from tmpAccount Where Account_Code = N'" & Trim(rsDummy!Parent_Account_Code & "") & "' and last_account = 0 "
    
    Set rsDummy2 = New ADODB.Recordset
    rsDummy2.Open ss, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy2.EOF Then
        mParent = Trim(rsDummy2!NewAccount_Code & "")
    Else
        mParent = Trim(DboParentAccount2.BoundText)
    End If
    StrNewAccountCode = AddNewAccount(mParent, Trim$(rsDummy!account_name & ""), CBool(rsDummy!last_account & ""), False, Trim$(rsDummy!account_name & ""), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(DepitOrCredit(0).value = True, 0, 1), 0, 0, 0, 1, False)
    SaveBransh_UserAccount StrNewAccountCode
   
    If Not CBool(rsDummy!last_account & "") Then
        ss = "Select * from tmpAccount "
        Set rsD = New ADODB.Recordset
        rsD.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
        rsD.AddNew
        rsD!Account_code = Trim$(rsDummy!Account_code & "")
        rsD!account_name = Trim$(rsDummy!account_name & "")
        rsD!Parent_Account_Code = Trim$(rsDummy!Parent_Account_Code & "")
        rsD!last_account = rsDummy!last_account
        rsD!NewAccount_Code = StrNewAccountCode
        rsD!NewParent_Account_Code = mParent
        rsD.update
        
    End If
    
    If CBool(rsDummy!last_account & "") Then
        If optMove(0).value = True Then
            s = "Update TblEmployee Set Account_code = N'" & Trim(StrNewAccountCode) & "' "
            s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
            Cn.Execute s
            
        ElseIf optMove(1).value = True Then
            s = "Update TblBoxesData Set Account_code = N'" & Trim(StrNewAccountCode) & "' ,parent_account = N'" & Trim(mParent) & "'"
            s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
            Cn.Execute s
            
            
        ElseIf optMove(5).value = True Then
            s = "Update BanksData Set Account_code = N'" & Trim(StrNewAccountCode) & "' ,parent_account = N'" & Trim(mParent) & "'"
            s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
            Cn.Execute s
            
            
        ElseIf optMove(2).value = True Then
            s = "Update ExpensesType Set Account_code = N'" & Trim(StrNewAccountCode) & "' ,parent_account = N'" & Trim(mParent) & "'"
            s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
            Cn.Execute s
        
        
        ElseIf optMove(3).value = True Then
            
            s = "Update TblCustemers Set  Account_code = N'" & Trim(StrNewAccountCode) & "' ,parent_account  = N'" & Trim(mParent) & "'"
            s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
            Cn.Execute s
        End If
        
        
        s = "Update DOUBLE_ENTREY_VOUCHERS1 Set Account_Code = N'" & Trim(StrNewAccountCode) & "'"
        s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
        Cn.Execute s
        
        s = "Update DOUBLE_ENTREY_VOUCHERS Set Account_Code = N'" & Trim(StrNewAccountCode) & "'"
        s = s & " Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
        Cn.Execute s
        DoEvents
        
'         s = "Delete ACCOUNTS Where Account_code = N'" & Trim(rsDummy!Account_code & "") & "'"
'         Cn.Execute s
    
    End If
    rsDummy.MoveNext
    
   
Loop

MsgBox " „ «·‰ﬁ·"

End Sub

Private Sub Command2_Click()
    If SystemOptions.UserInterface = EnglishInterface Then
        Dcombos.GetAccountingCodes Me.DboParentAccount2, False, True
    Else
        Dcombos.GetAccountingCodes Me.DboParentAccount2, False, True
    End If
  
Frame8.Visible = True

End Sub

Private Sub DboParentAccount2_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 789725
    End If

End Sub

Private Sub FgAccounts_Click()
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long
    Dim XNode As MSComctlLib.Node
    Dim StrAcountCode  As String

    With Me.FgAccounts
        
        LngMouseCol = .MouseCol
        LngMouseRow = .MouseRow

        If LngMouseRow <= 0 Then Exit Sub
        If LngMouseCol <> .ColIndex("Account_Name") Then
            Exit Sub
        End If

        If LngMouseCol = .ColIndex("Account_Name") Then
            If .cell(flexcpFontBold, LngMouseRow, LngMouseCol) = True Then
                StrAcountCode = .TextMatrix(LngMouseRow, .ColIndex("Account_Code"))
                Me.Retrive StrAcountCode
                Set XNode = Me.TrvAccounts.Nodes(StrAcountCode & "G")
                Me.TrvAccounts.Nodes(XNode.key).EnsureVisible
                Me.TrvAccounts.Nodes(XNode.key).Expanded = True
                Me.TrvAccounts.Nodes(XNode.key).Selected = True
            End If

        Else
            Exit Sub
        End If

    End With

End Sub

Private Sub FgAccounts_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    With Me.FgAccounts
        .ToolTipText = ""
        .MousePointer = flexDefault
        .cell(flexcpForeColor, 0, 0, .rows - 1, .Cols - 1) = vbBlack
        .cell(flexcpFontUnderline, 0, 0, .rows - 1, .Cols - 1) = False
        
        LngMouseCol = .MouseCol
        LngMouseRow = .MouseRow

        If LngMouseRow <= 0 Then Exit Sub
        If LngMouseCol <> .ColIndex("Account_Name") Then
            .MousePointer = flexDefault
            Exit Sub
        End If

        If LngMouseCol = .ColIndex("Account_Name") Then
            If .cell(flexcpFontBold, LngMouseRow, LngMouseCol) = True Then
                .MousePointer = flexHand
                .ToolTipText = "≈÷€ÿ Â‰« Õ Ï Ì„ﬂ‰ﬂ „‘«Âœ… «·Õ”«»«  «·›—⁄Ì… „‰ Â–« «·Õ”«»"
                .cell(flexcpForeColor, LngMouseRow, LngMouseCol) = vbBlue
                .cell(flexcpFontUnderline, LngMouseRow, LngMouseCol) = True
            End If

        Else
            .MousePointer = flexDefault
            .cell(flexcpForeColor, 0, 0, .rows - 1, .Cols - 1) = vbBlack
            .cell(flexcpFontUnderline, 0, 0, .rows - 1, .Cols - 1) = False
        End If

    End With

End Sub

Private Sub ChangeLang()
    DcAccountTypes.Clear
    DcAccountTypes.AddItem "Without"
    DcAccountTypes.AddItem "Balance Sheet"
    DcAccountTypes.AddItem "Income Statement"
cmdAdd.Caption = "Refresh"
Cmd(8).Caption = "Update"
Label10.Caption = "Branches"
Label9.Caption = "Users"
    DcAccountTab.Clear
    DcAccountTab.AddItem "Assets"
    DcAccountTab.AddItem "Liabilities"
    DcAccountTab.AddItem "Revenue"
    DcAccountTab.AddItem "Expenses"
    DcAccountTab.AddItem "Legall Acc."

    ChKBlock.Caption = "Block"
    Frame6.Caption = "Balance Type"
    DepitOrCredit(0).Caption = "Depit"
    DepitOrCredit(1).Caption = "Credit"
    Label4.Caption = "In Different Case"
    Differenttype(0).Caption = "Acess Deny"
    Differenttype(1).Caption = "Alarm Only"
    LblAccType.Caption = "Acc. Type"
     LblAccTab.Caption = "Acc. Class."
    Frame4.Caption = "Authority"
    Authority(0).Caption = "All Users"
    Authority(1).Caption = "Group"
    Authority(2).Caption = "User"

    Frame1.Caption = "Cost Center"
    Frame2.Caption = "C.C Type"
    Option1.Caption = "Fixed"
    Option2.Caption = "Not Fixed"
    Label3.Caption = "CC Name"
    Me.Caption = "Account Chart"
    Ele(1).Caption = "Account Data"
    Lbl(1).Caption = "Account#"
    LblCode.Caption = "Acc. Code"
    lblParentAcc.Caption = "Parent Account"
    LblNameA.Caption = "Name A"
    LblNameE.Caption = "Name E"
    'lbl(3).Caption = "Derived Account From This Acc"
    Lbl(4).Caption = "Note"
    Lbl(5).Visible = False
    Check1.Caption = "Budget Acc."
    Check2.Caption = "Cost Center"
    Check3.Caption = "Sum account"
    lblCurrency.Caption = "Curr"

    Ele(5).Caption = "Account Type"

    With FgAccounts
        .TextMatrix(0, .ColIndex("Account_ID")) = "Account ID"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code"
        .TextMatrix(0, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(0, .ColIndex("OpenAccount")) = "Opening Balance"
        .TextMatrix(0, .ColIndex("OpenAccountType")) = "Opening Balance State"

        .TextMatrix(0, .ColIndex("AccountState")) = "Account State"
        .TextMatrix(0, .ColIndex("DateCreated")) = "DateCreated"
        .TextMatrix(0, .ColIndex("CurrentAccount")) = "Current Account"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
 
    End With

    OptAccountType(0).Caption = "Final"
    OptAccountType(1).Caption = "Master"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

End Sub

Private Sub Form_Load()
    Me.Height = 10000
    Me.Width = 17600
    
   ' Me.left = (mdifrmmain.Width - Me.Width) / 2
   ' Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
        Me.left = (mdifrmmain.Width - Me.Width) / 2 - 1200
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    Dim Msg As String
    FillMylist
    Dim My_SQL As String
    My_SQL = "  select id,code from currency"

    fill_combo DCCURRENCY, My_SQL

    Dim GrdBack As ClsBackGroundPic

If dictParents Is Nothing Then Set dictParents = BuildParentsCache(Cn)

    ScreenNameArabic = " œ·Ì· «·Õ”«»« "
    ScreenNameEnglish = " Accounts Chart"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Msg = "ÌÃ» «‰  ·«ÕŸ «‰Â ·«Ì„ﬂ‰ ≈÷«›… Õ”«»«  ( «·⁄„·«¡ «Ê «·„Ê—œÌ‰) «Ê Õ”«»«  («·Œ“Ì‰… «Ê Õ”«»«  «·»‰Êﬂ)"
    Msg = Msg + " «Ê «·Õ”«»«  «·„ ›—⁄… „‰ Õ”«» «·„’—Ê›«  „‰ Œ·«· ‘«‘… «·œ·Ì· «·„Õ«”»Ï "
    Msg = Msg + " Ê–·ﬂ ·«‰ «·»—‰«„Ã ÌﬁÊ„  ·ﬁ«∆Ì« „‰ »≈÷«›… ﬂ· Â–Â «·Õ”«»«  „‰ «·‘«‘«  «·Œ«’… »ﬂ· »Â–Â «·»Ì«‰« "
    Msg = Msg + " »„⁄‰Ï «‰Â ›Ê— ≈÷«›… «Ê  ”ÃÌ· ⁄„Ì· ÃœÌœ „‰ ‘«‘… »Ì«‰«  «·⁄„·«¡ ›«‰ «·»—‰«„Ã"
    Msg = Msg + " ÌﬁÊ„ ⁄·Ï «·›Ê— »«‰‘«¡ Õ”«» Œ«’ ·Â–« «·⁄„Ì· ›Ï «·œ·Ì· «·„Õ«”»Ï (»‰›” «”„ «·⁄„Ì·)"
    Msg = Msg + "Êﬂ–·ﬂ «·Õ«· »«·‰”»… ··„Ê—œÌ‰ .Ê‰›” «·Õ«·… ›Ï Õ«·… ≈÷«›… Œ“‰… ÃœÌœ… «Ê »‰ﬂ ÃœÌœ..."
    Me.Lbl(5).Caption = Msg
    '----------------------------
    Set xAccCol = New Collection
    '----------------------------
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Dcombos = New ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        Dcombos.GetAccountingCodes Me.DboParentAccount
    Else
        Dcombos.GetAccountingCodes Me.DboParentAccount
    End If
  
'    Me.Height = 8605
'    Me.Width = 16000
    Resize_Form Me
    Set GrdBack = New ClsBackGroundPic

    With Me.FgAccounts
        Set .WallPaper = GrdBack.Picture
        .GridLines = flexGridNone
        .AutoSize 0, .Cols - 1, False

        If SystemOptions.UserInterface = ArabicInterface Then
            .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

    End With

    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    LoadData
    Me.TxtModflg.text = "R"
    Me.Retrive "r"
    Me.TrvAccounts.Nodes("r").EnsureVisible
    Me.TrvAccounts.Nodes("r").Expanded = True
    Me.TrvAccounts.Nodes("r").Selected = True
    Dim StrSQL As String
    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

    StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
    fill_combo Me.DcActivityType, StrSQL

Set ISButton1.ButtonImage = mdifrmmain.ImgLstTree.ListImages("GridOptions").Picture

Dim StrLogFileName  As String
StrLogFileName = App.path & "\Titles\" & Me.Name & ".txt"
    If Dir(StrLogFileName) <> "" Then
          ShowFormtitles Me
    End If
    
    
    'If OPEN_NEW_SCREEN = True Then
    'Cmd_Click (0)
    'End If

End Sub

Private Sub LoadData()
    ModTree.LoadTreeAccount Me.TrvAccounts
End Sub

Public Sub Retrive(StrAccountCode As String, _
                   Optional BolPutInCol As Boolean = True)

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Static IntIndexPut As Integer

    If BolPutInCol = True Then
        IntIndexPut = IntIndexPut + 1
        xAccCol.Add StrAccountCode, "A" & IntIndexPut
        IntCurrentIndex = IntIndexPut
    Else
        'IntCurrentIndex = IntIndexPut
    End If

    'IntIndexPut = IntIndexPut + 1
    'xAccCol.Add StrAccountCode, "A" & IntIndexPut
    'IntCurrentIndex = IntIndexPut
    
    StrSQL = "Select * From Accounts Where Account_Code='" & StrAccountCode & "'"
   ' StrSQL = StrSQL & " and Account_Code in (SELECT     Account_Code"
   ' StrSQL = StrSQL & " From dbo.TblAccountBranch"
   ' StrSQL = StrSQL & " WHERE    BranchID  in(" & Current_branchSql & ") )"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.TxtAccount_ID.text = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
        Me.TxtAccount_Code.text = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        Me.TxtAccount_Serial.text = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
        Me.TxtAccount_Name.text = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
        Me.TxtAccount_NameE.text = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbllevel.Caption = "«·„” ÊÏ : " & CountAs(Me.TxtAccount_Code.text)
        Else
            Me.lbllevel.Caption = "Level:" & CountAs(Me.TxtAccount_Code.text)
        End If

If SystemOptions.UserInterface = ArabicInterface Then
    Me.lbllevel.Caption = "«·„” ÊÏ : " & CountAs2(Me.TxtAccount_Code.text, , LevelMode_FromParents_Cache, , dictParents)
Else
    Me.lbllevel.Caption = "Level: " & CountAs2(Me.TxtAccount_Code.text, , LevelMode_FromParents_Cache, , dictParents)
End If

        Me.Check1.value = IIf(rs("mowazna").value = True, vbChecked, Unchecked)
        Me.Check2.value = IIf(rs("cost_center").value = True, vbChecked, Unchecked)
        Me.ChKBlock.value = IIf(rs("Block").value = True, vbChecked, Unchecked)
        DcAccountTypes.ListIndex = IIf(IsNull(rs("AccountTypes").value), -1, rs("AccountTypes").value)
        Me.DcAccountTab.ListIndex = IIf(IsNull(rs("AccountTab").value), -1, rs("AccountTab").value)
    
        If rs("DepitOrCredit").value = 0 Then
            DepitOrCredit(0).value = True
        ElseIf rs("DepitOrCredit").value = 1 Then
            DepitOrCredit(1).value = True
        Else
            DepitOrCredit(0).value = False
            DepitOrCredit(1).value = False
      
        End If
     
        If rs("Differenttype").value = 0 Then
            Differenttype(0).value = True
        ElseIf rs("Differenttype").value = 1 Then
            Differenttype(1).value = True
        Else
            Differenttype(0).value = False
            Differenttype(1).value = False
      
        End If
     
        Me.DcUserGroup.BoundText = ""
        Me.DcUser.BoundText = ""
    
        If rs("Authority").value = 0 Then
            Authority(0).value = True
           
        ElseIf rs("Authority").value = 1 Then
            Me.DcUserGroup.BoundText = IIf(IsNull(rs("UserGroupid").value), "", rs("UserGroupid").value)
            Authority(1).value = True
        ElseIf rs("Authority").value = 2 Then
            Authority(2).value = True
            Me.DcUser.BoundText = IIf(IsNull(rs("Userid").value), "", rs("Userid").value)
        Else
            Authority(0).value = False
            Authority(1).value = False
            Authority(2).value = False
      
        End If
     
        If IsNull(rs("cost_center_type").value) Then
            Option2.value = True
            
        Else

            If rs("cost_center_type").value = 0 Then
                Option2.value = True
            ElseIf rs("cost_center_type").value = 1 Then
                Option1.value = True
            End If

            Me.DcCostCenter.BoundText = IIf(IsNull(rs("cost_center_id").value), "", rs("cost_center_id").value)
        End If
     
        Me.Check3.value = IIf(rs("Sum_account").value = True, vbChecked, Unchecked)
     
        Me.DCCURRENCY.BoundText = IIf(IsNull(rs("currenct_code").value), "", rs("currenct_code").value)
        Me.DcActivityType.BoundText = IIf(IsNull(rs("ActivityTypeId").value), "", rs("ActivityTypeId").value)
   
        If rs("last_account").value = True Then
            Me.OptAccountType(0).value = True
            Me.OptAccountType(1).value = False
        Else
            Me.OptAccountType(0).value = False
            Me.OptAccountType(1).value = True
        End If

        Me.DboParentAccount.BoundText = IIf(IsNull(rs("Parent_Account_Code").value), "", rs("Parent_Account_Code").value)
    
    End If

    StrSQL = "Select * From Accounts Where Parent_Account_Code='" & StrAccountCode & "'"
    StrSQL = StrSQL & " and Account_Code in (SELECT     Account_Code"
    StrSQL = StrSQL & " From dbo.TblAccountBranch"
    StrSQL = StrSQL & " WHERE    BranchID  in(" & Current_branchSql & ") )"
    StrSQL = StrSQL + " Order By Accounts.last_account, Account_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim FirstPeriod As Date
    Dim AccountCode As String
    Dim CurrentAccount As String
    Dim opening_balance As Double
    getFirstPeriodDateInthisYear FirstPeriod

    With Me.FgAccounts
        .Clear flexClearScrollable, flexClearEverything

        If Not (rs.BOF Or rs.EOF) Then
            .rows = .FixedRows + rs.RecordCount

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Account_ID")) = IIf(IsNull(rs("Account_ID").value), "", rs("Account_ID").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                AccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                End If
            
                If Not IsNull(rs("DateCreated").value) Then
                    .TextMatrix(i, .ColIndex("DateCreated")) = DisplayDate(rs("DateCreated").value)
                End If
  
                CurrentAccount = Abs(val(GetActualAccountBalance(AccountCode, branch_id, FirstPeriod, Date, , False, opening_balance)))

                .TextMatrix(i, .ColIndex("CurrentAccount")) = FormatNumber(CurrentAccount, SystemOptions.SysDefCurrencyForamt, True, True, True)
               
                If CurrentAccount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountState")) = "„œÌ‰"
                    Else
                        .TextMatrix(i, .ColIndex("AccountState")) = "Debit"
                    End If
                       
                ElseIf CurrentAccount < 0 Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountState")) = "œ«∆‰"
                    Else
                        .TextMatrix(i, .ColIndex("AccountState")) = "Credit"
                    End If
                               
                End If
            
                .TextMatrix(i, .ColIndex("OpenAccount")) = FormatNumber(Abs(opening_balance), SystemOptions.SysDefCurrencyForamt, True, True, True)

                If opening_balance > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("OpenAccountType")) = "„œÌ‰"
                    Else
                        .TextMatrix(i, .ColIndex("OpenAccountType")) = "Debit"
                    End If
                       
                ElseIf opening_balance < 0 Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("OpenAccountType")) = "œ«∆‰"
                    Else
                        .TextMatrix(i, .ColIndex("OpenAccountType")) = "Credit"
                    End If
                               
                End If
          
                If rs("last_account").value = True Then
                    .cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Item").ExtractIcon
                    .cell(flexcpFontBold, i, .ColIndex("Account_Name")) = False
                
                Else
                    .cell(flexcpPicture, i, .ColIndex("Account_ID")) = Me.ImgLstChartTree.ListImages("Closed_Node").ExtractIcon
                    .cell(flexcpFontBold, i, .ColIndex("Account_Name")) = True
                    .cell(flexcpFontName, i, .ColIndex("Account_Name")) = "Tahoma"
                End If

                rs.MoveNext
            Next i

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignRightCenter
        End If

        .AutoSize 0, .Cols - 1, False
    End With
FillMylistData StrAccountCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModflg.text <> "R" Then

        Select Case Me.TxtModflg.text

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                SaveData

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    Do While xAccCol.count > 0
        xAccCol.Remove xAccCol.count
    Loop
 Set dictParents = Nothing
    Set xAccCol = Nothing
    Set Dcombos = Nothing
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
Dim X As Integer
Dim StrLogFileName As String
If SystemOptions.UserInterface = ArabicInterface Then
X = MsgBox("Â·  —Ìœ    ⁄œÌ· «·⁄‰«ÊÌ‰ ", vbInformation + vbYesNoCancel)
Else
X = MsgBox("Change Title  yes/no", vbInformation + vbYesNoCancel)
End If



StrLogFileName = App.path & "\Titles\" & Me.Name & ".txt"
    If Dir(StrLogFileName) = "" Then
             Exit Sub
    End If
    
       If X = vbYes Then
             
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        ElseIf X = vbNo Then
                ShowFormtitles Me
                
        End If
        
End Sub

Private Sub Label13_Click()
Dim i As Integer
ListGroupSelected.Clear
For i = 0 To ListGroupAll.ListCount - 1
ListGroupSelected.AddItem ListGroupAll.List(i)
ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
Next i
End Sub

Private Sub Label14_Click()
If ListGroupAll.ListIndex = -1 Then Exit Sub
ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
End Sub

Private Sub Label22_Click()
Dim i As Integer
ListUserSelect.Clear
For i = 0 To ListAllUser.ListCount - 1
ListUserSelect.AddItem ListAllUser.List(i)
ListUserSelect.ItemData(i) = ListAllUser.ItemData(i)
Next i
End Sub

Private Sub Label3_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
 
    sql = " SELECT   * "
    sql = sql & " from ACCOUNTS order by Account_ID"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If mId(rs("Account_Serial").value, 1, 2) = "00" Then
                rs("Account_Serial").value = mId(rs("Account_Serial").value, 3, Len(rs("Account_Serial").value) - 2)
                rs.update
            End If

            rs.MoveNext
        Next i
 
    End If

    MsgBox "Done"
End Sub

Private Sub Label31_Click()
Frame8.Visible = False
End Sub

Private Sub Label5_Click()
ListUserSelect.Clear
End Sub

Private Sub Label6_Click()
If ListUserSelect.ListIndex > -1 Then
ListUserSelect.RemoveItem ListUserSelect.ListIndex
End If
End Sub

Private Sub Label7_Click()
If ListGroupSelected.ListIndex > -1 Then
ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
End If
End Sub

Private Sub Label8_Click()
ListGroupSelected.Clear
End Sub

Private Sub LblSelect_Click()
If ListAllUser.ListIndex = -1 Then Exit Sub
ListUserSelect.AddItem ListAllUser.List(ListAllUser.ListIndex)
ListUserSelect.ItemData(ListUserSelect.NewIndex) = ListAllUser.ItemData(ListAllUser.ListIndex)
End Sub

Private Sub Option1_Click()

    If Option1.value = True Then
        DcCostCenter.Enabled = True
    Else
        DcCostCenter.Enabled = False
    End If

End Sub

Private Sub Option2_Click()

    If Option1.value = True Then
        DcCostCenter.Enabled = True
    Else
        DcCostCenter.Enabled = False
    End If

End Sub

Private Sub Text1_Change()
If Me.Text1.text = "Alex2025" Then
Command2.Visible = True
cmdRenew.Visible = True
Else
Command2.Visible = False
cmdRenew.Visible = False
End If

End Sub

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
        
    If Not Node Is Nothing Then
        If InStr(1, Node.key, "G", vbTextCompare) <> 0 Then
            StrTemp = Node.key
            StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
        Else
            StrTemp = Node.key
        End If

        If Me.TxtModflg.text = "R" Then
            Me.Retrive StrTemp
            
        Else
            Me.DboParentAccount.BoundText = StrTemp
        End If
    End If

End Sub
 
Private Sub TxtAccount_Name_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtAccount_NameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtModFlg_Change()

    Select Case TxtModflg.text

        Case "N"
            Me.TxtAccount_ID.Enabled = False
            Me.TxtAccount_Serial.Enabled = True
            Me.TxtAccount_Name.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False
        
        Case "E"
            Me.TxtAccount_ID.Enabled = False
            Me.TxtAccount_Serial.Enabled = True
            Me.TxtAccount_Name.Enabled = True
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(5).Enabled = False
            Cmd(7).Enabled = False

        Case "R"
            Me.TxtAccount_ID.Enabled = False
            Me.TxtAccount_Serial.Enabled = False
            Me.TxtAccount_Name.Enabled = False
        
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(4).Enabled = True
            Cmd(5).Enabled = True
            Cmd(7).Enabled = True
        
    End Select

End Sub

Private Sub MoveInCollection(IntDir As Integer)
    Dim StrTemp  As String

    If IntDir = 0 Then
        ' Õ—ﬂ ··√„«„
    
    ElseIf IntDir = 1 Then

        ' Õ—ﬂ ··Œ·›
        If IntCurrentIndex = 0 Then
            StrTemp = xAccCol.Item(1)
            IntCurrentIndex = 1
        Else
            StrTemp = xAccCol.Item(IntCurrentIndex - 1)
        End If

        Me.Retrive StrTemp, False
    End If

    'Set Buttons

End Sub

Private Sub GetUpLevel()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String

    If val(Me.TxtAccount_ID.text) <> 0 Then
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.text)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrTemp = IIf(IsNull(rs("Parent_Account_Code").value), "", rs("Parent_Account_Code").value)

            If Trim$(StrTemp) <> "" Then
                Me.Retrive StrTemp, True
            End If
        End If
    End If

End Sub

Private Sub DelAccount()

    Dim Msg As String
    Dim RsAcccounts As ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrNodeKey As String
    Dim XNode As MSComctlLib.Node
    Dim IntRes As Integer

    'If Not DoPremis(Do_Delete, "Frm_General_Journal", "«·œ·Ì· «·„Õ«”»Ï") Then Exit Sub
    On Error GoTo ErrTrap
    StrAccountCode = Trim(Me.TxtAccount_Code.text)

    'If left(StrAccountCode, 6) = "a1a2a3" Or left(StrAccountCode, 6) = "a1a2a1" Or left(StrAccountCode, 6) = "a2a3a1" Or left(StrAccountCode, 6) = "a3a1a4" Then
    '    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«»"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If
    If CheckDelAccount1(StrAccountCode) = False Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«» ·«‰Â Õ”«» —∆Ì”Ì ·œÌ… «»‰«¡"
        Else
            Msg = "Can't Delete this Account because it have Child Account"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If CheckDelAccount(StrAccountCode) = True Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "This Acccount will be Deleted " ' & Me.TxtAccount_ID.text
            Msg = Msg & CHR(13) & "Account Code :- " & Me.TxtAccount_Serial.text
            Msg = Msg & CHR(13) & "Account Name " & Me.TxtAccount_NameE.text
            Msg = Msg & CHR(13) & ""
            Msg = Msg & CHR(13) & "Sure you want delte ??"
          
        Else
          
            Msg = "”Ê› Ì „ Õ–› «·Õ”«» —ﬁ„:- " '& Me.TxtAccount_ID.text
            Msg = Msg & CHR(13) & " ﬂÊœ «Ê „”·”· «·Õ”«»:- " & Me.TxtAccount_Serial.text
            Msg = Msg & CHR(13) & "«”„ «·Õ”«» :- " & Me.TxtAccount_Name.text
            Msg = Msg & CHR(13) & ""
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ ⁄„·Ì… «·Õ–› ..øø"
        
        End If

        IntRes = MsgBox(Msg, vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbYesNo, App.Title)

        If IntRes = vbNo Then
            Exit Sub
        End If

        Set RsAcccounts = New ADODB.Recordset
        StrSQL = "select *  From  ACCOUNTS Where Account_Code = '" & StrAccountCode & "'"
        RsAcccounts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If RsAcccounts("last_account").value = True Then
            StrNodeKey = StrAccountCode & ""
        Else
            StrNodeKey = StrAccountCode & "G"
        End If
   Cn.Execute " delete from TblAccountBranch where Account_Code ='" & TxtAccount_Code.text & "'"
   Cn.Execute " delete from TblAccountUser where Account_Code ='" & TxtAccount_Code.text & "'"
   ListGroupSelected.Clear
  ListUserSelect.Clear
        CuurentLogdata ("D")
        RsAcccounts.delete
        RsAcccounts.Close
    
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Account was Deleted......."
        Else
            Msg = " „  ⁄„·Ì… «·Õ–› ...!"
        End If

        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Set XNode = Me.TrvAccounts.Nodes(StrNodeKey)
        Me.TrvAccounts.Nodes.Remove XNode.key
        Set RsAcccounts = Nothing
        'LoadData
    Else

        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Can't Delete this Account"
        Else
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«» ...!!"
        End If
 
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim rs                As ADODB.Recordset
    Dim StrSQL            As String
    Dim Msg               As String
    Dim BolLastAccount    As Boolean
    Dim StrNewAccountCode As String

    If Me.TxtAccount_Name.text = "" And Me.TxtAccount_NameE.text <> "" Then
        Me.TxtAccount_Name.text = Me.TxtAccount_NameE.text
    End If

    If Me.TxtAccount_Name.text <> "" And Me.TxtAccount_NameE.text = "" Then
        Me.TxtAccount_NameE.text = Me.TxtAccount_Name.text
    End If

    If Trim$(Me.TxtAccount_Name.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… «”„ «·Õ”«»"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = EnglishInterface Then

        If Trim$(Me.TxtAccount_NameE.text) = "" Then
            Msg = "Must Enter Account Name"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

    End If

    If Trim$(Me.DboParentAccount.BoundText) = "" Then
        '    If SystemOptions.UserInterface = EnglishInterface Then
        '        Msg = "Must Specify Parent Account"
        '    Else
        '        Msg = "ÌÃ»  ÕœÌœ «”„ «·Õ”«» «·—∆Ì”Ì «·–Ì ”Ê› Ì ›—⁄ „‰Â Â–« «·Õ”«»"
        '    End If
        '
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '        Exit Sub
    End If

    If Me.OptAccountType(0).value = True Then
        BolLastAccount = True
    Else
        BolLastAccount = False
    End If

    Dim cost_center_type As Integer
    Dim cost_center_id   As String
    Dim AccountTypes     As Integer
    Dim AccountTab       As Integer
    Dim DepitOrCreditv   As Integer
    Dim Differenttypev   As Integer
    Dim Authorityv       As Integer
 
    Dim UserGroupIdv     As Integer
    Dim UserIdv          As Integer
 
    Dim ChKBlockv        As Boolean
    Dim UserIdas         As Integer

    AccountTypes = DcAccountTypes.ListIndex
    AccountTab = DcAccountTab.ListIndex

    If DepitOrCredit(0).value = True Then
        DepitOrCreditv = 0
    ElseIf DepitOrCredit(1).value = True Then
        DepitOrCreditv = 1
    End If

    If Differenttype(0).value = True Then
        Differenttypev = 0
    ElseIf Differenttype(1).value = True Then
        Differenttypev = 1
    End If

    If Authority(0).value = True Then
        Authorityv = 0
    ElseIf Authority(1).value = True Then
        Authorityv = 1
    ElseIf Authority(2).value = True Then
        Authorityv = 2
    End If

    UserGroupIdv = val(Me.DcUserGroup.BoundText)
    UserIdv = val(Me.DcUser.BoundText)

    If Option1.value = True Then
        If DcCostCenter.BoundText = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Must Specify COST CENTER"
                       
            Else
                Msg = "ÌÃ»  ÕœÌœ «”„ „—ﬂ“ «· ﬂ·›…"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcCostCenter.SetFocus
            Sendkeys "{F4}"
    
            Exit Sub
        Else
            cost_center_id = DcCostCenter.BoundText
        End If

    End If

    If Option2.value = True Then
        cost_center_type = 0
    Else
        cost_center_type = 1
    End If

    If Me.TxtModflg.text = "N" Then
        Dim mowazna        As Boolean
        Dim cost_center    As Boolean
        Dim Sum_account    As Boolean
        Dim ActivityTypeId As Integer

        If Check1.value = vbChecked Then
            mowazna = 1
        Else
            mowazna = 0
        End If

        If Check2.value = vbChecked Then
            cost_center = 1
        Else
            cost_center = 0
        End If

        If Check3.value = vbChecked Then
            Sum_account = 1
        Else
            Sum_account = 0
        End If

        If ChKBlock.value = vbChecked Then
            ChKBlockv = True
        Else
            ChKBlockv = False
        End If

        StrNewAccountCode = ModAccounts.AddNewAccount(Me.DboParentAccount.BoundText, Trim$(Me.TxtAccount_Name.text), BolLastAccount, False, Trim$(Me.TxtAccount_NameE.text), DCCURRENCY.BoundText, mowazna, cost_center, Sum_account, , TxtAccount_Serial, cost_center_type, cost_center_id, val(Me.DcActivityType.BoundText), AccountTypes, AccountTab, DepitOrCreditv, Differenttypev, Authorityv, UserGroupIdv, UserIdv, ChKBlockv)
        SaveBransh_UserAccount StrNewAccountCode
        CuurentLogdata

        If StrNewAccountCode <> "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Saved"
            Else
                Msg = " „  ⁄„·Ì… «·Õ›Ÿ."
            End If
    
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.TxtModflg.text = "R"
        End If

    ElseIf Me.TxtModflg.text = "E" Then
        Cn.Execute " delete from TblAccountBranch where Account_Code ='" & TxtAccount_Code.text & "'"
        Cn.Execute " delete from TblAccountUser where Account_Code ='" & TxtAccount_Code.text & "'"
        StrSQL = "Select * From Accounts Where Account_ID=" & val(Me.TxtAccount_ID.text) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            '·Ê «‰ «·Õ”«» Õ”«» —∆Ì”Ì ÊÌ—Ìœ «‰ ÌÃ⁄·Â Õ”«» ‰Â«∆
            If rs("last_account").value = False And OptAccountType(0).value = True Then
                If GetAccountChilds(Me.TxtAccount_Code.text) > 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "This Account is Master Account  And have child "
                        Msg = Msg & CHR(13) & "Can't change to last Account"
                                               
                    Else
                                            
                        Msg = "Â–« «·Õ”«» Õ”«» —∆Ì”Ì ÊÌÕ ÊÏ ⁄·Ï Õ”«»«  „ ›—⁄… „‰Â"
                        Msg = Msg & CHR(13) & "Ê·« Ì„ﬂ‰  ⁄œÌ·Â ≈·Ï Õ”«» ‰Â«∆"
                    End If
                                               
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If

            ElseIf rs("last_account").value = True And Me.OptAccountType(1).value = True Then
                '«·Õ”«» «·„—«œ  ⁄œÌ·Â Õ”«» ‰Â«∆ ÊÌ—«œ  ÕÊÌ·Â ≈·Ï Õ”«» —∆Ì”Ï
            
            End If

            Dim currency_code As Integer

            If Not IsNull(TxtAccount_Code.text) Then
                If Me.DCCURRENCY.BoundText <> "" Then
                    currency_code = Me.DCCURRENCY.BoundText
                Else
                    currency_code = 1
                End If

                If Me.OptAccountType(0).value = True Then
                    BolLastAccount = True
                Else
                    BolLastAccount = False
                End If

                ModAccounts.EditAccount rs("Account_Code").value, Me.TxtAccount_Name.text, Me.TxtAccount_NameE.text, Check1.value, Check2.value, currency_code, Check3.value, TxtAccount_Serial, cost_center_type, cost_center_id, val(Me.DcActivityType.BoundText), AccountTypes, AccountTab, DepitOrCreditv, Differenttypev, Authorityv, UserGroupIdv, UserIdv, ChKBlock, BolLastAccount
                SaveBransh_UserAccount rs("Account_Code").value
                CuurentLogdata

                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Saved"
                Else
                  
                    Msg = " „  ⁄„·Ì… «·Õ›Ÿ."
                End If
                  
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.TxtModflg.text = "R"
            End If

            '  End If
        End If
    End If

    Dcombos.GetAccountingCodes Me.DboParentAccount
    LoadData
    Me.Retrive StrNewAccountCode
End Sub

Public Function GetAccountChilds(StrAccountCode As String) As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From Accounts Where Parent_Account_Code = '" & StrAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        GetAccountChilds = 0
    ElseIf rs.RecordCount = 0 Then
        GetAccountChilds = 0
    Else
        GetAccountChilds = rs.RecordCount
    End If

End Function
