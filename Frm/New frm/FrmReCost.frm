VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmReCost 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10065
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   17475
   Icon            =   "FrmReCost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   17475
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic12 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17475
      _cx             =   30824
      _cy             =   17754
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
      AutoSizeChildren=   7
      BorderWidth     =   6
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
         Height          =   9900
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   17415
         _cx             =   30718
         _cy             =   17462
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
         Caption         =   "«⁄«œ… «Õ ”«» «· ﬂ·›…|÷»ÿ «· ﬂ·›…|New Tab"
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9525
            Index           =   2
            Left            =   -17970
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   16801
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   9525
               Left            =   0
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   0
               Width           =   17325
               _cx             =   30559
               _cy             =   16801
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
               Begin VB.Frame FraHeader 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   630
                  Index           =   0
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   0
                  Width           =   17355
                  Begin VB.TextBox TxtModFlg 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H0000FF00&
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   0
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Text            =   "modflag"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   465
                  End
                  Begin ImpulseButton.ISButton btnLast 
                     Height          =   315
                     Index           =   0
                     Left            =   450
                     TabIndex        =   20
                     Top             =   240
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmReCost.frx":6852
                     ColorButton     =   16777215
                     AcclimateGrayTones=   -1  'True
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnNext 
                     Height          =   315
                     Index           =   0
                     Left            =   915
                     TabIndex        =   21
                     Top             =   240
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmReCost.frx":6BEC
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnPrevious 
                     Height          =   315
                     Index           =   0
                     Left            =   1515
                     TabIndex        =   22
                     Top             =   240
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmReCost.frx":6F86
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnFirst 
                     Height          =   315
                     Index           =   0
                     Left            =   2040
                     TabIndex        =   23
                     Top             =   240
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmReCost.frx":7320
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "≈⁄«œ… ≈Õ ”«» «· ﬂ·›…"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   18
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   495
                     Index           =   2
                     Left            =   8880
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   120
                     Width           =   4080
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   495
                  Left            =   5460
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   630
                  Width           =   10905
                  _cx             =   19235
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
                  Begin VB.TextBox TxtSerial1 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   0
                     Left            =   8850
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   120
                     Width           =   1410
                  End
                  Begin MSComCtl2.DTPicker XPDtbTrans 
                     Height          =   300
                     Index           =   0
                     Left            =   7095
                     TabIndex        =   26
                     Top             =   165
                     Width           =   990
                     _ExtentX        =   1746
                     _ExtentY        =   529
                     _Version        =   393216
                     Format          =   94371841
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo Dcbranch 
                     Bindings        =   "FrmReCost.frx":76BA
                     Height          =   315
                     Index           =   0
                     Left            =   2460
                     TabIndex        =   27
                     Top             =   165
                     Width           =   3915
                     _ExtentX        =   6906
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
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
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· «—ÌŒ"
                     Height          =   195
                     Index           =   2
                     Left            =   8265
                     TabIndex        =   30
                     Top             =   225
                     Width           =   570
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                     Height          =   195
                     Index           =   4
                     Left            =   10155
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   165
                     Width           =   570
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·›—⁄"
                     Height          =   195
                     Index           =   7
                     Left            =   6225
                     TabIndex        =   28
                     Top             =   165
                     Width           =   990
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   750
                  Left            =   0
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   1110
                  Width           =   16365
                  _cx             =   28866
                  _cy             =   1323
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
                  Begin VB.TextBox TxtRemarks 
                     Alignment       =   1  'Right Justify
                     Height          =   630
                     Index           =   0
                     Left            =   4830
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   32
                     TabStop         =   0   'False
                     Top             =   75
                     Width           =   10605
                  End
                  Begin ImpulseButton.ISButton ShowBtn 
                     Height          =   630
                     Left            =   120
                     TabIndex        =   33
                     Top             =   75
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   1111
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
                     ButtonImage     =   "FrmReCost.frx":76CF
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin MSComCtl2.DTPicker FrmDate 
                     Height          =   285
                     Index           =   0
                     Left            =   2040
                     TabIndex        =   34
                     Top             =   45
                     Width           =   1470
                     _ExtentX        =   2593
                     _ExtentY        =   503
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   94371841
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker ToDate 
                     Height          =   285
                     Index           =   0
                     Left            =   2040
                     TabIndex        =   35
                     Top             =   405
                     Width           =   1470
                     _ExtentX        =   2593
                     _ExtentY        =   503
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   94371841
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ï  «—ÌŒ"
                     Height          =   195
                     Index           =   1
                     Left            =   3780
                     TabIndex        =   38
                     Top             =   405
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„‰  «—ÌŒ"
                     Height          =   195
                     Index           =   0
                     Left            =   3780
                     TabIndex        =   37
                     Top             =   75
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     Height          =   525
                     Index           =   21
                     Left            =   15255
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   330
                     Width           =   1320
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic5 
                  Height          =   840
                  Left            =   210
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Top             =   8310
                  Width           =   16155
                  _cx             =   28496
                  _cy             =   1482
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
                  Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                     Height          =   690
                     Index           =   0
                     Left            =   240
                     TabIndex        =   40
                     TabStop         =   0   'False
                     Top             =   75
                     Width           =   5160
                     _cx             =   9102
                     _cy             =   1217
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
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”Ã· «·Õ«·Ì:"
                        Height          =   360
                        Index           =   0
                        Left            =   4470
                        RightToLeft     =   -1  'True
                        TabIndex        =   44
                        Top             =   285
                        Width           =   600
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "⁄œœ «·”Ã·« :"
                        Height          =   360
                        Index           =   1
                        Left            =   630
                        RightToLeft     =   -1  'True
                        TabIndex        =   43
                        Top             =   285
                        Width           =   630
                     End
                     Begin VB.Label LabCurrRec 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        ForeColor       =   &H00800000&
                        Height          =   345
                        Index           =   0
                        Left            =   3960
                        RightToLeft     =   -1  'True
                        TabIndex        =   42
                        Top             =   300
                        Width           =   420
                     End
                     Begin VB.Label LabCountRec 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        ForeColor       =   &H00C00000&
                        Height          =   360
                        Index           =   0
                        Left            =   60
                        RightToLeft     =   -1  'True
                        TabIndex        =   41
                        Top             =   285
                        Width           =   480
                     End
                  End
                  Begin MSDataListLib.DataCombo DCboUserName 
                     Height          =   315
                     Index           =   1
                     Left            =   7500
                     TabIndex        =   45
                     Top             =   75
                     Width           =   6315
                     _ExtentX        =   11139
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õ—— »Ê«”ÿ…  "
                     Height          =   360
                     Index           =   8
                     Left            =   13995
                     TabIndex        =   46
                     Top             =   75
                     Width           =   930
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                  Height          =   420
                  Index           =   0
                  Left            =   0
                  TabIndex        =   47
                  TabStop         =   0   'False
                  Top             =   9105
                  Width           =   17325
                  _cx             =   30559
                  _cy             =   741
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
                  Align           =   2
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
                  Begin ImpulseButton.ISButton btnNew 
                     Height          =   225
                     Index           =   0
                     Left            =   12585
                     TabIndex        =   48
                     ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
                     Top             =   75
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":DF31
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnSave 
                     Height          =   225
                     Index           =   0
                     Left            =   9375
                     TabIndex        =   49
                     ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
                     Top             =   75
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":14793
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnModify 
                     Height          =   225
                     Index           =   0
                     Left            =   11055
                     TabIndex        =   50
                     ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
                     Top             =   75
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":14B2D
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton BtnUndo 
                     Height          =   225
                     Index           =   0
                     Left            =   7680
                     TabIndex        =   51
                     ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
                     Top             =   75
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":1B38F
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnDelete 
                     Height          =   225
                     Index           =   0
                     Left            =   6060
                     TabIndex        =   52
                     ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
                     Top             =   75
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":1B729
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnCancel 
                     Height          =   225
                     Index           =   0
                     Left            =   1650
                     TabIndex        =   53
                     ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
                     Top             =   75
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":1BCC3
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton ISButton5 
                     Height          =   285
                     Index           =   0
                     Left            =   4860
                     TabIndex        =   54
                     TabStop         =   0   'False
                     ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   990
                     _ExtentX        =   1746
                     _ExtentY        =   503
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "ÿ»«⁄… "
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
                     ButtonImage     =   "FrmReCost.frx":1C05D
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton ISButton8 
                     Height          =   225
                     Index           =   0
                     Left            =   3300
                     TabIndex        =   55
                     TabStop         =   0   'False
                     ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   397
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
                     ButtonImage     =   "FrmReCost.frx":228BF
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                  Height          =   6225
                  Left            =   0
                  TabIndex        =   56
                  TabStop         =   0   'False
                  Top             =   1965
                  Width           =   16365
                  _cx             =   28866
                  _cy             =   10980
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
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   5070
                     Left            =   0
                     TabIndex        =   57
                     Top             =   135
                     Width           =   16335
                     _cx             =   28813
                     _cy             =   8943
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
                     AllowBigSelection=   -1  'True
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   1
                     Cols            =   17
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmReCost.frx":22C59
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
                     WallPaperAlignment=   0
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   210
                     Index           =   0
                     Left            =   14805
                     TabIndex        =   58
                     Top             =   5295
                     Visible         =   0   'False
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   370
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   " Õ–› ”ÿ—"
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
                     ButtonImage     =   "FrmReCost.frx":22EDD
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   795
                     Index           =   11
                     Left            =   0
                     TabIndex        =   59
                     TabStop         =   0   'False
                     Top             =   5250
                     Width           =   13965
                     _cx             =   24633
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
                     Begin VB.CommandButton Command9 
                        Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                        Height          =   465
                        Index           =   0
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   64
                        Top             =   135
                        Width           =   2205
                     End
                     Begin VB.TextBox TxtNoteSerial 
                        Alignment       =   1  'Right Justify
                        Enabled         =   0   'False
                        Height          =   465
                        Index           =   0
                        Left            =   2550
                        Locked          =   -1  'True
                        RightToLeft     =   -1  'True
                        TabIndex        =   63
                        Top             =   135
                        Width           =   3255
                     End
                     Begin VB.TextBox TxtNoteID 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Index           =   0
                        Left            =   8295
                        RightToLeft     =   -1  'True
                        TabIndex        =   62
                        Top             =   -105
                        Visible         =   0   'False
                        Width           =   2070
                     End
                     Begin VB.CommandButton Command5 
                        Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
                        Height          =   465
                        Index           =   0
                        Left            =   8565
                        RightToLeft     =   -1  'True
                        TabIndex        =   61
                        Top             =   135
                        Width           =   1710
                     End
                     Begin VB.CommandButton Command2 
                        Caption         =   "Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
                        Height          =   465
                        Index           =   0
                        Left            =   6765
                        RightToLeft     =   -1  'True
                        TabIndex        =   60
                        Top             =   135
                        Width           =   1710
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "—ﬁ„ «·ﬁÌœ"
                        Height          =   390
                        Index           =   35
                        Left            =   5535
                        RightToLeft     =   -1  'True
                        TabIndex        =   65
                        Top             =   255
                        Width           =   1125
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9525
            Index           =   3
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   16801
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
            Begin VB.CheckBox chkNewMethode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÿ—Ìﬁ… «·ÃœÌœ…"
               Height          =   465
               Left            =   7980
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   1140
               Width           =   1365
            End
            Begin VB.CheckBox chkIsMov 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ÕÊÌ·«  ›ﬁÿ"
               Height          =   465
               Left            =   8010
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   780
               Width           =   1365
            End
            Begin VB.CommandButton exportHeader 
               Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
               Height          =   375
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   3720
               Width           =   1488
            End
            Begin VB.CommandButton ExportMe 
               Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
               Height          =   375
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   9120
               Width           =   1488
            End
            Begin VB.TextBox TxtAttachedItemCode 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7530
               TabIndex        =   108
               Top             =   495
               Width           =   1500
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   630
               Index           =   1
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   900
               Width           =   5115
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Õ–› ﬁÌœ  "
               Height          =   465
               Index           =   1
               Left            =   6645
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   7620
               Visible         =   0   'False
               Width           =   1710
            End
            Begin VB.CommandButton Command5 
               Caption         =   "÷»ÿ «·ﬁÌœ "
               Height          =   465
               Index           =   1
               Left            =   8445
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   7620
               Width           =   1710
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   1
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   7380
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   465
               Index           =   1
               Left            =   2430
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   7620
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   465
               Index           =   1
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   7620
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.TextBox TxtSerial1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   14640
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   540
               Width           =   1260
            End
            Begin VB.Frame FraHeader 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   630
               Index           =   1
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   -210
               Width           =   17355
               Begin VB.CheckBox chkSalim 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ﬂ·›… ﬂ„Ì…"
                  Height          =   255
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000FF00&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Text            =   "modflag"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   465
               End
               Begin ImpulseButton.ISButton btnLast 
                  Height          =   315
                  Index           =   1
                  Left            =   450
                  TabIndex        =   80
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":23477
                  ColorButton     =   16777215
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnNext 
                  Height          =   315
                  Index           =   1
                  Left            =   915
                  TabIndex        =   81
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":23811
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnPrevious 
                  Height          =   315
                  Index           =   1
                  Left            =   1515
                  TabIndex        =   82
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":23BAB
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnFirst 
                  Height          =   315
                  Index           =   1
                  Left            =   2040
                  TabIndex        =   83
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":23F45
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷»ÿ «· ﬂ·›…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   0
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   120
                  Width           =   4080
               End
            End
            Begin VB.CommandButton cmdRecost 
               Caption         =   " ﬂ·Ì›"
               Height          =   315
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   4020
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   645
               Index           =   14
               Left            =   9345
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   915
               Width           =   6270
               Begin MSComCtl2.DTPicker FrmDate 
                  Height          =   330
                  Index           =   1
                  Left            =   3240
                  TabIndex        =   5
                  Top             =   210
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   94371843
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   330
                  Index           =   1
                  Left            =   210
                  TabIndex        =   6
                  Top             =   210
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   94371843
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   29
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   27
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   240
                  Width           =   420
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid grdMaster 
               Height          =   1980
               Left            =   -90
               TabIndex        =   9
               Top             =   1560
               Width           =   16575
               _cx             =   29236
               _cy             =   3492
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReCost.frx":242DF
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
               WallPaperAlignment=   0
               AccessibleName  =   "ReCostDet"
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin ImpulseButton.ISButton Btn_Update 
               Height          =   255
               Index           =   0
               Left            =   10095
               TabIndex        =   10
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   8865
               Visible         =   0   'False
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   450
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
               BackColor       =   14871017
               FontSize        =   9.75
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmReCost.frx":246B1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DCboUserName 
               Height          =   315
               Index           =   0
               Left            =   11745
               TabIndex        =   11
               Top             =   8805
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmReCost.frx":24A4B
               Height          =   315
               Index           =   1
               Left            =   30
               TabIndex        =   12
               Top             =   525
               Width           =   2340
               _ExtentX        =   4128
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
            Begin XtremeSuiteControls.PushButton cmdInsert 
               Height          =   600
               Left            =   6210
               TabIndex        =   13
               Top             =   930
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   1058
               _StockProps     =   79
               Caption         =   "«œ—«Ã «·”‰œ«  "
               UseVisualStyle  =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid grdDet 
               Height          =   3120
               Left            =   0
               TabIndex        =   67
               Top             =   4380
               Width           =   16455
               _cx             =   29025
               _cy             =   5503
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
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   31
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReCost.frx":24A60
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   420
               Index           =   1
               Left            =   0
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   9105
               Width           =   17325
               _cx             =   30559
               _cy             =   741
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
               Align           =   2
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
               Begin ImpulseButton.ISButton btnNew 
                  Height          =   225
                  Index           =   1
                  Left            =   12585
                  TabIndex        =   70
                  ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
                  Top             =   75
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":24F3C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnSave 
                  Height          =   225
                  Index           =   1
                  Left            =   9375
                  TabIndex        =   71
                  ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
                  Top             =   75
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":2B79E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   225
                  Index           =   1
                  Left            =   11055
                  TabIndex        =   72
                  ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
                  Top             =   75
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":2BB38
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton BtnUndo 
                  Height          =   225
                  Index           =   1
                  Left            =   7680
                  TabIndex        =   73
                  ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
                  Top             =   75
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":3239A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnDelete 
                  Height          =   225
                  Index           =   1
                  Left            =   6060
                  TabIndex        =   74
                  ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
                  Top             =   75
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":32734
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnCancel 
                  Height          =   225
                  Index           =   1
                  Left            =   1650
                  TabIndex        =   75
                  ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
                  Top             =   75
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":32CCE
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton ISButton5 
                  Height          =   285
                  Index           =   1
                  Left            =   4860
                  TabIndex        =   76
                  TabStop         =   0   'False
                  ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   503
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄… "
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
                  ButtonImage     =   "FrmReCost.frx":33068
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton ISButton8 
                  Height          =   225
                  Index           =   1
                  Left            =   3300
                  TabIndex        =   77
                  TabStop         =   0   'False
                  ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
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
                  ButtonImage     =   "FrmReCost.frx":398CA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   300
               Index           =   1
               Left            =   12720
               TabIndex        =   87
               Top             =   510
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   529
               _Version        =   393216
               Format          =   94371841
               CurrentDate     =   38784
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   630
               Index           =   1
               Left            =   330
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   8430
               Width           =   5160
               _cx             =   9102
               _cy             =   1111
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
               Begin VB.Label LabCountRec 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00C00000&
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   255
                  Width           =   480
               End
               Begin VB.Label LabCurrRec 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   1
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   270
                  Width           =   420
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·”Ã·« :"
                  Height          =   420
                  Index           =   3
                  Left            =   630
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   165
                  Width           =   630
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ã· «·Õ«·Ì:"
                  Height          =   420
                  Index           =   2
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   165
                  Width           =   600
               End
            End
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Left            =   10020
               TabIndex        =   102
               Top             =   510
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItemID1 
               Height          =   315
               Left            =   3030
               TabIndex        =   109
               Top             =   510
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComDlg.CommonDialog cd 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’‰›"
               Height          =   195
               Index           =   9
               Left            =   9120
               TabIndex        =   110
               Top             =   540
               Width           =   870
            End
            Begin VB.Label lblStatus2 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   10500
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   7710
               Visible         =   0   'False
               Width           =   5145
            End
            Begin VB.Label Status 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   5610
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   8190
               Width           =   7155
            End
            Begin VB.Label CountNo 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   12810
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   8190
               Width           =   3015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·„Œ“‰"
               Height          =   195
               Index           =   6
               Left            =   11865
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   525
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   525
               Index           =   5
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1080
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   525
               Index           =   3
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   0
               Width           =   1320
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   390
               Index           =   1
               Left            =   5415
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   7740
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒÂ"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   20
               Left            =   13755
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   555
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   195
               Index           =   23
               Left            =   2175
               TabIndex        =   16
               Top             =   555
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Õ—— »Ê«”ÿ…  "
               Height          =   195
               Index           =   24
               Left            =   15165
               TabIndex        =   15
               Top             =   8805
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”‰œ"
               Height          =   195
               Index           =   31
               Left            =   15885
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   555
               Width           =   1020
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9525
            Index           =   0
            Left            =   18060
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   16801
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
            Begin VB.Frame Frame1 
               Caption         =   "Frame1"
               Height          =   765
               Left            =   -240
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   6930
               Visible         =   0   'False
               Width           =   16575
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ê« Ì— „‘ —Ì«  „Œ«·›… ·”‰œ«  «” ·«„Â«"
                  Height          =   405
                  Index           =   5
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   180
                  Width           =   3015
               End
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ê« Ì— „‘ —Ì«  »œÊ‰ ”‰œ«  «” ·«„"
                  Height          =   405
                  Index           =   4
                  Left            =   2970
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   180
                  Width           =   2895
               End
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«  «·«” ·«„ »œÊ‰ ›Ê« Ì—"
                  Height          =   405
                  Index           =   3
                  Left            =   5580
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   180
                  Width           =   2415
               End
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ê« Ì— „»Ì⁄«  „Œ«·›… ·”‰œ«  ’—›Â«"
                  Height          =   405
                  Index           =   2
                  Left            =   7950
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   180
                  Width           =   2895
               End
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ê« Ì— „»Ì⁄«  »œÊ‰ ”‰œ«  ’—›"
                  Height          =   405
                  Index           =   1
                  Left            =   10830
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   180
                  Width           =   2895
               End
               Begin VB.OptionButton optOut 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«  «·’—› »œÊ‰ ›Ê« Ì—"
                  Height          =   405
                  Index           =   0
                  Left            =   13590
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   180
                  Width           =   2415
               End
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   645
               Index           =   0
               Left            =   9390
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   480
               Width           =   6270
               Begin MSComCtl2.DTPicker FrmDate 
                  Height          =   330
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   123
                  Top             =   210
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   94371843
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   330
                  Index           =   2
                  Left            =   210
                  TabIndex        =   124
                  Top             =   210
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   94371843
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   11
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   240
                  Width           =   420
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   10
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   240
                  Width           =   375
               End
            End
            Begin VB.Frame FraHeader 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   630
               Index           =   2
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   -210
               Width           =   17355
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000FF00&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Text            =   "modflag"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   465
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ﬂ·›… ﬂ„Ì…"
                  Height          =   255
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1335
               End
               Begin ImpulseButton.ISButton btnLast 
                  Height          =   315
                  Index           =   2
                  Left            =   450
                  TabIndex        =   117
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":39C64
                  ColorButton     =   16777215
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnNext 
                  Height          =   315
                  Index           =   2
                  Left            =   915
                  TabIndex        =   118
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":39FFE
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnPrevious 
                  Height          =   315
                  Index           =   2
                  Left            =   1515
                  TabIndex        =   119
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":3A398
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btnFirst 
                  Height          =   315
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   120
                  Top             =   240
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   16777215
                  FontSize        =   12
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmReCost.frx":3A732
                  ColorButton     =   16777215
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "÷»ÿ «· ﬂ·›…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   3
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   120
                  Width           =   4080
               End
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   630
               Index           =   2
               Left            =   6990
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   8850
               Visible         =   0   'False
               Width           =   6285
            End
            Begin VB.TextBox TxtAttachedItemCode2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7530
               TabIndex        =   112
               Top             =   495
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VSFlex8UCtl.VSFlexGrid grdMaster2 
               Height          =   6600
               Left            =   150
               TabIndex        =   127
               Top             =   1110
               Width           =   16575
               _cx             =   29236
               _cy             =   11642
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   27
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReCost.frx":3AACC
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
               WallPaperAlignment=   0
               AccessibleName  =   "ReCostDet"
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmReCost.frx":3AF21
               Height          =   315
               Index           =   2
               Left            =   30
               TabIndex        =   128
               Top             =   525
               Visible         =   0   'False
               Width           =   2340
               _ExtentX        =   4128
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
            Begin XtremeSuiteControls.PushButton cmdInsert2 
               Height          =   600
               Left            =   7470
               TabIndex        =   129
               Top             =   540
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   1058
               _StockProps     =   79
               Caption         =   "«œ—«Ã «·”‰œ«  "
               UseVisualStyle  =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   420
               Index           =   2
               Left            =   0
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   9105
               Width           =   17325
               _cx             =   30559
               _cy             =   741
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
               Align           =   2
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
            End
            Begin MSDataListLib.DataCombo DCboStoreName2 
               Height          =   315
               Left            =   10020
               TabIndex        =   131
               Top             =   510
               Visible         =   0   'False
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItemID2 
               Height          =   315
               Left            =   3030
               TabIndex        =   132
               Top             =   510
               Visible         =   0   'False
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   600
               Left            =   90
               TabIndex        =   148
               Top             =   8310
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   1058
               _StockProps     =   79
               Caption         =   "«‰‘«¡ «·”‰œ«  "
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   195
               Index           =   17
               Left            =   2175
               TabIndex        =   140
               Top             =   555
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   525
               Index           =   15
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   0
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   525
               Index           =   14
               Left            =   6150
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   1080
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·„Œ“‰"
               Height          =   195
               Index           =   13
               Left            =   11865
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   525
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label CountNo2 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   12810
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   8190
               Width           =   3015
            End
            Begin VB.Label Status2 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   5610
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   8190
               Width           =   7155
            End
            Begin VB.Label lblStatus22 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   10500
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   7710
               Visible         =   0   'False
               Width           =   5145
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’‰›"
               Height          =   195
               Index           =   12
               Left            =   9120
               TabIndex        =   133
               Top             =   540
               Visible         =   0   'False
               Width           =   870
            End
         End
      End
   End
End
Attribute VB_Name = "FrmReCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim RevenueAccount As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double
Public mIndex As Integer
Dim mTableName As String
Dim mTableName2 As String
Dim mTableName3 As String
Dim cProgress As ClsProgress
  Dim BolFrmLoaded As Boolean


Private Sub cmdInsert2_Click()
    FillGrid3
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DcboItemID1_LostFocus()
DcboItemID1_Validate False

End Sub

Private Sub DcboItemID1_Validate(Cancel As Boolean)
On Error Resume Next
If val(DcboItemID1.Tag) = val(DcboItemID1.BoundText) And val(DcboItemID1.BoundText) <> 0 Then Exit Sub
If val(DcboItemID1.BoundText) = 0 Then Exit Sub

Dim UnitID As Long
Dim UnitName As String

    Me.TxtAttachedItemCode.Text = GetItemCode(val(Me.DcboItemID1.BoundText))

If Me.TxtModFlg(mIndex).Text <> "R" Then
  GetDefaultItemUnit val(Me.DcboItemID1.BoundText), UnitID, UnitName
    

     
    'fillgrid
    
    
  End If

DcboItemID1.Tag = DcboItemID1.BoundText
End Sub



Private Sub DcboItemID1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmItemSearch.RetrunType = 2020
        FrmItemSearch.show vbModal
        
    End If
End Sub

Private Sub exportHeader_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\grdMaster.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
  '  Me.Grid.saveGrid StrFileName, flexFileExcel, True
  '  OpenFile StrFileName
    
         On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Report"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.grdMaster.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    

End Sub

Private Sub ExportMe_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\grdDet.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
  '  Me.Grid.saveGrid StrFileName, flexFileExcel, True
  '  OpenFile StrFileName
    
         On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.FileName = "Report"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.FileName & ".xls"
Me.grdDet.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    

End Sub

Private Sub PushButton1_Click()
   lblStatus22.Visible = True
    lblStatus22.Caption = "Ì „ ÷»ÿ «·ﬁÌœ »—Ã«¡ «·«‰ Ÿ«—"
    Status2.Caption = ""
    CountNo2.Caption = ""
CreateIssueVoucher
    Status2.Caption = ""
    CountNo2.Caption = ""
     lblStatus22.Visible = False
    
End Sub

Private Sub TxtAttachedItemCode_KeyDown(KeyCode As Integer, _
                                        Shift As Integer)
Dim UnitID As Long
Dim UnitName As String
    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode.Text = "" Then
            Me.DcboItemID1.BoundText = ""
        Else
            Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.Text))
            GetDefaultItemUnit val(Me.DcboItemID1.BoundText), UnitID, UnitName

        End If
    End If

End Sub
Private Sub DcboItemID1_Click(Area As Integer)
    
    DcboItemID1_Validate False
    
End Sub
Private Sub DisplayDetails()
'  Dim mTable As String
'  mTable = grdMaster.AccessibleName
'
'
'
'    grdMaster.Rows = 1
'    Set rsGrdArray(0) = OpenGridRs(s)
'
'    grdArray = Array(grdMaster.Name, grdDet.Name)
'    Static IsStart As Boolean
'    If IsStart Then Exit Sub
'    IsStart = True
'    'ValidateFormGrid Me, Array(rsGrd, rsGrd2), grdArray
'    ValidateFormGrid Me, Array(rsGrdArray(0), rsGrdArray(1)), grdArray
'
'    If DB_CreateTable("ReCostDet", True, "Id ", True) = True Then
'        DB_CreateField "ReCostDet", "ReCostDetID", adInteger, adColNullable, , , "    ", False, True
'        DB_CreateField "ReCostDet", "ID", adVarWChar, adColNullable, 15, "''", "    ", False, True
'        DB_CreateField "ReCostDet", "SerID", adVarWChar, adColNullable, 15, "''", "    ", False, True
'        DB_CreateField "ReCostDet", "Transaction_ID", adInteger, adColNullable, , , "    ", False, True
'    End If
'
'
'      If DB_CreateTable("ReCostDet2", True, "Id ", True) = True Then
'        DB_CreateField "ReCostDet2", "ReCostDetID", adInteger, adColNullable, , , "    ", False, True
'        DB_CreateField "ReCostDet2", "ID", adVarWChar, adColNullable, 15, "''", "    ", False, True
'        DB_CreateField "ReCostDet2", "SerID", adVarWChar, adColNullable, 15, "''", "    ", False, True
'        DB_CreateField "ReCostDet2", "Transaction_ID", adInteger, adColNullable, , , "    ", False, True
'        DB_CreateField "ReCostDet2", "Transaction_ID", adInteger, adColNullable, , , "    ", False, True
'        DB_CreateField "ReCostDet2", "NoteSerial1", adInteger, adColNullable, , , "    ", False, True
'        DB_CreateField "ReCostDet2", "Total", adDouble, adColNullable, , , "    ", False, True
'    End If
'
End Sub




Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg(mIndex).Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow2
 End Select
End If
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Long
    With FG
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("ItemName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub

Private Sub cmdInsert_Click()
'If Me.TxtModFlg(mIndex).Text <> "R" Then


    '    DB_CreateField "Transaction_Details", "OldcostPrice", adDouble, adColNullable, , , "    ", False, True
        
       
          
    FillGrid2
'End If
End Sub



Private Sub Command2_Click(Index As Integer)
If Me.TxtModFlg(mIndex).Text = "R" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID(mIndex).Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID(mIndex).Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update TblReCost set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1((mIndex)).Text) & " "
        RsSavRec.Requery
         FindRec val(TxtSerial1(1).Text)
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
End Sub

Private Sub Command5_Click(Index As Integer)
If TxtSerial1(mIndex).Text = "" Then
    MsgBox "·«Ì„ﬂ‰ ÷»ÿ «·ﬁÌœ ﬁ»· Õ›Ÿ «·Õ—ﬂ…"
    Exit Sub
Else

    Dim i As Long
'    Set cProgress = New ClsProgress
'    BolFrmLoaded = True
'    cProgress.ProgressType = Waiting
'    cProgress.StartProgress
    lblStatus2.Visible = True
    lblStatus2.Caption = "Ì „ ÷»ÿ «·ﬁÌœ »—Ã«¡ «·«‰ Ÿ«—"
    Status.Caption = ""
    CountNo.Caption = ""
    
    Dim s As String
   
        '=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & " And IsNull(AdvanceID,0) <> 0"
            
    
    For i = 1 To grdMaster.Rows - 1
      With grdMaster

        
'       If val(.TextMatrix(i, .ColIndex("DiffTotal"))) <> 0 And val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 19 Or val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 10 Or val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 992 Then
If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 19 Or val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 10 Or val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 992 Then

'                    s = "Delete DOUBLE_ENTREY_VOUCHERS Where Transaction_ID "
'            s = s & " In (Select  Transaction_ID From " & mTableName2 & " ) And IsNull(AdvanceID,0) <> 0"
'            s = s & " and Transaction_ID = " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
'            Cn.Execute s, , adExecuteNoRecords
            If chkNewMethode Then
                CreateVoucNew i
            Else
                CreateVouc i
            End If
            DoEvents
            
            
            CountNo.Caption = "”ÿ— —ﬁ„ " & i
            Status.Caption = "÷»ÿ «·Õ—ﬂ… —ﬁ„ " & .TextMatrix(i, .ColIndex("NoteSerial1"))
        End If
       End With
    Next
'    If BolFrmLoaded = True Then
'        cProgress.StopProgess
'        Set cProgress = Nothing
'    End If
End If
s = "Update TblReCostCalc Set  EntryCreated = 1 Where Id =" & val(TxtSerial1(mIndex))
Cn.Execute s, , adExecuteNoRecords
Command5(mIndex).Enabled = False
lblStatus2.Caption = " „ ÷»ÿ «·ﬁÌÊœ"
Status.Caption = ""
MsgBox " „ ÷»ÿ «·ﬁÌÊœ"
lblStatus2.Visible = True
End Sub

Private Sub Command9_Click(Index As Integer)
ShowGL_cc TxtNoteSerial(mIndex).Text, , 200
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    
    If mIndex = 1 Then
        mTableName = "TblReCostCalc"
        mTableName2 = "TblReCostCalcDet"
        mTableName3 = "TblReCostCalcDet2"
    ElseIf mIndex = 0 Then
        mTableName = "TblReCost"
        mTableName2 = "TblReCostDet"
        
    End If

 'Wael

    
     
    
    
    
      TabMain.TabVisible(1) = False
     TabMain.TabVisible(0) = False
     TabMain.TabVisible(2) = False
     TabMain.TabVisible(mIndex) = True
     TabMain.CurrTab = mIndex
    If mIndex <> 2 Then
        conection = "select * from " & mTableName & "  order by  ID "
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Me.TxtModFlg(mIndex).Text = "R"
    End If
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    
    
    Dcombos.GetBranches Me.Dcbranch(mIndex)
    Dcombos.GetUsers Me.DCboUserName(mIndex)
    Dcombos.GetStores Me.DCboStoreName
    
    Dcombos.GetItemsNames Me.DcboItemID1, 0
    
    
   ' BtnLast_Click
    FrmDate(0).value = Date
    ToDate(0).value = Date
    FrmDate(1).value = Date
    ToDate(1).value = Date
    
        FrmDate(2).value = Date
    ToDate(2).value = Date
    
    
    
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If mIndex <> 2 Then
    BtnLast_Click mIndex
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click mIndex
    End If
    
   Me.Refresh
ErrTrap:
End Sub



Public Sub FiLLRec()

  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg(mIndex).Text = "E" Then
                 StrSQL = "Delete From " & mTableName2 & "  Where ReCostID =" & val(TxtSerial1(mIndex).Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                If mIndex = 1 Then
                    StrSQL = "Delete From " & mTableName3 & "  Where ReCostID =" & val(TxtSerial1(mIndex).Text) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                End If
              End If
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch(mIndex).BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec.Fields("Remarks").value = TxtRemarks(mIndex).Text
    RsSavRec.Fields("UserID").value = val(DCboUserName(mIndex).BoundText)
    
    

    If mIndex = 1 Then
        RsSavRec.Fields("FrmDate").value = FrmDate(mIndex).value
        RsSavRec.Fields("ToDate").value = ToDate(mIndex).value
        RsSavRec.Fields("StoreID").value = val(Me.DCboStoreName.BoundText)
        RsSavRec.Fields("ItemID").value = val(Me.DcboItemID1.BoundText)
    End If







    RsSavRec.update

''//////////////////////////
    If mIndex = 0 Then
        saveDetails
    ElseIf mIndex = 1 Then
        saveDetails2
    End If
 

      Select Case Me.TxtModFlg(mIndex).Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then

                Me.Refresh
                FiLLTXT
                TxtModFlg(mIndex) = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else

                Me.Refresh
                FiLLTXT
                TxtModFlg(mIndex) = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                btnNew_Click mIndex
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg(mIndex) = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                Me.Refresh
                FiLLTXT
                TxtModFlg(mIndex) = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                Me.Refresh
                FiLLTXT
                TxtModFlg(mIndex) = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
Private Sub saveDetails()
    Set RsDevsub = New ADODB.Recordset
    Dim s As String, StrSQL As String
    StrSQL = "SELECT  *  from " & mTableName2 & "  Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Long
    With FG
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Item_ID"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("ReCostID").value = val(Me.TxtSerial1(mIndex).Text)
                 RsDevsub("IDRef").value = IIf((.TextMatrix(i, .ColIndex("IDRef"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDRef"))))
                 RsDevsub("Transaction_ID").value = IIf((.TextMatrix(i, .ColIndex("Transaction_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
                 RsDevsub("NoteSerial1").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial1"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial1")))
                 RsDevsub("Transaction_Date").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Null, .TextMatrix(i, .ColIndex("Transaction_Date")))
                 RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val(.TextMatrix(i, .ColIndex("StoreID"))))
                 RsDevsub("Item_ID").value = IIf((.TextMatrix(i, .ColIndex("Item_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Item_ID"))))
                 RsDevsub("UnitId").value = IIf((.TextMatrix(i, .ColIndex("UnitId"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitId"))))
                 RsDevsub("ShowQty").value = IIf((.TextMatrix(i, .ColIndex("ShowQty"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShowQty"))))
                 .TextMatrix(i, .ColIndex("showPrice")) = ModItemCostPrice.GetCostItemPrice(val(.TextMatrix(i, .ColIndex("Item_ID"))), 0, "", , SystemOptions.SysMainStockCostMethod, , , IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Date, .TextMatrix(i, .ColIndex("Transaction_Date"))), , val(.TextMatrix(i, .ColIndex("UnitId"))), val(.TextMatrix(i, .ColIndex("StoreID"))))
'               salimhere
                '.TextMatrix(i, .ColIndex("showPrice")) = ModItemCostPrice.GetCostItemPrice(val(.TextMatrix(i, .ColIndex("Item_ID"))), 0, "", , SystemOptions.SysMainStockCostMethod, , , IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Date, .TextMatrix(i, .ColIndex("Transaction_Date"))), , val(.TextMatrix(i, .ColIndex("UnitId"))))
                 RsDevsub("showPrice").value = IIf((.TextMatrix(i, .ColIndex("showPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("showPrice"))))
                 RsDevsub("Price").value = val(.TextMatrix(i, .ColIndex("showPrice")))
                 RsDevsub("Valu").value = val(.TextMatrix(i, .ColIndex("ShowQty"))) * val(.TextMatrix(i, .ColIndex("showPrice")))
               Dim RsUnitData As ADODB.Recordset
               StrSQL = "Select * From TblItemsUnits Where ItemID=" & val(.TextMatrix(i, .ColIndex("Item_ID")))
               StrSQL = StrSQL + " AND UnitID=" & val(.TextMatrix(i, .ColIndex("UnitId")))
               Set RsUnitData = New ADODB.Recordset
               RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
               RsDevsub("Quantity").value = RsUnitData("UnitFactor").value * val(.TextMatrix(i, .ColIndex("ShowQty")))
            End If
             RsDevsub("Price").value = val(IIf((.TextMatrix(i, FG.ColIndex("Price")) = ""), 0, val(.TextMatrix(i, .ColIndex("Price"))))) / RsUnitData("UnitFactor").value
             RsDevsub.update
            Cn.Execute "Update dbo.Transaction_Details set FlgReCost=1,ReCostID=" & val(TxtSerial1(mIndex).Text) & " where ID=" & val(.TextMatrix(i, .ColIndex("IDRef"))) & ""
      End If
     Next i
    End With
    

End Sub

Private Sub saveDetails2()
On Error Resume Next
    Set RsDevsub = New ADODB.Recordset
    Dim s As String, StrSQL As String
     s = "Delete " & mTableName2 & "  Where ReCostID =" & val(Me.TxtSerial1(mIndex).Text)
    Cn.Execute s, , adExecuteNoRecords
     s = "Delete " & mTableName3 & "  Where ReCostID =" & val(Me.TxtSerial1(mIndex).Text)
    Cn.Execute s, , adExecuteNoRecords
'     Set cProgress = New ClsProgress
'        BolFrmLoaded = True
'        cProgress.ProgressType = Waiting
'        cProgress.StartProgress


      StrSQL = "SELECT  *  from " & mTableName2 & "  Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Long
    With grdMaster
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Transaction_ID"))) <> 0 Then
                RsDevsub.AddNew
                 RsDevsub("ReCostID").value = val(Me.TxtSerial1(mIndex).Text)
                 RsDevsub("IDRef").value = IIf((.TextMatrix(i, .ColIndex("IDRef"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDRef"))))
                 RsDevsub("Transaction_ID").value = IIf((.TextMatrix(i, .ColIndex("Transaction_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
                 RsDevsub("NoteSerial1").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial1"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial1")))
                 RsDevsub("Transaction_Date").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Null, .TextMatrix(i, .ColIndex("Transaction_Date")))
                 RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val(.TextMatrix(i, .ColIndex("StoreID"))))
                 RsDevsub("StoreID2").value = IIf((.TextMatrix(i, .ColIndex("StoreID2"))) = "", Null, val(.TextMatrix(i, .ColIndex("StoreID2"))))
                 
                 RsDevsub("CusID").value = IIf((.TextMatrix(i, .ColIndex("CusID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CusID"))))
                 RsDevsub("Transaction_Type").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Type"))) = "", Null, val(.TextMatrix(i, .ColIndex("Transaction_Type"))))
                 RsDevsub("Note_Value").value = IIf((.TextMatrix(i, .ColIndex("Note_Value"))) = "", Null, val(.TextMatrix(i, .ColIndex("Note_Value"))))

                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("NoteID").value = IIf((.TextMatrix(i, .ColIndex("NoteID"))) = "", Null, val(.TextMatrix(i, .ColIndex("NoteID"))))
                RsDevsub("NoteSerial").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial"))) = "", Null, val(.TextMatrix(i, .ColIndex("NoteSerial"))))


                 RsDevsub("FixesAssetsID").value = IIf((.TextMatrix(i, .ColIndex("FixesAssetsID"))) = "", Null, val(.TextMatrix(i, .ColIndex("FixesAssetsID"))))
                 RsDevsub("Emp_ID").value = IIf((.TextMatrix(i, .ColIndex("Emp_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Emp_ID"))))
                 RsDevsub("DepartementID").value = IIf((.TextMatrix(i, .ColIndex("DepartementID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DepartementID"))))
                 RsDevsub("OldTotal").value = IIf((.TextMatrix(i, .ColIndex("OldTotal"))) = "", Null, val(.TextMatrix(i, .ColIndex("OldTotal"))))
                 RsDevsub("DiffTotal").value = IIf((.TextMatrix(i, .ColIndex("DiffTotal"))) = "", Null, val(.TextMatrix(i, .ColIndex("DiffTotal"))))
                 RsDevsub("Doctype").value = IIf((.TextMatrix(i, .ColIndex("Doctype"))) = "", Null, val(.TextMatrix(i, .ColIndex("Doctype"))))

                 
               '  RsDevsub("CusID").value = IIf((.TextMatrix(i, .ColIndex("CusID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CusID"))))
                RsDevsub.update
        
             
      End If
     Next i
    End With
   ' mTableName3 = mTableName2 & "2"
    RsDevsub.Close
    StrSQL = "SELECT  *  from " & mTableName3 & "  Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim xx As Long
    With grdDet
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("Item_ID"))) <> 0 Then
        RsDevsub.AddNew
                 RsDevsub("ReCostID").value = val(Me.TxtSerial1(mIndex).Text)
                 RsDevsub("IDRef").value = IIf((.TextMatrix(i, .ColIndex("IDRef"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDRef"))))
                 RsDevsub("Transaction_ID").value = IIf((.TextMatrix(i, .ColIndex("Transaction_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
                 RsDevsub("NoteSerial1").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial1"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial1")))
                 RsDevsub("Transaction_Date").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Null, .TextMatrix(i, .ColIndex("Transaction_Date")))
                 RsDevsub("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", Null, val(.TextMatrix(i, .ColIndex("StoreID"))))
                 RsDevsub("StoreID2").value = IIf((.TextMatrix(i, .ColIndex("StoreId2"))) = "", Null, val(.TextMatrix(i, .ColIndex("StoreId2"))))
                 RsDevsub("Item_ID").value = IIf((.TextMatrix(i, .ColIndex("Item_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Item_ID"))))
                 RsDevsub("UnitId").value = IIf((.TextMatrix(i, .ColIndex("UnitId"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitId"))))
                 RsDevsub("CusID").value = IIf((.TextMatrix(i, .ColIndex("CusID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CusID"))))
                 
                 RsDevsub("ShowQty").value = IIf((.TextMatrix(i, .ColIndex("ShowQty"))) = "", Null, val(.TextMatrix(i, .ColIndex("ShowQty"))))
                 RsDevsub("Transaction_Type").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Type"))) = "", Null, val(.TextMatrix(i, .ColIndex("Transaction_Type"))))
                 RsDevsub("OldshowPrice").value = IIf((.TextMatrix(i, .ColIndex("OldshowPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("OldshowPrice"))))

                 RsDevsub("FixesAssetsID").value = IIf((.TextMatrix(i, .ColIndex("FixesAssetsID"))) = "", Null, val(.TextMatrix(i, .ColIndex("FixesAssetsID"))))
                 RsDevsub("Emp_ID").value = IIf((.TextMatrix(i, .ColIndex("Emp_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Emp_ID"))))
                 RsDevsub("DepartementID").value = IIf((.TextMatrix(i, .ColIndex("DepartementID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DepartementID"))))

                 If RsDevsub("Transaction_Type").value = 21 Then
                    RsDevsub("OldCostPrice").value = IIf((.TextMatrix(i, .ColIndex("CostPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("CostPrice"))))
                     If SystemOptions.TypicalProduction = False Then
                        RsDevsub("CostPrice").value = IIf((.TextMatrix(i, .ColIndex("CostPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("CostPrice"))))
                        'RsDevsub("CostPrice").value = ModItemCostPrice.GetCostItemPrice(.TextMatrix(i, .ColIndex("Item_ID")), 0, , , SystemOptions.SysMainStockCostMethod, , , RsDevsub("Transaction_Date").value, val(.TextMatrix(i, .ColIndex("Transaction_ID"))), RsDevsub("UnitID").value, val(RsDevsub("StoreID").value))
'                        If RsDevsub("CostPrice").value = 0 Then
'                            RsDevsub("CostPrice").value = ModItemCostPrice.GetCostItemPrice(.TextMatrix(i, .ColIndex("Item_ID")), 0, , , LastPurPriceType, , , RsDevsub("Transaction_Date").value, val(.TextMatrix(i, .ColIndex("Transaction_ID"))), val(RsDevsub("UnitID").value), val(RsDevsub("StoreID").value))
'                        End If
                    Else
                        RsDevsub("CostPrice").value = 0
                    
                    End If
                    .TextMatrix(i, .ColIndex("Diff")) = val(RsDevsub("OldCostPrice").value) - val(RsDevsub("CostPrice").value)
                 Else
                   ' .TextMatrix(i, .ColIndex("showPrice")) = ModItemCostPrice.GetCostItemPrice(val(.TextMatrix(i, .ColIndex("Item_ID"))), 0, "", , SystemOptions.SysMainStockCostMethod, , , IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Date, .TextMatrix(i, .ColIndex("Transaction_Date"))), , val(.TextMatrix(i, .ColIndex("UnitId"))), val(.TextMatrix(i, .ColIndex("StoreID"))))
                    RsDevsub("showPrice").value = IIf((.TextMatrix(i, .ColIndex("showPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("showPrice"))))
                    RsDevsub("OldshowPrice").value = IIf((.TextMatrix(i, .ColIndex("OldshowPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("OldshowPrice"))))
                    RsDevsub("Price").value = val(.TextMatrix(i, .ColIndex("showPrice")))
                    RsDevsub("ItemCostPrice").value = val(.TextMatrix(i, .ColIndex("showPrice")))
                    '.TextMatrix(i, .ColIndex("Diff")) = Round(val(val(RsDevsub("OldshowPrice").value & "") - val(RsDevsub("showPrice").value & "")), 4)
                    RsDevsub("Diff").value = val(.TextMatrix(i, .ColIndex("Diff")))
                    RsDevsub("Valu").value = val(.TextMatrix(i, .ColIndex("ShowQty"))) * val(.TextMatrix(i, .ColIndex("showPrice")))
                 '   Dim RsUnitData As ADODB.Recordset
'                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & val(.TextMatrix(i, .ColIndex("Item_ID")))
'                    StrSQL = StrSQL + " AND UnitID=" & val(.TextMatrix(i, .ColIndex("UnitId")))
'                    Set RsUnitData = New ADODB.Recordset
'                    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
'                       RsDevsub("Quantity").value = RsUnitData("UnitFactor").value * val(.TextMatrix(i, .ColIndex("ShowQty")))
'                        RsDevsub("Price").value = val(IIf((.TextMatrix(i, .ColIndex("Price")) = ""), 0, val(.TextMatrix(i, .ColIndex("Price"))))) / RsUnitData("UnitFactor").value
'                    Else
'                        RsDevsub("Price").value = val(IIf((.TextMatrix(i, .ColIndex("Price")) = ""), 0, val(.TextMatrix(i, .ColIndex("Price")))))
'                    End If

                
                End If
                RsDevsub("Diff").value = val(.TextMatrix(i, .ColIndex("Diff")))
                 
             
             RsDevsub.update
             
            s = "Update dbo.Transaction_Details set FlgReCost=1,ReCostID=" & val(TxtSerial1(mIndex).Text) & " , "
            If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) <> 21 Then
                s = s & " showPrice = " & val(.TextMatrix(i, .ColIndex("showPrice")))
                s = s & " ,OldshowPrice = " & val(.TextMatrix(i, .ColIndex("OldshowPrice")))
               ' s = s & " ,ItemCostPrice = " & val(.TextMatrix(i, .ColIndex("showPrice")))
            Else
                s = s & " CostPrice = " & val(RsDevsub("CostPrice").value & "")
            End If
            s = s & " where ID=" & val(.TextMatrix(i, .ColIndex("IDRef"))) & ""
            Cn.Execute s
            
         
      End If
        CountNo.Caption = "”ÿ— —ﬁ„ " & i
        Status.Caption = "Õ›Ÿ  «·Õ—ﬂ… —ﬁ„ " & RsDevsub("NoteSerial1") & ""
     ' DoEvents
     Next i
    End With
    
    
'    s = " update Transaction_Details"
'    '   --set CostPrice=  isnull( dbo.GetItemCostPrice('01/01/2000',  dbo.Transactions.Transaction_Date ,  dbo.Transaction_Details.Item_ID)  ,0)
's = s & " SET    CostPrice = ROUND("
's = s & "            CONVERT("
's = s & "                FLOAT,"
's = s & "                ("
's = s & "                    SELECT SUM(T.Total) / SUM(T.Quantity) Quantity"
's = s & "                    FROM   ("
's = s & "                               SELECT 'Total' = CASE"
's = s & "                                                     WHEN TT2.ItemDiscountType = 1"
's = s & "                               OR TT2.ItemDiscountType = 0 THEN TT2.Quantity * TT2.Price WHEN TT2.ItemDiscountType = 2"
's = s & "                                  THEN ((TT2.Quantity * TT2.Price) -TT2.ItemDiscount) WHEN TT2.ItemDiscountType = 3 THEN (TT2.Quantity * TT2.Price)"
's = s & "                                  * (1 -(TT2.ItemDiscount / 100)) ELSE 0 END,"
's = s & "                               TT2.Quantity FROM dbo.Transaction_Details TT2 INNER JOIN dbo.Transactions TT1 ON"
's = s & "                               TT2.Transaction_ID = TT1.Transaction_ID"
'
's = s & "                               WHERE ("
's = s & "                                   TT1.Transaction_Type = 28"
's = s & "                                   OR TT1.Transaction_Type = 3 "
's = s & "                                   OR TT1.Transaction_Type = 20"
's = s & "                                   OR TT1.Transaction_Type = 34"
's = s & "                                   OR TT1.Transaction_Type = 0"
's = s & "                                   OR TT1.Transaction_Type = 15"
's = s & "                               )"
's = s & "                               AND TT1.Transaction_Date >= '1-1-1900'"
's = s & "                               AND TT1.Transaction_Date <= dbo.Transactions.Transaction_Date"
's = s & "                               AND TT2.Item_ID = dbo.Transaction_Details.Item_ID"
's = s & "                               AND TT1.Transaction_ID <> Transactions.Transaction_ID"
's = s & "                           ) T"
's = s & "                ),"
's = s & " 3"
's = s & "            ),"
's = s & " 3"
's = s & "        )"
'
'
's = s & " From dbo.Transactions"
's = s & "        INNER JOIN dbo.Transaction_Details"
's = s & "             ON  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
's = s & " WHERE  ("
's = s & "            dbo.Transactions.Transaction_Type = 21"
's = s & "            OR dbo.Transactions.Transaction_Type = 9"
's = s & "        )"
'
'If Me.TxtModFlg(mIndex).Text = "N" Then
''sql = sql & "  And (dbo.Transaction_Details.ReCostID Is Null)"
'End If
'If Not IsNull(FrmDate(mIndex).value) Then
'    s = s & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
'End If
'If Not IsNull(ToDate(mIndex).value) Then
'    s = s & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
'End If
'If DCboStoreName.Text <> "" Then
'    s = s & "  And dbo.transactions.StoreID =" & val(DCboStoreName.BoundText)
'End If
'
'If dcBranch(1).Text <> "" Then
'    s = s & "  And dbo.transactions.BranchId =" & val(dcBranch(1).BoundText)
'End If
'
'
'
'
'
'Cn.Execute s
    
    Status.Caption = " „ Õ›Ÿ «·Õ—ﬂ« "
    Status.Caption = ""
    CountNo.Caption = "⁄œœ «·«”ÿ—" & i
'     If BolFrmLoaded = True Then
'            cProgress.StopProgess
'            Set cProgress = Nothing
'        End If
End Sub
Private Sub CreateNote(ByRef NoteID As Long, ByVal mDate As Date, ByVal mTransType As Integer, ByVal mValue As Double, NoteSerial As Double, ByVal mBranchNo As Integer, ByVal NoteSerial1 As String)
     Dim StrSqlDel As String
     Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
    Set RsNotesGeneral = New ADODB.Recordset
  '  RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
  
        '
        StrSqlDel = "delete From Notes where noteid=" & val(NoteID)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        general_noteid = val(NoteID)
    

    

    'If SngTemp = 0 Then TxtNoteSerial.Text = "":   GoTo novalue
    RsNotesGeneral.AddNew
    
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    
    general_noteid = RsNotesGeneral("NoteID").value
    RsNotesGeneral.update
    NoteID = general_noteid
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    RsNotesGeneral("NoteDate").value = mDate
    If mTransType = 19 Then
        RsNotesGeneral("NoteType").value = 180  ' «–‰ «÷«›… ' 180
    Else
        RsNotesGeneral("NoteType").value = 190  ' «–‰ «÷«›… ' 180
    End If
    
    
    RsNotesGeneral("Note_Value").value = Abs(mValue)
    RsNotesGeneral("Note_ValueSales").value = Abs(mValue)
    
    RsNotesGeneral("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
    RsNotesGeneral("NoteSerial1").value = NoteSerial1
    NoteSerial = RsNotesGeneral("NoteSerial").value
  '  RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
    
       RsNotesGeneral("Remark").value = NoteSerial
     '
      
    
 RsNotesGeneral("OldNoteSerial1").value = NoteSerial
 '  Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    RsNotesGeneral("sanad_year").value = year(mDate)
    RsNotesGeneral("sanad_month").value = Month(mDate)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral("branch_no").value = val(mBranchNo)
    RsNotesGeneral.update
 
novalue:

End Sub
Private Sub CreateVouc(ByVal mRow As Long, Optional ByVal mType As Integer = 0)
 Dim LngDevID As Long
 Dim Line1 As Double
Dim Line2 As Double
Dim OtherInformation As New ClsGLOther
 Dim usedaccount As Integer
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim mStoreId As Integer
    Dim mStoreId2 As Integer
    Dim mItemCode As String
    Dim mTransID As Long
    Dim mDoctype As Integer, mOldTotal As Double
    Dim mCustId As Integer, NoteID As Long, mNoteSerial As Double, mDate As Date, mBranchNo As Integer
    Dim mNoteSerial1 As String
    Dim mFixesAssetsID As Integer, mEmp_ID As Double, mDepartementID As Double, project_id As Integer, mTransType As Integer
    Dim mValue As Double, SngTemp As Double
    Dim mDebitCredit As Integer
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    Dim DebitAccount As String, general_noteid As Long
Dim CreditAccount As String
    If mType = 0 Then
        With grdMaster
            mStoreId = val(.TextMatrix(mRow, .ColIndex("StoreID")))
            mStoreId2 = val(.TextMatrix(mRow, .ColIndex("StoreID2")))
            mStoreId2 = IIf(mStoreId2 = 0, mStoreId, mStoreId2)
            'mItemCode = (.TextMatrix(mRow, .ColIndex("Item_ID")))
            mValue = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
            mOldTotal = val(.TextMatrix(mRow, .ColIndex("OldTotal")))
            mTransID = val(.TextMatrix(mRow, .ColIndex("Transaction_ID")))
    '        mNoteSerial = val(.TextMatrix(mRow, .ColIndex("NoteSerial")))
            NoteID = val(.TextMatrix(mRow, .ColIndex("NoteID")))
            mDoctype = val(.TextMatrix(mRow, .ColIndex("Doctype")))
            mCustId = val(.TextMatrix(mRow, .ColIndex("CusID")))
            mNoteSerial1 = Trim(.TextMatrix(mRow, .ColIndex("NoteSerial1")))
            mFixesAssetsID = val(.TextMatrix(mRow, .ColIndex("FixesAssetsID")))
            mDate = (.TextMatrix(mRow, .ColIndex("Transaction_Date")))
            mEmp_ID = val(.TextMatrix(mRow, .ColIndex("Emp_ID")))
            mDepartementID = val(.TextMatrix(mRow, .ColIndex("DepartementID")))
            mTransType = val(.TextMatrix(mRow, .ColIndex("Transaction_Type")))
            mBranchNo = val(.TextMatrix(mRow, .ColIndex("BranchId")))
           ' project_id = val(.TextMatrix(mRow, .ColIndex("project_id")))
        End With
    Else
            With grdMaster2
                mStoreId = val(.TextMatrix(mRow, .ColIndex("StoreID")))
                mStoreId2 = val(.TextMatrix(mRow, .ColIndex("StoreID")))
                
                'mItemCode = (.TextMatrix(mRow, .ColIndex("Item_ID")))
                mValue = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
                mOldTotal = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
                mTransID = val(.TextMatrix(mRow, .ColIndex("Transaction_ID2")))
                TxtSerial1(1) = val(.TextMatrix(mRow, .ColIndex("Transaction_ID")))
        '        mNoteSerial = val(.TextMatrix(mRow, .ColIndex("NoteSerial")))
                NoteID = 0 'val(.TextMatrix(mRow, .ColIndex("NoteID")))
                mDoctype = val(.TextMatrix(mRow, .ColIndex("Doctype")))
                mCustId = val(.TextMatrix(mRow, .ColIndex("CusID")))
                mNoteSerial1 = Trim(.TextMatrix(mRow, .ColIndex("NoteSerial12")))
                mFixesAssetsID = val(.TextMatrix(mRow, .ColIndex("FixesAssetsID")))
                mDate = (.TextMatrix(mRow, .ColIndex("Transaction_Date")))
                mEmp_ID = val(.TextMatrix(mRow, .ColIndex("Emp_ID")))
                mDepartementID = val(.TextMatrix(mRow, .ColIndex("DepartementID")))
                mTransType = 19
                mBranchNo = val(.TextMatrix(mRow, .ColIndex("BranchId")))
            End With
                
                
                

    End If
    If mTransType <> 19 Then
    mTransType = mTransType
    End If
    If mTransID = 49694 Then
        mTransType = mTransType
    End If
    '    s = "Select Top 1 DD.Double_Entry_Vouchers_ID LngDevID,DEV_ID_Line_No,Notes_ID "
    's = s & " from DOUBLE_ENTREY_VOUCHERS DD WHERE Transaction_ID =" & mTransID & " Order By DEV_ID_Line_No DESC"
    'rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ' If mTransID <> 10 Then
        s = "Delete Notes Where NoteId In (Select Transactions.NoteId From Transactions Where Transactions.NoteId =Notes.NoteID  and  Transactions.Transaction_ID = " & mTransID & ")"
          Cn.Execute s
         s = "Delete DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID In (Select Transactions.NoteId From Transactions Where Transactions.NoteId =Notes_ID  and  Transactions.Transaction_ID = " & mTransID & ")"
         Cn.Execute s
    'End If
'    If Not rsDummy.EOF Then
'        LngDevID = val(rsDummy!LngDevID & "")
'        LngDevNO = val(rsDummy!DEV_ID_Line_No & "")
'        general_noteid = val(rsDummy!notes_id & "")
'    Else
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        LngDevNO = 1
        
       ' Exit Sub
      ' mValue = mOldTotal
       CreateNote NoteID, mDate, mTransType, mValue, mNoteSerial, mBranchNo, mNoteSerial1
       s = "Update transactions Set  NoteID = " & NoteID & " , NoteSerial = " & mNoteSerial & " Where Transaction_ID = " & mTransID
       general_noteid = NoteID
       Cn.Execute s
       
   ' End If
    If LngDevID = 0 Then Exit Sub
 '   s = "Select DD.Double_Entry_Vouchers_ID  from DOUBLE_ENTREY_VOUCHERS DD WHERE Transaction_ID =  " &
   ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = mValue

    If SngTemp <> 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
        OtherInformation.NextAccount_Code = get_store_Account(mStoreId, "Account_Code")
        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
                End If
            End If

            Dim UseCustomerAcc As Integer

            If val(mDoctype) > 0 Then
                getDocAccounts val(mDoctype), StrTempAccountCode, , , , , usedaccount, , , , , UseCustomerAcc
        
                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                
                ElseIf usedaccount = 0 And UseCustomerAcc = 0 Then
                
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                ElseIf usedaccount = 0 And UseCustomerAcc = 1 Then
                 
                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(mCustId))
                       
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
            End If

            DebitAccount = StrTempAccountCode
    
            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & mNoteSerial1 & "  " & "„‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
            Else
                StrTempDes = "Issue Voucher No.  " & mNoteSerial1 & "  "
            End If

            Line1 = setfoxy_Line
            LngDevNO = LngDevNO + 1
           'EditW
            mDebitCredit = IIf(SngTemp < 0, 1, 0)
            
            mDebitCredit = IIf(SngTemp < 0, 0, 1)
            mDebitCredit = 1
            
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, val(mTransID), , val(TxtSerial1(1)), , , , , , Line1, , , , , , mFixesAssetsID, , , val(mBranchNo), , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
            '«·„Œ“Ê‰ ›Ì «·›—⁄
            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If
        
            If val(mDoctype) > 0 Then
                getDocAccounts val(mDoctype), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
            End If

            CreditAccount = StrTempAccountCode
    
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & mNoteSerial1 & "  „‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
            Else
                StrTempDes = "Issue Voucher No. " & mNoteSerial1
            End If
    
            LngDevNO = LngDevNO + 1
            Line2 = setfoxy_Line
           'EditW
            mDebitCredit = IIf(SngTemp < 0, 1, 0)
            mDebitCredit = IIf(SngTemp < 0, 0, 1)
            mDebitCredit = 0
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(mIndex)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            If mTransType = 19 Then
                Account_Code_dynamic = get_account_code_branch(1, my_branch)
            Else
                Account_Code_dynamic = get_store_Account(mStoreId2, "Account_Code")
            End If
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If
'mDoctype = 0
            If val(mDoctype) > 0 Then
                getDocAccounts val(mDoctype), StrTempAccountCode, , , , , usedaccount, , , , , UseCustomerAcc

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 And UseCustomerAcc = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                ElseIf usedaccount = 0 And UseCustomerAcc = 1 Then
                 
                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(mCustId))
                
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
            End If
            DebitAccount = StrTempAccountCode
            
            Line1 = setfoxy_Line

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & mNoteSerial1 & "  " & "„‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
            Else
                StrTempDes = "Issue Voucher No.  " & mNoteSerial1 & "  "
            End If
    
            LngDevNO = LngDevNO + 1
        'EditW
            mDebitCredit = IIf(SngTemp < 0, 1, 0)
            mDebitCredit = IIf(SngTemp < 0, 0, 1)
            'project_id
            mDebitCredit = 0
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line1, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
           ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

            If val(mDoctype) > 0 Then
                getDocAccounts val(mDoctype), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    Account_Code_dynamic = StrTempAccountCode
                ElseIf usedaccount = 0 Then
        
                    Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")
                End If

            Else
                
             ' If mTransType = 10 Then
             '   Account_Code_dynamic = get_store_Account(mStoreId2, "Account_Code")
                
            'Else
               Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")
            'End If
                

            End If
        
            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            CreditAccount = StrTempAccountCode
            OtherInformation.NextAccount_Code = DebitAccount
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & mNoteSerial1 & " „‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›… "
            Else
                StrTempDes = "Issue Voucher No. " & mNoteSerial1
            End If

            Line2 = setfoxy_Line
         
            LngDevNO = LngDevNO + 1

            'EditW
            mDebitCredit = IIf(SngTemp < 0, 0, 1)
            mDebitCredit = IIf(SngTemp < 0, 1, 0)
            'project_id
            mDebitCredit = 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If


        End If

        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If
ErrTrap:
End Sub




Private Sub CreateVoucNew(ByVal mRow As Long, Optional ByVal mType As Integer = 0)
 Dim LngDevID As Long
 Dim Line1 As Double
Dim Line2 As Double
Dim OtherInformation As New ClsGLOther
 Dim usedaccount As Integer
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim mStoreId As Integer
    Dim mStoreId2 As Integer
    Dim mItemCode As String
    Dim mTransID As Long
    Dim mDoctype As Integer, mOldTotal As Double
    Dim mCustId As Integer, NoteID As Long, mNoteSerial As Double, mDate As Date, mBranchNo As Integer
    Dim mNoteSerial1 As String
    Dim mFixesAssetsID As Integer, mEmp_ID As Double, mDepartementID As Double, project_id As Integer, mTransType As Integer
    Dim mValue As Double, SngTemp As Double
    Dim mDebitCredit As Integer
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    Dim DebitAccount As String, general_noteid As Long
Dim CreditAccount As String
    If mType = 0 Then
        With grdMaster
            mStoreId = val(.TextMatrix(mRow, .ColIndex("StoreID")))
            mStoreId2 = val(.TextMatrix(mRow, .ColIndex("StoreID2")))
            mStoreId2 = IIf(mStoreId2 = 0, mStoreId, mStoreId2)
            'mItemCode = (.TextMatrix(mRow, .ColIndex("Item_ID")))
            mValue = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
            mValue = val(.TextMatrix(mRow, .ColIndex("OldTotal")))
            mTransID = val(.TextMatrix(mRow, .ColIndex("Transaction_ID")))
    '        mNoteSerial = val(.TextMatrix(mRow, .ColIndex("NoteSerial")))
            NoteID = val(.TextMatrix(mRow, .ColIndex("NoteID")))
            mDoctype = val(.TextMatrix(mRow, .ColIndex("Doctype")))
            mCustId = val(.TextMatrix(mRow, .ColIndex("CusID")))
            mNoteSerial1 = Trim(.TextMatrix(mRow, .ColIndex("NoteSerial1")))
            mFixesAssetsID = val(.TextMatrix(mRow, .ColIndex("FixesAssetsID")))
            mDate = (.TextMatrix(mRow, .ColIndex("Transaction_Date")))
            mEmp_ID = val(.TextMatrix(mRow, .ColIndex("Emp_ID")))
            mDepartementID = val(.TextMatrix(mRow, .ColIndex("DepartementID")))
            mTransType = val(.TextMatrix(mRow, .ColIndex("Transaction_Type")))
            mBranchNo = val(.TextMatrix(mRow, .ColIndex("BranchId")))
           ' project_id = val(.TextMatrix(mRow, .ColIndex("project_id")))
        End With
    Else
            With grdMaster2
                mStoreId = val(.TextMatrix(mRow, .ColIndex("StoreID")))
                mStoreId2 = val(.TextMatrix(mRow, .ColIndex("StoreID")))
                
                'mItemCode = (.TextMatrix(mRow, .ColIndex("Item_ID")))
                mValue = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
                mOldTotal = val(.TextMatrix(mRow, .ColIndex("DiffTotal")))
                mTransID = val(.TextMatrix(mRow, .ColIndex("Transaction_ID2")))
                TxtSerial1(1) = val(.TextMatrix(mRow, .ColIndex("Transaction_ID")))
        '        mNoteSerial = val(.TextMatrix(mRow, .ColIndex("NoteSerial")))
                NoteID = 0 'val(.TextMatrix(mRow, .ColIndex("NoteID")))
                mDoctype = val(.TextMatrix(mRow, .ColIndex("Doctype")))
                mCustId = val(.TextMatrix(mRow, .ColIndex("CusID")))
                mNoteSerial1 = Trim(.TextMatrix(mRow, .ColIndex("NoteSerial12")))
                mFixesAssetsID = val(.TextMatrix(mRow, .ColIndex("FixesAssetsID")))
                mDate = (.TextMatrix(mRow, .ColIndex("Transaction_Date")))
                mEmp_ID = val(.TextMatrix(mRow, .ColIndex("Emp_ID")))
                mDepartementID = val(.TextMatrix(mRow, .ColIndex("DepartementID")))
                mTransType = 19
                mBranchNo = val(.TextMatrix(mRow, .ColIndex("BranchId")))
            End With
                
                
                

    End If
    If mTransType <> 19 Then
    mTransType = mTransType
    End If
    If mTransID = 49694 Then
        mTransType = mTransType
    End If

        Dim RsTemp  As ADODB.Recordset
     Set RsTemp = New ADODB.Recordset
    StrSQL = "select * From Transactions where ReturnID = " & mTransID
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        mStoreId2 = IIf(IsNull(RsTemp("StoreID").value), "", RsTemp("StoreID").value)
    End If
    '    s = "Select Top 1 DD.Double_Entry_Vouchers_ID LngDevID,DEV_ID_Line_No,Notes_ID "
    's = s & " from DOUBLE_ENTREY_VOUCHERS DD WHERE Transaction_ID =" & mTransID & " Order By DEV_ID_Line_No DESC"
    'rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If mTransID <> 10 Then
          s = "Delete Notes Where NoteId In (Select Transactions.NoteId From Transactions Where Transactions.NoteId =Notes.NoteID  and  Transactions.Transaction_ID = " & mTransID & ")"
          Cn.Execute s
         s = "Delete DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID In (Select Transactions.NoteId From Transactions Where Transactions.NoteId =Notes_ID  and  Transactions.Transaction_ID = " & mTransID & ")"
         Cn.Execute s
    'End If
'    If Not rsDummy.EOF Then
'        LngDevID = val(rsDummy!LngDevID & "")
'        LngDevNO = val(rsDummy!DEV_ID_Line_No & "")
'        general_noteid = val(rsDummy!notes_id & "")
'    Else
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        LngDevNO = 1
        
       ' Exit Sub
      ' mValue = mOldTotal

        
       
       CreateNote NoteID, mDate, mTransType, mValue, mNoteSerial, mBranchNo, mNoteSerial1
       s = "Update transactions Set  NoteID = " & NoteID & " , NoteSerial = " & mNoteSerial & " Where Transaction_ID = " & mTransID
       general_noteid = NoteID
       Cn.Execute s
       
   ' End If
    If LngDevID = 0 Then Exit Sub
 '   s = "Select DD.Double_Entry_Vouchers_ID  from DOUBLE_ENTREY_VOUCHERS DD WHERE Transaction_ID =  " &
   ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = mValue

    If SngTemp <> 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
        
        
        If mTransType = 19 Then
            OtherInformation.NextAccount_Code = get_store_Account(mStoreId, "Account_Code")
            If detect_inventory_work_type = 1 Then
                Account_Code_dynamic = get_account_code_branch(1, my_branch)
            
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else
    
                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
                    End If
                End If
    
                Dim UseCustomerAcc As Integer
    
                If val(mDoctype) > 0 Then
                    getDocAccounts val(mDoctype), StrTempAccountCode, , , , , usedaccount, , , , , UseCustomerAcc
            
                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··”‰œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    
                    ElseIf usedaccount = 0 And UseCustomerAcc = 0 Then
                    
                        StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                    ElseIf usedaccount = 0 And UseCustomerAcc = 1 Then
                     
                        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(mCustId))
                           
                    End If
    
                Else
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                End If
    
                DebitAccount = StrTempAccountCode
        
                'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "  √–‰ ’—›  —ﬁ„     " & mNoteSerial1 & "  " & "„‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
                Else
                    StrTempDes = "Issue Voucher No.  " & mNoteSerial1 & "  "
                End If
    
                Line1 = setfoxy_Line
                LngDevNO = LngDevNO + 1
               'EditW
                mDebitCredit = IIf(SngTemp < 0, 1, 0)
                
                mDebitCredit = IIf(SngTemp < 0, 0, 1)
                mDebitCredit = 1
                
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, val(mTransID), , val(TxtSerial1(1)), , , , , , Line1, , , , , , mFixesAssetsID, , , val(mBranchNo), , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
        
                '«·„Œ“Ê‰ ›Ì «·›—⁄
                Account_Code_dynamic = get_account_code_branch(0, my_branch)
            
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else
    
                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
             
                    End If
                End If
            
                If val(mDoctype) > 0 Then
                    getDocAccounts val(mDoctype), , StrTempAccountCode, , , , , usedaccount
    
                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
            
                    ElseIf usedaccount = 0 Then
            
                        StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
                    End If
    
                Else
                    StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
                End If
    
                CreditAccount = StrTempAccountCode
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰ ’—›  —ﬁ„ " & mNoteSerial1 & "  „‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
                Else
                    StrTempDes = "Issue Voucher No. " & mNoteSerial1
                End If
        
                LngDevNO = LngDevNO + 1
                Line2 = setfoxy_Line
               'EditW
                mDebitCredit = IIf(SngTemp < 0, 1, 0)
                mDebitCredit = IIf(SngTemp < 0, 0, 1)
                mDebitCredit = 0
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(mIndex)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
        
            ElseIf detect_inventory_work_type = 2 Then
                If mTransType = 19 Then
                    Account_Code_dynamic = get_account_code_branch(1, my_branch)
                Else
                    Account_Code_dynamic = get_store_Account(mStoreId2, "Account_Code")
                End If
            
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else
    
                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
             
                    End If
                End If
    'mDoctype = 0
                If val(mDoctype) > 0 Then
                    getDocAccounts val(mDoctype), StrTempAccountCode, , , , , usedaccount, , , , , UseCustomerAcc
    
                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··”‰œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
            
                    ElseIf usedaccount = 0 And UseCustomerAcc = 0 Then
            
                        StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                    ElseIf usedaccount = 0 And UseCustomerAcc = 1 Then
                     
                        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(mCustId))
                    
                    End If
    
                Else
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                End If
                DebitAccount = StrTempAccountCode
                
                Line1 = setfoxy_Line
    
                'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "  √–‰ ’—›  —ﬁ„     " & mNoteSerial1 & "  " & "„‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›…"
                Else
                    StrTempDes = "Issue Voucher No.  " & mNoteSerial1 & "  "
                End If
        
                LngDevNO = LngDevNO + 1
            'EditW
                mDebitCredit = IIf(SngTemp < 0, 1, 0)
                mDebitCredit = IIf(SngTemp < 0, 0, 1)
                'project_id
                mDebitCredit = 0
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line1, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
    
                '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
               ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    
                If val(mDoctype) > 0 Then
                    getDocAccounts val(mDoctype), , StrTempAccountCode, , , , , usedaccount
    
                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                        Account_Code_dynamic = StrTempAccountCode
                    ElseIf usedaccount = 0 Then
            
                        Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")
                    End If
    
                Else
                    
                 ' If mTransType = 10 Then
                 '   Account_Code_dynamic = get_store_Account(mStoreId2, "Account_Code")
                    
                'Else
                   Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")
                'End If
                    
    
                End If
            
                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                    GoTo ErrTrap
                End If
        
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
                CreditAccount = StrTempAccountCode
                OtherInformation.NextAccount_Code = DebitAccount
                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰ ’—›  —ﬁ„ " & mNoteSerial1 & " „‰ Õ—ﬂ… ÷»ÿ «· ﬂ·›… "
                Else
                    StrTempDes = "Issue Voucher No. " & mNoteSerial1
                End If
    
                Line2 = setfoxy_Line
             
                LngDevNO = LngDevNO + 1
    
                'EditW
                mDebitCredit = IIf(SngTemp < 0, 0, 1)
                mDebitCredit = IIf(SngTemp < 0, 1, 0)
                'project_id
                mDebitCredit = 1
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(SngTemp), mDebitCredit, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                    GoTo ErrTrap
                End If
    
    
            End If
        Else
                    If detect_inventory_work_type = 1 Then
                    ' 1«·„Œ“Ê‰ ›Ì «·›—⁄
                        Account_Code_dynamic = get_account_code_branch(0, my_branch)
            
                        If Account_Code_dynamic = "NO branch" Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            Cmd(2).Enabled = True
                            GoTo ErrTrap
                        Else
        
                            If Account_Code_dynamic = "NO account" Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                Cmd(2).Enabled = True
                                GoTo ErrTrap
                 
                            End If
                        End If
    
                    StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
        
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "√–‰  ÕÊÌ· »÷«∆⁄ »Ì‰ «·„Œ«“‰  —ﬁ„ " & mNoteSerial1
                    Else
                        StrTempDes = "  Moving Items Vchr  No. " & mNoteSerial1
                    End If
            
                    LngDevNO = 0
    
    
                   mDebitCredit = IIf(SngTemp < 0, 0, 1)
                mDebitCredit = IIf(SngTemp < 0, 1, 0)
                'project_id
                mDebitCredit = 1
               
 
    
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
         
                    LngDevNO = 1
    
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
        
                ElseIf detect_inventory_work_type = 2 Then
                    Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")
    
                    If Account_Code_dynamic = "" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄      " & mStoreId, vbCritical
                        GoTo ErrTrap
                    End If
        
                    StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
    
                    ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & mNoteSerial1
                    Else
                        StrTempDes = " Moving Items Vchr  No. " & mNoteSerial1
                    End If
        
                    LngDevNO = 1
    
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
    
                    '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
        
                    Account_Code_dynamic = get_store_Account(mStoreId2, "Account_Code")
    
                    If Account_Code_dynamic = "" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    " & mStoreId2, vbCritical
                        GoTo ErrTrap
                    End If
        
                    StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
    
                    ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & mNoteSerial1
                    Else
                        StrTempDes = " Moving Items Vchr  No. " & mNoteSerial1
                    End If
        
                    LngDevNO = 0
    
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    Dim BranchId1  As Integer
                    Dim BranchID2  As Integer
                    Dim DeptSide1 As String
                    Dim CreditSide1 As String
                    Dim noteid1 As Double


                    BranchId1 = GetInventoryBranch(mStoreId)
                    
                    BranchID2 = GetInventoryBranch(mStoreId2)
        LngDevNO = 1
    If BranchId1 <> BranchID2 Then
    
     DeptSide1 = getBranchCurrentAccount(BranchId1)
    CreditSide1 = getBranchCurrentAccount(BranchID2)
    LngDevNO = LngDevNO + 1
    If CreditSide1 <> "" Then
         If ModAccounts.AddNewDev(LngDevID, LngDevNO, CreditSide1, SngTemp, 0, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
    End If
                    LngDevNO = LngDevNO + 1
                    If DeptSide1 <> "" Then
         If ModAccounts.AddNewDev(LngDevID, LngDevNO, DeptSide1, SngTemp, 1, StrTempDes, general_noteid, , , , mDate, user_id, mTransID, , val(TxtSerial1(1)), , , , , , Line2, , , , , , mFixesAssetsID, , , mBranchNo, , , , , , , mDepartementID, mEmp_ID, , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    End If
                    noteid1 = val(general_noteid)
                    updateNotesValueAndNobytext noteid1, CDbl(SngTemp)
                    
                    
    End If

        End If
        End If
        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If
ErrTrap:
End Sub

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch(mIndex).BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    TxtRemarks(mIndex).Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)



    If mIndex = 1 Then

        FrmDate(mIndex).value = IIf(IsNull(RsSavRec.Fields("FrmDate").value), Date, RsSavRec.Fields("FrmDate").value)
        ToDate(mIndex).value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
        DCboStoreName.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
        
        DcboItemID1.BoundText = IIf(IsNull(RsSavRec.Fields("ItemID").value), "", RsSavRec.Fields("ItemID").value)
        
        Command5(mIndex).Enabled = IIf(IsNull(RsSavRec.Fields("EntryCreated").value), True, Not RsSavRec.Fields("EntryCreated").value)
        Command5(mIndex).Enabled = True
    End If
  '  TxtNoteID(mIndex).Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
  '  TxtNoteSerial(mIndex).Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
   ' FrmDate(mIndex).value = IIf(IsNull(RsSavRec.Fields("FrmDate").value), Date, RsSavRec.Fields("FrmDate").value)
   ' ToDate(mIndex).value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
     LabCurrRec(mIndex).Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec(mIndex).Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
If mIndex = 0 Then
    FullGridData
ElseIf mIndex = 1 Then
    FullGridData2
End If

ErrTrap:
End Sub
Function GetValue() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(Valu) AS TotalValue"
sql = sql & " From dbo.TblReCostDet"
sql = sql & " Where (ReCostID = " & TxtSerial1(mIndex).Text & ")"
sql = sql & " GROUP BY StoreID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetValue = IIf(IsNull(rs2("TotalValue").value), 0, rs2("TotalValue").value)
Else
GetValue = 0
End If
End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    ≈⁄«œ… ≈Õ ”«» «· ﬂ·›… " & TxtSerial1(mIndex).Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblReCost"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1(mIndex))
Notevalue = 0
notytype = 9085
Notevalue = GetValue()
BranchID = val(Dcbranch(mIndex).BoundText)
NoteDate = (XPDtbTrans(mIndex).value)

If Notevalue > 0 Then

                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (TxtNoteSerial(mIndex)), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID(mIndex).Text = NoteID
                                                     TxtNoteSerial(mIndex).Text = NoteSerial

CREATE_VOUCHER_GE val(TxtNoteID(mIndex).Text), BranchID, user_id, NoteDate
'rs.Resync adAffectCurrent


     End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
    Dim i As Long
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    ≈⁄«œ… ≈Õ ”«» «· ﬂ·›…" & TxtSerial1(1).Text
    notes_id = general_noteid
    my_branch = val(Dcbranch(1).BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    sql = " SELECT     SUM(Valu) AS TotalValue, StoreID"
    sql = sql & "    From dbo.TblReCostDet"
    sql = sql & "  Where (ReCostID = " & val(TxtSerial1(mIndex).Text) & ")"
    sql = sql & "  GROUP BY StoreID"
   rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
    line_no = 1
    rs2.MoveFirst
   For i = 1 To rs2.RecordCount
   Notevalue = IIf(IsNull(rs2("TotalValue").value), 0, rs2("TotalValue").value)
   StoreID6 = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            If Notevalue > 0 Then

                             StrAccountCodeDebt = get_account_code_branch(1, my_branch)
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  ﬂ·›… «·„»Ì⁄«   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            StrAccountCodeCridet = get_store_Account(StoreID6, "Account_Code")
                            line_no = line_no + 1
                                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  «·„Œ“Ê‰  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                             line_no = line_no + 1
            End If
         rs2.MoveNext
      Next i
      End If

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click(Index As Integer)
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
'      If Dcbranch(mIndex).Text = "" And val(Dcbranch(mIndex).BoundText) = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
'            Else
'            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'         End If
''            Dcbranch(mIndex).Enabled = True
''            Dcbranch(mIndex).SetFocus
'            Exit Sub
'     End If
    ' -------------------------------------- txtmodflg type -------------------
    
    Select Case Me.TxtModFlg(mIndex).Text
            '------------------------------ new record ----------------------------
        Case "N"
                '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec

        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id(mTableName, "ID", "")
    Me.TxtSerial1(mIndex).Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 1
   sql = "  SELECT     dbo.TblReCostDet.ReCostID, dbo.TblReCostDet.IDRef, dbo.TblReCostDet.Transaction_ID, dbo.TblReCostDet.NoteSerial1, dbo.TblReCostDet.Transaction_Date,"
   sql = sql & "                    dbo.TblReCostDet.ShowQty, dbo.TblReCostDet.showPrice, dbo.TblReCostDet.Price, dbo.TblReCostDet.Valu, dbo.TblReCostDet.Quantity, dbo.TblReCostDet.StoreID,"
   sql = sql & "                   dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblReCostDet.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
   sql = sql & "                   dbo.TblReCostDet.UnitID , dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
   sql = sql & "     FROM         dbo.TblReCostDet LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblUnites ON dbo.TblReCostDet.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblItems ON dbo.TblReCostDet.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblStore ON dbo.TblReCostDet.StoreID = dbo.TblStore.StoreID"
   sql = sql & "   Where (dbo.TblReCostDet.ReCostID = " & val(TxtSerial1(1).Text) & ")"

Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Long
     With FG
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDRef")) = IIf(IsNull(Rs1("IDRef").value), "", Rs1("IDRef").value)
                   .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs1("Transaction_ID").value), "", Rs1("Transaction_ID").value)
                   .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                   .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rs1("Transaction_Date").value), "", Rs1("Transaction_Date").value)
                   .TextMatrix(i, .ColIndex("ShowQty")) = IIf(IsNull(Rs1("ShowQty").value), "", Rs1("ShowQty").value)
                   .TextMatrix(i, .ColIndex("showPrice")) = IIf(IsNull(Rs1("showPrice").value), "", Rs1("showPrice").value)
                   .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs1("Price").value), "", Rs1("Price").value)
                   .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), "", Rs1("Quantity").value)
                   .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), "", Rs1("StoreID").value)
                   .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(Rs1("Item_ID").value), "", Rs1("Item_ID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(Rs1("UnitId").value), "", Rs1("UnitId").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                Else
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
           End If
                   Rs1.MoveNext
             Next i
        End With

        Exit Sub
ErrTrap:
    End Sub
    
    
 Sub FullGridData2()
 
          

           
           

 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    grdMaster.Clear flexClearScrollable, flexClearEverything
    grdMaster.Rows = 1
   sql = "  SELECT     dbo.TblReCostCalcDet.ReCostID,dbo.TblReCostCalcDet.DiffTotal,dbo.TblReCostCalcDet.OldTotal, dbo.TblReCostCalcDet.IDRef, dbo.TblReCostCalcDet.Transaction_ID, dbo.TblReCostCalcDet.NoteSerial1, dbo.TblReCostCalcDet.Transaction_Date,"
   sql = sql & "                   TblReCostCalcDet.Transaction_Type,TblReCostCalcDet.CusID,tc.CusName,TblReCostCalcDet.Note_Value,"
   sql = sql & "                    dbo.TblReCostCalcDet.StoreID, TblReCostCalcDet.BranchId,TblReCostCalcDet.NoteID,TblReCostCalcDet.NoteSerial,"
   sql = sql & "                   dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,TS2.StoreName StoreName2,TS2.StoreID StoreID2 ,"
    sql = sql & "  TblReCostCalcDet.CusID, TblReCostCalcDet.FixesAssetsID, TblReCostCalcDet.DepartementID, TblReCostCalcDet.Emp_ID,TblReCostCalcDet.Doctype"
  
   sql = sql & "     FROM         dbo.TblReCostCalcDet "
   sql = sql & "         LEFT OUTER JOIN dbo.TblCustemers AS tc"
   sql = sql & "               ON  dbo.TblReCostCalcDet.CusID = tc.CusID"
   sql = sql & "    LEFT OUTER JOIN                dbo.TblStore ON dbo.TblReCostCalcDet.StoreID = dbo.TblStore.StoreID"
      sql = sql & "    LEFT OUTER JOIN                dbo.TblStore TS2 ON dbo.TblReCostCalcDet.StoreID2 =TS2.StoreID"
   sql = sql & "   Where (dbo.TblReCostCalcDet.ReCostID = " & val(TxtSerial1(mIndex).Text) & ")"
   sql = sql & "   Order By TblReCostCalcDet.Transaction_ID,TblReCostCalcDet.NoteSerial1"
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Long
     With grdMaster
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDRef")) = IIf(IsNull(Rs1("IDRef").value), "", Rs1("IDRef").value)
                   .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs1("Transaction_ID").value), "", Rs1("Transaction_ID").value)
                   .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                   .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rs1("Transaction_Date").value), "", Rs1("Transaction_Date").value)
                   .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(Rs1("Note_Value").value), "", Rs1("Note_Value").value)
                   .TextMatrix(i, .ColIndex("DiffTotal")) = IIf(IsNull(Rs1("DiffTotal").value), "", Rs1("DiffTotal").value)
                   .TextMatrix(i, .ColIndex("OldTotal")) = IIf(IsNull(Rs1("OldTotal").value), "", Rs1("OldTotal").value)
                    .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs1("NoteSerial").value), "", Rs1("NoteSerial").value)
                    .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs1("NoteID").value), "", Rs1("NoteID").value)
                    .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs1("BranchId").value), "", Rs1("BranchId").value)

                   .TextMatrix(i, .ColIndex("FixesAssetsID")) = IIf(IsNull(Rs1("FixesAssetsID").value), "", Rs1("FixesAssetsID").value)
                   .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs1("CusID").value), "", Rs1("CusID").value)
                   .TextMatrix(i, .ColIndex("DepartementID")) = IIf(IsNull(Rs1("DepartementID").value), "", Rs1("DepartementID").value)
                   .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs1("Emp_ID").value), "", Rs1("Emp_ID").value)
                    .TextMatrix(i, .ColIndex("Doctype")) = IIf(IsNull(Rs1("Doctype").value), "", Rs1("Doctype").value)
                    .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(Rs1("Transaction_Type").value), "", Rs1("Transaction_Type").value)
                   .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), "", Rs1("StoreID").value)
                    .TextMatrix(i, .ColIndex("StoreId2")) = IIf(IsNull(Rs1("StoreID2").value), "", Rs1("StoreID2").value)
                   .TextMatrix(i, .ColIndex("StoreName2")) = IIf(IsNull(Rs1("StoreName2").value), "", Rs1("StoreName2").value)
    
                    If Rs1("Transaction_Type").value = 19 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰ ’—›"
                    ElseIf Rs1("Transaction_Type").value = 10 Or Rs1("Transaction_Type").value = 992 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰"
                    ElseIf Rs1("Transaction_Type").value = 11 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«” ·«„ „‰ «·„Œ«“‰"
                        
                    Else
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "„»Ì⁄« "
                    End If
                   
                  ' .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(Rs1("Item_ID").value), "", Rs1("Item_ID").value)
                 '  .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                  ' .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(Rs1("UnitId").value), "", Rs1("UnitId").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                 '  .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                  ' .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                Else
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                  ' .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                  ' .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
           End If
                   Rs1.MoveNext
             Next i
        End With


   grdDet.Clear flexClearScrollable, flexClearEverything
    grdDet.Rows = 1
   sql = "  SELECT     dbo.TblReCostCalcDet2.ReCostID, dbo.TblReCostCalcDet2.IDRef, dbo.TblReCostCalcDet2.Transaction_ID, dbo.TblReCostCalcDet2.NoteSerial1, dbo.TblReCostCalcDet2.Transaction_Date,"
   sql = sql & "                    dbo.TblReCostCalcDet2.ShowQty, dbo.TblReCostCalcDet2.showPrice, dbo.TblReCostCalcDet2.Price, dbo.TblReCostCalcDet2.Valu, dbo.TblReCostCalcDet2.Quantity, dbo.TblReCostCalcDet2.StoreID,"
   sql = sql & "                   dbo.TblStore.StoreName,TblReCostCalcDet2.Diff, dbo.TblStore.StoreNamee, dbo.TblReCostCalcDet2.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
   sql = sql & "                   dbo.TblReCostCalcDet2.UnitID ,TblReCostCalcDet2.Transaction_Type,TblReCostCalcDet2.OldshowPrice, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
   sql = sql & " , TblReCostCalcDet2.CusID, TblReCostCalcDet2.FixesAssetsID, TblReCostCalcDet2.DepartementID, TblReCostCalcDet2.Emp_ID,TblReCostCalcDet2.Doctype,TblReCostCalcDet2.Doctype"
    sql = sql & "                   ,TS2.StoreName StoreName2,TS2.StoreID StoreID2 "
   sql = sql & "     FROM         dbo.TblReCostCalcDet2 LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblUnites ON dbo.TblReCostCalcDet2.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblItems ON dbo.TblReCostCalcDet2.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblStore ON dbo.TblReCostCalcDet2.StoreID = dbo.TblStore.StoreID"
   sql = sql & "    LEFT OUTER JOIN                dbo.TblStore TS2 ON dbo.TblReCostCalcDet2.StoreID2 =TS2.StoreID"
   sql = sql & "   Where (dbo.TblReCostCalcDet2.ReCostID = " & val(TxtSerial1(mIndex).Text) & ")"
sql = sql & "   Order By TblReCostCalcDet2.Transaction_ID,TblReCostCalcDet2.NoteSerial1"
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     'Dim i As Integer
     With grdDet
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDRef")) = IIf(IsNull(Rs1("IDRef").value), "", Rs1("IDRef").value)
                   .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs1("Transaction_ID").value), "", Rs1("Transaction_ID").value)
                   .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                   .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rs1("Transaction_Date").value), "", Rs1("Transaction_Date").value)
                   .TextMatrix(i, .ColIndex("ShowQty")) = IIf(IsNull(Rs1("ShowQty").value), "", Rs1("ShowQty").value)
                   .TextMatrix(i, .ColIndex("showPrice")) = Round(IIf(IsNull(Rs1("showPrice").value), "", Rs1("showPrice").value), 3)
                   .TextMatrix(i, .ColIndex("Price")) = Round(IIf(IsNull(Rs1("Price").value), "", Rs1("Price").value), 3)
                   .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(i, .ColIndex("Diff")) = IIf(IsNull(Rs1("Diff").value), "", Rs1("Diff").value)
                   .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), "", Rs1("Quantity").value)
                   .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs1("StoreID").value), "", Rs1("StoreID").value)
                   .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(Rs1("Item_ID").value), "", Rs1("Item_ID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(Rs1("UnitId").value), "", Rs1("UnitId").value)
                    .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(Rs1("Transaction_Type").value), "", Rs1("Transaction_Type").value)
                   .TextMatrix(i, .ColIndex("FixesAssetsID")) = IIf(IsNull(Rs1("FixesAssetsID").value), "", Rs1("FixesAssetsID").value)
                   .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs1("CusID").value), "", Rs1("CusID").value)
                   .TextMatrix(i, .ColIndex("DepartementID")) = IIf(IsNull(Rs1("DepartementID").value), "", Rs1("DepartementID").value)
                   .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs1("Emp_ID").value), "", Rs1("Emp_ID").value)
                    .TextMatrix(i, .ColIndex("Doctype")) = IIf(IsNull(Rs1("Doctype").value), "", Rs1("Doctype").value)
                    
                   .TextMatrix(i, .ColIndex("StoreId2")) = IIf(IsNull(Rs1("StoreID2").value), "", Rs1("StoreID2").value)
                   .TextMatrix(i, .ColIndex("StoreName2")) = IIf(IsNull(Rs1("StoreName2").value), "", Rs1("StoreName2").value)
                    .TextMatrix(i, .ColIndex(("OldshowPrice"))) = Round(IIf(IsNull(Rs1("OldshowPrice").value), "", Rs1("OldshowPrice").value), 3)
                    If Rs1("Transaction_Type").value = 19 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰ ’—›"
                    ElseIf Rs1("Transaction_Type").value = 10 Or Rs1("Transaction_Type").value = 992 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰"
                    ElseIf Rs1("Transaction_Type").value = 11 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«” ·«„ „‰ «·„Œ«“‰"
                        
                    Else
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "„»Ì⁄« "
                    End If
            If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreName").value), "", Rs1("StoreName").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                Else
                   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs1("StoreNamee").value), "", Rs1("StoreNamee").value)
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
           End If
                   Rs1.MoveNext
             Next i
        End With

        Exit Sub
ErrTrap:
    End Sub
    Private Sub RemoveGridRow2()
    With Me.FG
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub

Private Sub ISButton5_Click(Index As Integer)
print_report
End Sub

Sub FillGrid()
Dim sql As String
Dim i As Long
  FG.Clear flexClearScrollable, flexClearEverything
  FG.Rows = 1
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.Transaction_Details.Quantity, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1, dbo.Transactions.StoreID,"
sql = sql & "                       dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty,"
sql = sql & "                       dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
sql = sql & "                       dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.Transaction_Details.ReCostID,"
sql = sql & "                       dbo.Transaction_Details.FlgReCost"
sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                       dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
sql = sql & "  Where (dbo.transactions.Transaction_Type = 19) And (dbo.Transaction_Details.ShowPrice = 0) "
If Me.TxtModFlg(mIndex).Text = "N" Then
sql = sql & "  And (dbo.Transaction_Details.ReCostID Is Null)"
End If
If Me.TxtModFlg(mIndex).Text = "E" Then
sql = sql & "  And ((dbo.Transaction_Details.FlgReCost Is Null) or (dbo.Transaction_Details.ReCostID =" & val(TxtSerial1(1).Text) & " ))"
End If
If Not IsNull(FrmDate(mIndex).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(mIndex).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(mIndex).value, True) & ""
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With FG
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex(("Ser"))) = i
.TextMatrix(i, .ColIndex(("IDRef"))) = IIf(IsNull(rs2("ID").value), "", rs2("ID").value)
.TextMatrix(i, .ColIndex(("Transaction_ID"))) = IIf(IsNull(rs2("Transaction_ID").value), "", rs2("Transaction_ID").value)
.TextMatrix(i, .ColIndex(("NoteSerial1"))) = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
.TextMatrix(i, .ColIndex(("Transaction_Date"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)
.TextMatrix(i, .ColIndex(("StoreID"))) = IIf(IsNull(rs2("StoreID").value), "", rs2("StoreID").value)
.TextMatrix(i, .ColIndex(("Fullcode"))) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex(("Item_ID"))) = IIf(IsNull(rs2("Item_ID").value), "", rs2("Item_ID").value)
.TextMatrix(i, .ColIndex(("UnitId"))) = IIf(IsNull(rs2("UnitId").value), "", rs2("UnitId").value)
.TextMatrix(i, .ColIndex(("ShowQty"))) = IIf(IsNull(rs2("ShowQty").value), "", rs2("ShowQty").value)
.TextMatrix(i, .ColIndex(("showPrice"))) = IIf(IsNull(rs2("showPrice").value), "", rs2("showPrice").value)
.TextMatrix(i, .ColIndex(("Quantity"))) = IIf(IsNull(rs2("Quantity").value), "", rs2("Quantity").value)
.TextMatrix(i, .ColIndex(("Price"))) = IIf(IsNull(rs2("Price").value), "", rs2("Price").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
.TextMatrix(i, .ColIndex(("ItemName"))) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
.TextMatrix(i, .ColIndex(("UnitName"))) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
Else
.TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
.TextMatrix(i, .ColIndex(("ItemName"))) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
.TextMatrix(i, .ColIndex(("UnitName"))) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
End If
rs2.MoveNext
Next i
End If
End With
End Sub



Sub FillGrid3()
Dim sql As String
Dim i As Long, j As Long
Dim mTime As Date
Dim mTime2 As Date
mTime = Time
Dim MinDate As Date
Dim s As String

If SystemOptions.CostStarting = True Then
     Dim FirstPeriodDateInthisYear  As Date
     getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                               
    MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
Else
    MinDate = "1-1-1900"
End If

  grdMaster2.Clear flexClearScrollable, flexClearEverything
  grdMaster2.Rows = 1
  
Dim rs2 As ADODB.Recordset
Dim RsDet  As New ADODB.Recordset
Set rs2 = New ADODB.Recordset



s = " SELECT * FROM ("
s = s & " SELECT Nots,"
s = s & "        TN.NoteSerial1,"
s = s & "        TN.Transaction_ID,"
s = s & "        Transaction_Date"
s = s & " FROM   ("
s = s & "            SELECT SUM(td.Quantity) Qty,"
s = s & "                   t.Nots,"
s = s & "                   td.Item_ID,"
s = s & "                   td.UnitId,"
s = s & "                   NoteSerial1,"
s = s & "                   t.Transaction_ID,"
s = s & "                   t.Transaction_Date"
s = s & "            FROM   Transaction_Details td"
s = s & "                   INNER JOIN Transactions AS t"
s = s & "                        ON  t.Transaction_ID = td.Transaction_ID"
s = s & "            Where t.Transaction_Type = 21"
s = s & "            Group By"
s = s & "                   td.Transaction_ID,"
s = s & "                   td.Item_ID,"
s = s & "                   td.UnitId,"
s = s & "                   Nots,"
s = s & "                   NoteSerial1,"
s = s & "                   t.Transaction_ID,"
s = s & "                   t.Transaction_Date"
s = s & "        ) TN"
s = s & " WHERE  TN.Item_ID IN (SELECT Item_ID"
s = s & "                      FROM   ("
s = s & "                                 SELECT *"
s = s & "                                 FROM   ("
s = s & "                                            SELECT SUM(td.Quantity) Qty,"
s = s & "                                                   td.Transaction_ID,"
s = s & "                                                   td.Item_ID,"
s = s & "                                                   td.UnitId,"
s = s & "                                                   NoteSerial1"
s = s & "                                            FROM   Transaction_Details td"
s = s & "                                                   INNER JOIN Transactions AS t"
s = s & "                                                        ON  t.Transaction_ID = td.Transaction_ID"
s = s & "                                            Where t.Transaction_Type = 19"
s = s & "                                            Group By"
s = s & "                                                   td.Transaction_ID,"
s = s & "                                                   td.Item_ID,"
s = s & "                                                   td.UnitId,"
s = s & "                                                   NoteSerial1"
s = s & "                                        ) N"
s = s & "                                 Where N.Transaction_ID = TN.nots"
s = s & "                                         AND ROUND(TN.Qty,2) <> ROUND(N.Qty,2)"
s = s & "                                        AND N.Item_ID = TN.Item_ID"
s = s & "                                        AND N.UnitId = TN.UnitId"
s = s & "                             ) H)"
s = s & " Group By"
s = s & "        Nots,"
s = s & "        TN.NoteSerial1,"
s = s & "        TN.Transaction_ID,"
s = s & "        TN.Transaction_Date"

s = s & " Union all"
s = s & " SELECT Nots,"
s = s & "        NoteSerial1,"
s = s & "        Transactions.Transaction_ID,"
s = s & "        transactions.Transaction_Date"
s = s & " From Transaction_Details"
s = s & "        INNER JOIN Transactions"
s = s & "             ON  Transactions.Transaction_ID = Transaction_Details.Transaction_ID"
s = s & "        INNER JOIN TblItems AS ti"
s = s & "             ON  ti.ItemID = Transaction_Details.Item_ID"
s = s & " Where Transaction_Type = 21"
s = s & "        AND ItemType = 0"
s = s & "        AND Item_ID NOT IN (SELECT td.Item_ID"
s = s & "                            FROM   Transaction_Details td"
s = s & "                                   INNER JOIN Transactions AS t"
s = s & "                                        ON  t.Transaction_ID = td.Transaction_ID"
s = s & "                            Where t.Transaction_Type = 19"
s = s & "                                   AND t.Transaction_ID = Transactions.Nots)"
s = s & " Group By"
s = s & "        Nots,"
s = s & "        NoteSerial1,"
s = s & "        Transactions.Transaction_ID,"
s = s & "        transactions.Transaction_Date"
s = s & "        "
s = s & "        "
s = s & "   Union all"
s = s & " SELECT Nots,"
s = s & "        NoteSerial1,"
s = s & "        Transaction_ID,"
s = s & "        Transaction_Date"
s = s & " FROM   ("
s = s & "            SELECT Transactions.Transaction_ID,"
s = s & "                   Nots,"
s = s & "                   Transaction_Date,"
s = s & "                   Transactions.NoteSerial1,"
s = s & "                   VV = ("
s = s & "                       SELECT SUM(ISNULL(td.CostPrice, 0) * ISNULL(td.ShowQty, 0))"
s = s & "                       FROM   Transaction_Details AS td"
s = s & "                       Where td.Transaction_ID = transactions.Transaction_ID"
s = s & "                   ),"
s = s & "                   VV2 = ("
s = s & "                       SELECT SUM(ISNULL(ShowQty, 0) * ISNULL(showPrice, 0))"
s = s & "                       FROM   Transaction_Details DD"
s = s & "                              INNER JOIN Transactions B"
s = s & "                                   ON  B.Transaction_ID = DD.Transaction_ID"
s = s & "                       WHERE  B.nots = CAST(Transactions.Transaction_ID AS NVARCHAR(50))"
s = s & "                              AND B.Transaction_Type = 19"
s = s & "                   )"
s = s & "            From transactions"
s = s & "            WHERE  ("
s = s & "                       transactions.Transaction_Type = 21"
s = s & "                       OR Transactions.Transaction_Type = 9"
s = s & "                   )"
                  
's = s & "                   --   AND (Transactions.Transaction_Date >= '01-01-2018')
's = s & "                   -- AND (Transactions.Transaction_Date <= '01-17-2018')"
's = s & "                   --  AND Transactions.Transaction_ID = 1076"
       
s = s & "        ) NN"
s = s & " Where Abs(vv - vv2) > 2"
s = s & " Group By"
s = s & " Transaction_ID,"
s = s & "        Nots,"
       
s = s & "        NoteSerial1,"
s = s & "        Transaction_Date"
s = s & "   Union all"

s = s & " SELECT t.Nots,"
s = s & "        t.NoteSerial1,"
s = s & "        t.Transaction_ID,"
s = s & "        t.Transaction_Date"
s = s & " FROM   Transactions         AS t"
s = s & "        INNER JOIN Transaction_Details"
s = s & "             ON  Transaction_Details.Transaction_ID = t.Transaction_ID"
s = s & "        INNER JOIN TblItems  AS ti"
s = s & "             ON  ti.ItemID = Transaction_Details.Item_ID"
s = s & " Where Transaction_Type = 21"
s = s & "        AND ItemType = 0"
s = s & "        AND Nots NOT IN (SELECT t2.Transaction_ID"
s = s & "                         FROM   Transactions AS t2"
s = s & "                         WHERE  t2.Transaction_Type = 19)"
s = s & " Group By"
s = s & "        t.Transaction_ID,"
s = s & "        t.Nots,"
s = s & "        t.NoteSerial1,"
s = s & "        t.Transaction_Date"


s = s & "   Union all "
 



s = s & "   SELECT Nots as Transaction_ID ,"
s = s & "          "
s = s & "          NoteSerial1,Transaction_ID as Nots,"
s = s & "          Transaction_Date"
s = s & "   From Transactions"
s = s & "   Where 1 = 1 "
If Not IsNull(FrmDate(mIndex).value) Then
    s = s & "  And Transaction_Date >=" & SQLDate(FrmDate(mIndex).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
    s = s & "  And Transaction_Date <=" & SQLDate(ToDate(mIndex).value, True) & ""
End If

s = s & "   and Transaction_Type = 21 AND"
s = s & "   Nots IN"
s = s & "   (SELECT Transaction_ID"
s = s & "   From Transactions"
s = s & "   Where Transaction_Type = 19"
s = s & "          AND NoteId NOT IN (SELECT NoteId"
s = s & "                             FROM   Notes))"

s = s & "   Union all"

s = s & "   SELECT Nots as Transaction_ID ,"
s = s & "          "
s = s & "          NoteSerial1,Transaction_ID as Nots,"
s = s & "          Transaction_Date"

s = s & "   From Transactions"
s = s & "   Where"
s = s & "   Transaction_Type = 21 AND"
s = s & "   Nots IN (SELECT Transaction_ID"
s = s & "                   FROM   ("
s = s & "                              SELECT NoteID,"
s = s & "                                     NoteSerial,"
s = s & "                                     t.Transaction_ID,"
s = s & "                                     NoteSerial1,"
s = s & "                                     t.Transaction_Date,"
s = s & "                                     VV = ISNULL("
s = s & "                                         ("
s = s & "                                             SELECT SUM(ISNULL(td.CostPrice, 0) * ISNULL(td.ShowQty, 0))"
s = s & "                                             FROM   Transaction_Details AS td"
s = s & "                                                    INNER JOIN Transactions AS tt"
s = s & "                                                         ON  tt.Transaction_ID = td.Transaction_ID"
s = s & "                                             WHERE  CAST(tt.Transaction_ID AS NVARCHAR(50)) = t.nots"
s = s & "                                                    AND Transaction_Type = 21"
s = s & "                                         ),"
s = s & "   0"
s = s & "                                     ),"
s = s & "                                     EntryOut = ISNULL("
s = s & "                                         ("
s = s & "                                             SELECT SUM(dev.[Value])"
s = s & "                                             FROM   DOUBLE_ENTREY_VOUCHERS AS dev"
s = s & "                                             WHERE  dev.Account_Code = ("
s = s & "                                                        SELECT ts.Account_Code"
s = s & "                                                        FROM   TblStore AS ts"
s = s & "                                                        Where ts.StoreId = t.StoreId"
s = s & "                                                    )"
s = s & "                                                    AND dev.Notes_ID = t.NoteId"
s = s & "                                         ),"
s = s & "   0"
s = s & "                                     ),"
s = s & "                                     ISNULL(SUM((ISNULL(ShowQty, 0) * ISNULL(showPrice, 0))), 0) VALUE"
s = s & "                              FROM   Transactions AS t"
s = s & "                                     INNER JOIN Transaction_Details AS td"
s = s & "                                          ON  td.Transaction_ID = t.Transaction_ID"
s = s & "                              Where t.Transaction_Type = 19"

If Not IsNull(FrmDate(mIndex).value) Then
    s = s & "  And t.Transaction_Date >=" & SQLDate(FrmDate(mIndex).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
    s = s & "  And t.Transaction_Date <=" & SQLDate(ToDate(mIndex).value, True) & ""
End If

s = s & "                              Group By"
s = s & "                                     t.NoteId,"
s = s & "                                     NoteID,"
s = s & "                                     t.nots,"
s = s & "                                     t.NoteSerial1,"
s = s & "                                     t.Transaction_ID,"
s = s & "                                     t.Transaction_Date,"
s = s & "                                     t.StoreID,"
s = s & "                                     NoteSerial"
s = s & "                          ) T"
s = s & "                   Where "
s = s & "                        "
s = s & "                          (ABS(ISNULL(T.EntryOut, 0) - VV) > 1))"




s = s & "        Union all"

s = s & " SELECT Nots            AS Transaction_ID,"
s = s & "        dsd.NoteSerial1,"
s = s & "        Transaction_ID  AS Nots,"

s = s & "        dsd.Transaction_Date"
s = s & " FROM   Transactions       dsd"
s = s & " WHERE  nots IN (SELECT t.nots"
s = s & "                 FROM   Transactions AS t"
s = s & "                 Where t.Transaction_Type = 19"
s = s & "                        AND t.nots IN (SELECT d.Transaction_ID"
s = s & "                                       FROM   Transactions d"
s = s & "                                       WHERE  d.Transaction_Type = 21)"
s = s & "                 Group By"
s = s & "                        nots"
s = s & "                 HAVING (COUNT(*) > 1))"
                
s = s & " ) J"
                         
s = s & " Where 1 = 1                         "
If Not IsNull(FrmDate(mIndex).value) Then
    s = s & "  And Transaction_Date >=" & SQLDate(FrmDate(mIndex).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
    s = s & "  And Transaction_Date <=" & SQLDate(ToDate(mIndex).value, True) & ""
End If


s = s & "  GROUP BY    Transaction_ID,"
s = s & "         Nots,"
s = s & "         NoteSerial1,"
s = s & "         Transaction_Date"


'



rs2.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
' Set cProgress = New ClsProgress
'        BolFrmLoaded = True
'        cProgress.ProgressType = Waiting
'        cProgress.StartProgress
With grdMaster2
    If rs2.RecordCount > 0 Then
        .Rows = .Rows + rs2.RecordCount
        rs2.MoveFirst
        
        For i = 1 To .Rows - 1
            If .Rows <= i Then Exit Sub
            If Not rs2.EOF Then
                .TextMatrix(i, .ColIndex(("Ser"))) = i
                '.TextMatrix(I, .ColIndex(("IDRef"))) = IIf(IsNull(Rs2("ID").value), "", Rs2("ID").value)
                .TextMatrix(i, .ColIndex(("Transaction_ID"))) = IIf(IsNull(rs2("Transaction_ID").value), "", rs2("Transaction_ID").value)
                .TextMatrix(i, .ColIndex(("Notes"))) = IIf(IsNull(rs2("Nots").value), "", rs2("Nots").value)
                '.TextMatrix(i, .ColIndex(("Transaction_Type"))) = IIf(IsNull(rs2("Transaction_Type").value), "", rs2("Transaction_Type").value)
                '.TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)
                
                .TextMatrix(i, .ColIndex(("NoteSerial1"))) = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
                '.TextMatrix(i, .ColIndex(("NoteID"))) = IIf(IsNull(rs2("NoteID").value), "", rs2("NoteID").value)
                '.TextMatrix(i, .ColIndex(("BranchId"))) = IIf(IsNull(rs2("BranchId").value), "", rs2("BranchId").value)
                
                
                .TextMatrix(i, .ColIndex(("Transaction_Date"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)
                
                 
                
            
            
            
            
            
                rs2.MoveNext
            End If
            DoEvents
        Next i
    End If
End With

MsgBox " „ «·«œ—«Ã"




'
'       If BolFrmLoaded = True Then
'            cProgress.StopProgess
'            Set cProgress = Nothing
'        End If

End Sub


Sub FillGrid2()
Dim sql As String
Dim i As Long, j As Long
Dim mTime As Date
Dim mTime2 As Date
mTime = Time
Dim MinDate As Date


If SystemOptions.CostStarting = True Then
     Dim FirstPeriodDateInthisYear  As Date
     getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                               
    MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
Else
    MinDate = "1-1-1900"
End If

  grdMaster.Clear flexClearScrollable, flexClearEverything
  grdMaster.Rows = 1
  grdDet.Clear flexClearScrollable, flexClearEverything
  grdDet.Rows = 1
Dim rs2 As ADODB.Recordset
Dim RsDet  As New ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = " SELECT     DISTINCT dbo.Transactions.Transaction_Date,Transactions.FixesAssetsID,Transactions.Emp_ID,Transactions.DepartementID, dbo.Transactions.CusID"
sql = sql & " ,dbo.Transactions.Doctype, dbo.Transactions.Transaction_ID,Transactions.NoteID, Transactions.NoteSerial,Transactions.BranchId,  dbo.Transactions.NoteSerial1, dbo.Transactions.StoreID,"
sql = sql & "                       dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.Transaction_Type"
',Notes.Note_Value,"
sql = sql & "           ,tc.CusName,tc.CusNamee"
sql = sql & " ,Note_Value = (Select Sum(IsNull(ShowQty,0) *  IsNull(showPrice,0) ) From Transaction_Details DD Where DD.Transaction_ID = dbo.Transactions.Transaction_ID )"
'sql = sql & " ,Note_Value2 = (Select Sum(IsNull(ShowQty,0) *  IsNull(OldshowPrice,0) ) From Transaction_Details DD Where DD.Transaction_ID = dbo.Transactions.Transaction_ID )"
sql = sql & " ,T2.StoreID StoreID2,ts.StoreName StoreName2 "

sql = sql & "  FROM            Transactions LEFT OUTER JOIN"
sql = sql & "                           TblStore ON Transactions.StoreID = TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                           Notes ON Notes.NoteID = Transactions.NoteId LEFT OUTER JOIN"
sql = sql & "                           TblCustemers AS tc ON Transactions.CusID = tc.CusID LEFT OUTER JOIN"
sql = sql & "                           Transactions AS T2 ON Transactions.Transaction_ID = T2.ReturnID LEFT OUTER JOIN                          TblStore AS ts ON ts.StoreID = T2.StoreID"

'If DcboItemID1.Text <> "" Then
    sql = sql & "               Left Outer join            Transaction_Details  ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "
'End If
'sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
'sql = sql & "                       dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
'sql = sql & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
'sql = sql & "                                 LEFT OUTER JOIN Notes ON Notes.NoteID = transactions.NoteId AND Notes.NoteType = 180"
'sql = sql & "                                 LEFT OUTER JOIN TblCustemers AS tc On Transactions.CusID =tc.CusID"
'sql = sql & "                       LEFT OUTER JOIN  Transactions T2 ON Transactions.Transaction_ID = T2.ReturnID"
'sql = sql & "                         LEFT Outer JOIN TblStore AS ts ON ts.StoreID = T2.StoreID"

If chkIsMov.value Then
    sql = sql & "  Where ( dbo.transactions.Transaction_Type = 10 )"
Else
    sql = sql & "  Where (dbo.transactions.Transaction_Type = 19 Or  dbo.transactions.Transaction_Type = 10 Or dbo.transactions.Transaction_Type = 992 Or  dbo.transactions.Transaction_Type = 11 )"
End If


'And (dbo.Transaction_Details.ShowPrice = 0) "
If Me.TxtModFlg(mIndex).Text = "N" Then
'sql = sql & "  And (dbo.Transaction_Details.ReCostID Is Null)"
End If
If Me.TxtModFlg(mIndex).Text = "E" Then
    sql = sql & "  And ((dbo.Transaction_Details.FlgReCost Is Null) or (dbo.Transaction_Details.ReCostID =" & val(TxtSerial1(mIndex).Text) & " ))"
End If
If Not IsNull(FrmDate(mIndex).value) Then
    sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
    sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
End If
If DCboStoreName.Text <> "" Then
    sql = sql & "  And dbo.transactions.StoreID =" & val(DCboStoreName.BoundText)
End If

If DcboItemID1.Text <> "" Then
    sql = sql & "  And dbo.Transaction_Details.Item_ID =" & val(DcboItemID1.BoundText)
End If


If Dcbranch(1).Text <> "" And val(Dcbranch(1).BoundText) <> 0 Then
    sql = sql & "  And dbo.transactions.BranchId =" & val(Dcbranch(1).BoundText)
End If


sql = sql & "   Order By transactions.Transaction_Date,transactions.Transaction_ID,transactions.NoteSerial1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
' Set cProgress = New ClsProgress
'        BolFrmLoaded = True
'        cProgress.ProgressType = Waiting
'        cProgress.StartProgress
With grdMaster
If rs2.RecordCount > 0 Then
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst

For i = 1 To .Rows - 1
If .Rows <= i Then Exit Sub
.TextMatrix(i, .ColIndex(("Ser"))) = i
'.TextMatrix(I, .ColIndex(("IDRef"))) = IIf(IsNull(Rs2("ID").value), "", Rs2("ID").value)
.TextMatrix(i, .ColIndex(("Transaction_ID"))) = IIf(IsNull(rs2("Transaction_ID").value), "", rs2("Transaction_ID").value)
.TextMatrix(i, .ColIndex(("Transaction_Type"))) = IIf(IsNull(rs2("Transaction_Type").value), "", rs2("Transaction_Type").value)
.TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)

.TextMatrix(i, .ColIndex(("NoteSerial"))) = IIf(IsNull(rs2("NoteSerial").value), "", rs2("NoteSerial").value)
.TextMatrix(i, .ColIndex(("NoteID"))) = IIf(IsNull(rs2("NoteID").value), "", rs2("NoteID").value)
.TextMatrix(i, .ColIndex(("BranchId"))) = IIf(IsNull(rs2("BranchId").value), "", rs2("BranchId").value)


    
If rs2("Transaction_Type").value = 19 Then
    .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«–‰ ’—›"
ElseIf rs2("Transaction_Type").value = 992 Or rs2("Transaction_Type").value = 10 Then
        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = " ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
ElseIf rs2("Transaction_Type").value = 992 Or rs2("Transaction_Type").value = 11 Then
                        .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "«” ·«„ „‰ „Œ“‰"
Else
    .TextMatrix(i, .ColIndex(("Transaction_TypeName"))) = "„»Ì⁄« "
    
    
End If

.TextMatrix(i, .ColIndex(("Doctype"))) = IIf(IsNull(rs2("docType").value), "", rs2("docType").value)
.TextMatrix(i, .ColIndex(("NoteSerial1"))) = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
.TextMatrix(i, .ColIndex(("Transaction_Date"))) = IIf(IsNull(rs2("Transaction_Date").value), "", rs2("Transaction_Date").value)
.TextMatrix(i, .ColIndex(("StoreID"))) = IIf(IsNull(rs2("StoreID").value), "", rs2("StoreID").value)
.TextMatrix(i, .ColIndex(("Note_Value"))) = IIf(IsNull(rs2("Note_Value").value), "", rs2("Note_Value").value)

.TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)
.TextMatrix(i, .ColIndex(("FixesAssetsID"))) = IIf(IsNull(rs2("FixesAssetsID").value), "", rs2("FixesAssetsID").value)
.TextMatrix(i, .ColIndex(("Emp_ID"))) = IIf(IsNull(rs2("Emp_ID").value), "", rs2("Emp_ID").value)
.TextMatrix(i, .ColIndex(("DepartementID"))) = IIf(IsNull(rs2("DepartementID").value), "", rs2("DepartementID").value)

.TextMatrix(i, .ColIndex(("CusID"))) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)
.TextMatrix(i, .ColIndex(("StoreID2"))) = IIf(IsNull(rs2("StoreID2").value), "", rs2("StoreID2").value)
.TextMatrix(i, .ColIndex(("StoreName2"))) = IIf(IsNull(rs2("StoreName2").value), "", rs2("StoreName2").value)
If SystemOptions.UserInterface = ArabicInterface Then
    .TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
    .TextMatrix(i, .ColIndex(("CusName"))) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
Else
    .TextMatrix(i, .ColIndex(("StoreName"))) = IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
    .TextMatrix(i, .ColIndex(("CusName"))) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
End If
'GetItemCostPrice
'QryItemsTransactionsTotals
'QryItemsTransactionsTotalsByStores
          



sql = " SELECT    Transaction_Details.FlgReCost, dbo.Transaction_Details.ReCostID,dbo.Transaction_Details.Quantity,dbo.Transactions.Doctype,dbo.Transactions.CusID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1, dbo.Transactions.StoreID,"
sql = sql & "                       dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty,Transactions.FixesAssetsID,Transactions.Emp_ID,Transactions.DepartementID, dbo.Transactions.CusID,"
sql = sql & "                       dbo.Transaction_Details.showPrice, dbo.Transaction_Details.OldshowPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"

sql = sql & "                       dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.Transaction_Details.ReCostID,"
sql = sql & "                       dbo.Transaction_Details.FlgReCost,dbo.Transaction_Details.CostPrice"
sql = sql & " ,T2.StoreID StoreID2,ts.StoreName StoreName2 ,"

  'StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "

If SystemOptions.AllowCostPerStore = False Then
    '     sql = sql + "AvCost = (Select " & " CONVERT(money, Total / TotalQty, 3) AS AvCost FROM dbo.QryItemsTransactionsTotals(28 , 3,20, '1-1-1900', '" & Format(CDate(ToDate(mIndex).value), "MM/DD/YYYY") & "',Transaction_Details.Item_ID,dbo.Transactions.Transaction_ID))"
          If chkSalim.value = vbUnchecked Then
          
          
        sql = sql + "AvCost =Round(CONVERT(FLOAT, (SELECT SUM(T.Total) / SUM(T.Quantity)     Quantity"
      sql = sql + " FROM   ("
        sql = sql + "           SELECT 'Total' = CASE"
        sql = sql + "                                   WHEN TT2.ItemDiscountType = 1"
        sql = sql + "           OR TT2.ItemDiscountType = 0 THEN TT2.Quantity *"
        sql = sql + "              TT2.Price"
        sql = sql + "              WHEN TT2.ItemDiscountType = 2 THEN ((TT2.Quantity * TT2.Price) -TT2.ItemDiscount)"
        sql = sql + "              WHEN TT2.ItemDiscountType = 3 THEN (TT2.Quantity * TT2.Price) * (1 -(TT2.ItemDiscount / 100))"
        sql = sql + "              ELSE 0"
        sql = sql + "              END,"
        sql = sql + "              TT2.Quantity"
        sql = sql + "           From"
        sql = sql + "           dbo.Transaction_Details TT2"
        sql = sql + "           INNER JOIN dbo.Transactions TT1"
        sql = sql + "           ON TT2.Transaction_ID = TT1.Transaction_ID"
        sql = sql + "           WHERE ("
        sql = sql + "               TT1.Transaction_Type = 28"
        sql = sql + "               OR TT1.Transaction_Type = 3"
        sql = sql + "               OR TT1.Transaction_Type = 20"
        sql = sql + "               OR TT1.Transaction_Type = 34"
        sql = sql + "               OR TT1.Transaction_Type = 0"
        sql = sql + "               OR TT1.Transaction_Type = 15"
        sql = sql + "           )"
        sql = sql + "           AND TT1.Transaction_Date >=  " & SQLDate(MinDate, True) & ""
        sql = sql + "           AND TT1.Transaction_Date <=  dbo.Transactions.Transaction_Date"
        sql = sql + "           AND TT2.Item_ID = dbo.Transaction_Details.Item_ID"
        sql = sql + "           AND TT1.Transaction_ID <> Transactions.Transaction_ID"
        sql = sql + "       )                                  T),3),3)"
       
          
        Else
sql = sql + "price as AvCost "

'        sql = sql & "AvCost= (  "
'        sql = sql & "SELECT          dbo.Transaction_Details.NewCost"
 
 
'sql = sql & " FROM         dbo.Transactions TT1  INNER JOIN"
'sql = sql & "                       dbo.Transaction_Details TT2  ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'sql = sql & "  WHERE     (dbo.TransactionTypes.StockEffect <>0) "
'sql = sql & "  AND (dbo.Transactions.Transaction_Date <= " & SQLDate(Transaction_Date, True) & ")"

'sql = sql & "  AND               (TT2.Item_ID =  Item_ID )"
'
' sql = sql & " AND (dbo.TT1.StoreID =  StoreID) "
'
 
'sql = sql & "  AND (TT1.Transaction_ID <>  dbo.transactions.Transaction_Date  )"
'sql = sql & "  )"
 


        End If
         
         
Else
       ' sql = sql + "AvCost = (Select " & " CONVERT(money, Total / TotalQty, 3) AS AvCost FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '1-1-1900', '" & Format(CDate(ToDate(mIndex).value), "MM/DD/YYYY") & "',dbo.Transactions.StoreID,Transaction_Details.Item_ID,dbo.Transactions.Transaction_ID))"
        
           If chkSalim.value = vbUnchecked Then
        sql = sql + "AvCost =Round(CONVERT(FLOAT, (SELECT SUM(T.Total) / SUM(T.Quantity)     Quantity"
        Else
        sql = sql + "newcost as AvCost "
        End If
        sql = sql + " FROM   ("
        sql = sql + "           SELECT 'Total' = CASE"
        sql = sql + "                                   WHEN TT2.ItemDiscountType = 1"
        sql = sql + "           OR TT2.ItemDiscountType = 0 THEN TT2.Quantity *"
        sql = sql + "              TT2.Price"
        sql = sql + "              WHEN TT2.ItemDiscountType = 2 THEN ((TT2.Quantity * TT2.Price) -TT2.ItemDiscount)"
        sql = sql + "              WHEN TT2.ItemDiscountType = 3 THEN (TT2.Quantity * TT2.Price) * (1 -(TT2.ItemDiscount / 100))"
        sql = sql + "              ELSE 0"
        sql = sql + "              END,"
        sql = sql + "              TT2.Quantity"
        sql = sql + "           From"
        sql = sql + "           dbo.Transaction_Details TT2"
        sql = sql + "           INNER JOIN dbo.Transactions TT1"
        sql = sql + "           ON TT2.Transaction_ID = TT1.Transaction_ID"
        sql = sql + "           WHERE ("
        sql = sql + "               TT1.Transaction_Type = 28"
        sql = sql + "               OR TT1.Transaction_Type = 3"
        sql = sql + "               OR TT1.Transaction_Type = 20"
        sql = sql + "               OR TT1.Transaction_Type = 34"
        sql = sql + "               OR TT1.Transaction_Type = 0"
        sql = sql + "               OR TT1.Transaction_Type = 15"
        sql = sql + "           )"
        sql = sql + "           AND TT1.Transaction_Date >= " & SQLDate(MinDate, True) & ""
        sql = sql + "           AND TT1.Transaction_Date <=  dbo.Transactions.Transaction_Date"
        sql = sql + "           AND TT2.Item_ID = dbo.Transaction_Details.Item_ID"
        sql = sql + "                     AND"
        sql = sql + "             TT1.storeid =dbo.Transactions.StoreID"
        sql = sql + "           AND TT1.Transaction_ID <> Transactions.Transaction_ID"
        sql = sql + "       )                                  T),3),3)"
       
         
End If

sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                       dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
sql = sql & "                       LEFT OUTER JOIN  Transactions T2 ON Transactions.Transaction_ID = T2.ReturnID"
sql = sql & "                         LEFT Outer JOIN TblStore AS ts ON ts.StoreID = T2.StoreID"
sql = sql & "  Where (dbo.transactions.Transaction_Type = 19 Or dbo.transactions.Transaction_Type = 10 Or dbo.transactions.Transaction_Type = 992  Or  dbo.transactions.Transaction_Type = 11 ) "
sql = sql & " And Transactions.Transaction_ID = " & val(grdMaster.TextMatrix(i, grdMaster.ColIndex(("Transaction_ID"))))
'And (dbo.Transaction_Details.ShowPrice = 0) "
If Me.TxtModFlg(mIndex).Text = "N" Then
'sql = sql & "  And (dbo.Transaction_Details.ReCostID Is Null)"
End If
If Me.TxtModFlg(mIndex).Text = "E" Then
sql = sql & "  And ((dbo.Transaction_Details.FlgReCost Is Null) or (dbo.Transaction_Details.ReCostID =" & val(TxtSerial1(mIndex).Text) & " ))"
End If
If Not IsNull(FrmDate(mIndex).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
End If
If Not IsNull(ToDate(mIndex).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
End If

If Dcbranch(1).Text <> "" Then
    sql = sql & "  And dbo.transactions.BranchId =" & val(Dcbranch(1).BoundText)
End If

If DcboItemID1.Text <> "" Then
    sql = sql & "  And dbo.Transaction_Details.Item_ID =" & val(DcboItemID1.BoundText)
End If

sql = sql & "   Order By  transactions.Transaction_Date,transactions.Transaction_ID,transactions.NoteSerial1"


If RsDet.State = 1 Then RsDet.Close
RsDet.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText



      '  Do While Rs2.State = adStateExecuting
      
      '  Loop
Dim mCost As Double
  Dim mRowNo As Long

    With grdDet
        If RsDet.RecordCount > 0 Then
            mRowNo = .Rows
            .Rows = .Rows + RsDet.RecordCount
            RsDet.MoveFirst
            
            For j = mRowNo To .Rows - 1
                mCost = val(RsDet!AvCost & "")
                If mCost = 0 And SystemOptions.CostStarting Then
                                                
                        getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                        Dim LastPurchaseDate As String
                        Dim LastPurchasePrice As Double
                        Dim LastPurchaseqty As Double
                        GetlastPurchasedata 20, val(RsDet!Item_ID & ""), "01/01/1900", Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                        
                         
                         mCost = LastPurchasePrice

                    
                End If
                
                .TextMatrix(j, .ColIndex(("Ser"))) = j
                .TextMatrix(j, .ColIndex(("IDRef"))) = IIf(IsNull(RsDet("ID").value), "", RsDet("ID").value)
                .TextMatrix(j, .ColIndex(("Transaction_ID"))) = IIf(IsNull(RsDet("Transaction_ID").value), "", RsDet("Transaction_ID").value)
                .TextMatrix(j, .ColIndex(("Transaction_Type"))) = IIf(IsNull(RsDet("Transaction_Type").value), "", RsDet("Transaction_Type").value)
                .TextMatrix(j, .ColIndex(("Doctype"))) = IIf(IsNull(RsDet("docType").value), "", RsDet("docType").value)
                .TextMatrix(j, .ColIndex(("CusID"))) = IIf(IsNull(RsDet("CusID").value), "", RsDet("CusID").value)
                .TextMatrix(j, .ColIndex(("FixesAssetsID"))) = IIf(IsNull(RsDet("FixesAssetsID").value), "", RsDet("FixesAssetsID").value)
                .TextMatrix(j, .ColIndex(("Emp_ID"))) = IIf(IsNull(RsDet("Emp_ID").value), "", RsDet("Emp_ID").value)
                .TextMatrix(j, .ColIndex(("DepartementID"))) = IIf(IsNull(RsDet("DepartementID").value), "", RsDet("DepartementID").value)
                

If RsDet("Transaction_ID").value = 49694 Or RsDet("Transaction_ID").value = 55583 Then
.TextMatrix(j, .ColIndex(("Ser"))) = j
End If
                If RsDet("Transaction_Type").value = 19 Then
                     .TextMatrix(j, .ColIndex(("Transaction_TypeName"))) = "«–‰ ’—›"
                ElseIf RsDet("Transaction_Type").value = 992 Or RsDet("Transaction_Type").value = 10 Then
                        .TextMatrix(j, .ColIndex(("Transaction_TypeName"))) = " ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
                ElseIf RsDet("Transaction_Type").value = 992 Or RsDet("Transaction_Type").value = 11 Then
                        .TextMatrix(j, .ColIndex(("Transaction_TypeName"))) = "«” ·«„ „‰ „Œ“‰"
                
                Else
                    .TextMatrix(j, .ColIndex(("Transaction_TypeName"))) = "„»Ì⁄« "
                    .TextMatrix(j, .ColIndex(("CostPrice"))) = Round(IIf(IsNull(RsDet("CostPrice").value), "", RsDet("CostPrice").value), 3)
                '    .TextMatrix(J, .ColIndex(("OldCostPrice"))) = IIf(IsNull(RsDet("OldCostPrice").value), "", RsDet("OldCostPrice").value)
                    
                End If
         
                .TextMatrix(j, .ColIndex(("NoteSerial1"))) = IIf(IsNull(RsDet("NoteSerial1").value), "", RsDet("NoteSerial1").value)
                .TextMatrix(j, .ColIndex(("Transaction_Date"))) = IIf(IsNull(RsDet("Transaction_Date").value), "", RsDet("Transaction_Date").value)
                .TextMatrix(j, .ColIndex(("StoreID"))) = IIf(IsNull(RsDet("StoreID").value), "", RsDet("StoreID").value)
                .TextMatrix(j, .ColIndex(("Fullcode"))) = IIf(IsNull(RsDet("Fullcode").value), "", RsDet("Fullcode").value)
                .TextMatrix(j, .ColIndex(("Item_ID"))) = IIf(IsNull(RsDet("Item_ID").value), "", RsDet("Item_ID").value)
                .TextMatrix(j, .ColIndex(("UnitId"))) = IIf(IsNull(RsDet("UnitId").value), "", RsDet("UnitId").value)
                .TextMatrix(j, .ColIndex(("ShowQty"))) = IIf(IsNull(RsDet("ShowQty").value), "", RsDet("ShowQty").value)
                If val(RsDet!ReCostID & "") = val(TxtSerial1(mIndex).Text) Then
                    If val(RsDet!OldshowPrice & "") <> 0 Then
                        .TextMatrix(j, .ColIndex(("showPrice"))) = Round(IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value), 3)
                        '.TextMatrix(j, .ColIndex(("showPrice"))) = (IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value))
                        .TextMatrix(j, .ColIndex(("OldshowPrice"))) = Round(IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value), 3)
                        '.TextMatrix(j, .ColIndex(("OldshowPrice"))) = (IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value))
                    Else
                        .TextMatrix(j, .ColIndex(("showPrice"))) = Round(IIf(IsNull(RsDet("showPrice").value), 0, RsDet("showPrice").value), 3)
                        .TextMatrix(j, .ColIndex(("showPrice"))) = (IIf(IsNull(RsDet("showPrice").value), 0, RsDet("showPrice").value))
                        .TextMatrix(j, .ColIndex(("OldshowPrice"))) = Round(IIf(IsNull(RsDet("showPrice").value), 0, RsDet("showPrice").value), 3)
                        .TextMatrix(j, .ColIndex(("OldshowPrice"))) = (IIf(IsNull(RsDet("showPrice").value), 0, RsDet("showPrice").value))
                        
                    End If
                Else
                    If val(RsDet!FlgReCost & "") = 1 And val(RsDet!OldshowPrice & "") <> 0 Then
                        .TextMatrix(j, .ColIndex(("showPrice"))) = Round(IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value), 3)
                        '.TextMatrix(j, .ColIndex(("showPrice"))) = (IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value))
                        
                        .TextMatrix(j, .ColIndex(("OldshowPrice"))) = Round(IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value), 3)
                        '.TextMatrix(j, .ColIndex(("OldshowPrice"))) = (IIf(IsNull(RsDet("OldshowPrice").value), 0, RsDet("OldshowPrice").value))
                       
                    Else
                        .TextMatrix(j, .ColIndex(("showPrice"))) = Round(IIf(IsNull(RsDet("showPrice").value), "", RsDet("showPrice").value), 3)
                       ' .TextMatrix(j, .ColIndex(("showPrice"))) = (IIf(IsNull(RsDet("showPrice").value), "", RsDet("showPrice").value))
                        .TextMatrix(j, .ColIndex(("OldshowPrice"))) = Round(IIf(IsNull(RsDet("showPrice").value), "", RsDet("showPrice").value), 3)
                       ' .TextMatrix(j, .ColIndex(("OldshowPrice"))) = (IIf(IsNull(RsDet("showPrice").value), "", RsDet("showPrice").value))
                    End If
                End If
                
                
'                If val(RsDet!OldshowPrice & "") <> 0 Then
'                    .TextMatrix(j, .ColIndex(("OldshowPrice2"))) = IIf(IsNull(RsDet("OldshowPrice2").value), "", RsDet("OldshowPrice").value)
'                Else
'                    .TextMatrix(j, .ColIndex(("OldshowPrice2"))) = IIf(IsNull(RsDet("showPrice").value), "", RsDet("showPrice").value)
'                End If
                .TextMatrix(j, .ColIndex(("Quantity"))) = IIf(IsNull(RsDet("Quantity").value), "", RsDet("Quantity").value)
                .TextMatrix(j, .ColIndex(("Price"))) = IIf(IsNull(RsDet("Price").value), "", RsDet("Price").value)
                
                .TextMatrix(j, .ColIndex(("StoreID2"))) = IIf(IsNull(RsDet("StoreID2").value), "", RsDet("StoreID2").value)
                .TextMatrix(j, .ColIndex(("StoreName2"))) = IIf(IsNull(RsDet("StoreName2").value), "", RsDet("StoreName2").value)
               '---------------------
                If RsDet("Transaction_Type").value = 21 Then
'                    RsDet("OldCostPrice").value = IIf((.TextMatrix(J, .ColIndex("CostPrice"))) = "", Null, val(.TextMatrix(J, .ColIndex("CostPrice"))))
                     If SystemOptions.TypicalProduction = False Then
                       .TextMatrix(j, .ColIndex("CostPrice")) = mCost 'ModItemCostPrice.GetCostItemPrice(.TextMatrix(j, .ColIndex("Item_ID")), 0, , , SystemOptions.SysMainStockCostMethod, , , RsDet("Transaction_Date").value, val(.TextMatrix(j, .ColIndex("Transaction_ID"))), RsDet("UnitID").value, val(RsDet("StoreID").value))
                        If val(.TextMatrix(j, .ColIndex("CostPrice"))) = 0 Then
                            .TextMatrix(j, .ColIndex("CostPrice")) = mCost 'ModItemCostPrice.GetCostItemPrice(.TextMatrix(j, .ColIndex("Item_ID")), 0, , , LastPurPriceType, , , RsDet("Transaction_Date").value, val(.TextMatrix(j, .ColIndex("Transaction_ID"))), val(RsDet("UnitID").value), val(RsDet("StoreID").value))
                        End If
                    Else
                        .TextMatrix(j, .ColIndex("CostPrice")) = 0
'
                    End If
                   ' .TextMatrix(j, .ColIndex("Diff")) = val(RsDet("OldCostPrice").value) - val(RsDet("CostPrice").value)
                 Else
                    '.TextMatrix(j, .ColIndex("showPrice")) = ModItemCostPrice.GetCostItemPrice(val(.TextMatrix(j, .ColIndex("Item_ID"))), 0, "", , SystemOptions.SysMainStockCostMethod, , , IIf((.TextMatrix(j, .ColIndex("Transaction_Date"))) = "", Date, .TextMatrix(j, .ColIndex("Transaction_Date"))), , val(.TextMatrix(j, .ColIndex("UnitId"))), val(.TextMatrix(j, .ColIndex("StoreID"))))
                    .TextMatrix(j, .ColIndex("showPrice")) = mCost
                    .TextMatrix(j, .ColIndex(("Valu"))) = val(.TextMatrix(j, .ColIndex(("showPrice")))) * val(.TextMatrix(j, .ColIndex(("ShowQty"))))
                   ' .TextMatrix(j, .ColIndex("showPrice")) = (IIf(IsNull(RsDet("AvCost").value), 0, RsDet("AvCost").value))
                    '.TextMatrix(j, .ColIndex("OldshowPrice")) = ModItemCostPrice.GetCostItemPrice(val(.TextMatrix(j, .ColIndex("Item_ID"))), 0, "", , SystemOptions.SysMainStockCostMethod, , , IIf((.TextMatrix(j, .ColIndex("Transaction_Date"))) = "", Date, .TextMatrix(j, .ColIndex("Transaction_Date"))), , val(.TextMatrix(j, .ColIndex("UnitId"))), val(.TextMatrix(j, .ColIndex("StoreID"))))
                    'RsDevsub("showPrice").value = IIf((.TextMatrix(J, .ColIndex("showPrice"))) = "", Null, val(.TextMatrix(J, .ColIndex("showPrice"))))
                    'RsDevsub("OldshowPrice").value = IIf((.TextMatrix(J, .ColIndex("OldshowPrice"))) = "", Null, val(.TextMatrix(J, .ColIndex("OldshowPrice"))))
                    'RsDevsub("Price").value = val(.TextMatrix(J, .ColIndex("showPrice")))
                    'RsDevsub("ItemCostPrice").value = val(.TextMatrix(J, .ColIndex("showPrice")))0.
                    '.TextMatrix(j, .ColIndex("Diff")) = Round(val(Round(val(.TextMatrix(j, .ColIndex(("OldshowPrice")))), 3) - Round(val(.TextMatrix(j, .ColIndex("showPrice"))), 3)), 3)
                    .TextMatrix(j, .ColIndex("Diff")) = (val((val(.TextMatrix(j, .ColIndex(("OldshowPrice"))))) - (val(.TextMatrix(j, .ColIndex("showPrice"))))))
'                    RsDevsub("Diff").value = val(.TextMatrix(J, .ColIndex("Diff")))
'                    RsDevsub("Valu").value = val(.TextMatrix(J, .ColIndex("ShowQty"))) * val(.TextMatrix(J, .ColIndex("showPrice")))
                    Dim RsUnitData As ADODB.Recordset
'                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & val(.TextMatrix(J, .ColIndex("Item_ID")))
'                    StrSQL = StrSQL + " AND UnitID=" & val(.TextMatrix(J, .ColIndex("UnitId")))
'                    Set RsUnitData = New ADODB.Recordset
'                    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
'                       RsDevsub("Quantity").value = RsUnitData("UnitFactor").value * val(.TextMatrix(J, .ColIndex("ShowQty")))
'                        RsDevsub("Price").value = val(IIf((.TextMatrix(J, .ColIndex("Price")) = ""), 0, val(.TextMatrix(J, .ColIndex("Price"))))) / RsUnitData("UnitFactor").value
'                    Else
'                        RsDevsub("Price").value = val(IIf((.TextMatrix(J, .ColIndex("Price")) = ""), 0, val(.TextMatrix(J, .ColIndex("Price")))))
'                    End If
                End If
                

                
                
                 grdMaster.TextMatrix(i, grdMaster.ColIndex("OldTotal")) = Round(val(grdMaster.TextMatrix(i, grdMaster.ColIndex("OldTotal"))) + Round((val(grdDet.TextMatrix(j, grdDet.ColIndex("showPrice"))) * val(grdDet.TextMatrix(j, grdDet.ColIndex("ShowQty")))), 3), 3)
                 grdMaster.TextMatrix(i, grdMaster.ColIndex("DiffTotal")) = Round(val(grdMaster.TextMatrix(i, grdMaster.ColIndex("DiffTotal"))) + Round((val(grdDet.TextMatrix(j, grdDet.ColIndex("Diff"))) * val(grdDet.TextMatrix(j, grdDet.ColIndex("ShowQty")))), 3), 3)
               '---------------------DiffTotal
                
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(j, .ColIndex(("StoreName"))) = IIf(IsNull(RsDet("StoreName").value), "", RsDet("StoreName").value)
                    .TextMatrix(j, .ColIndex(("ItemName"))) = IIf(IsNull(RsDet("ItemName").value), "", RsDet("ItemName").value)
                    .TextMatrix(j, .ColIndex(("UnitName"))) = IIf(IsNull(RsDet("UnitName").value), "", RsDet("UnitName").value)
                Else
                    .TextMatrix(j, .ColIndex(("StoreName"))) = IIf(IsNull(RsDet("StoreNamee").value), "", RsDet("StoreNamee").value)
                    .TextMatrix(j, .ColIndex(("ItemName"))) = IIf(IsNull(RsDet("ItemNamee").value), "", RsDet("ItemNamee").value)
                    .TextMatrix(j, .ColIndex(("UnitName"))) = IIf(IsNull(RsDet("UnitNamee").value), "", RsDet("UnitNamee").value)
                End If
                CountNo.Caption = "”ÿ— —ﬁ„ " & j
                Status.Caption = "÷»ÿ «·’‰› " & RsDet!ItemName & " , Õ—ﬂ… —ﬁ„ " & RsDet!NoteSerial1 & ""
                RsDet.MoveNext
               ' .GetSelection J, .ColIndex(("ItemName")), J, .ColIndex(("ItemName"))

            Next j
        End If
        
     
    End With





rs2.MoveNext
 DoEvents
Next i
End If
End With

MsgBox " „ «·«œ—«Ã"




'
'       If BolFrmLoaded = True Then
'            cProgress.StopProgess
'            Set cProgress = Nothing
'        End If

End Sub
Private Sub cmdRecost_Click()
Dim i As Long
Dim mItemNo As Long
Dim mUnitNo As Long
Dim mStoreId As Long
Dim mQty As Double
Dim mTransaction_ID As Long
Dim Transaction_Date As Date
Dim ShowPrice As Double
With grdDet
For i = 1 To .Rows - 1
    
     mItemNo = val(.TextMatrix(i, .ColIndex(("Item_ID"))))
     mUnitNo = val(.TextMatrix(i, .ColIndex(("UnitId"))))
     mStoreId = val(.TextMatrix(i, .ColIndex(("StoreID"))))
    mTransaction_ID = val(.TextMatrix(i, .ColIndex(("Transaction_ID"))))
     Transaction_Date = (.TextMatrix(i, .ColIndex(("Transaction_Date"))))
    
    .TextMatrix(i, .ColIndex(("OldshowPrice2"))) = .TextMatrix(i, .ColIndex(("showPrice")))
     ShowPrice = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, "", , SystemOptions.SysMainStockCostMethod, , , Transaction_Date, mTransaction_ID, mUnitNo, val(mStoreId))
    'showPrice = ModItemCostPrice.GetCostItemPrice(i, 0, , , SystemOptions.SysMainStockCostMethod, , , Transaction_Date, mTransaction_ID, mUnitNo, val(mStoreNo))
    .TextMatrix(i, .ColIndex(("showPrice"))) = ShowPrice
    .TextMatrix(i, .ColIndex(("ItemCostPrice"))) = ShowPrice
    

Next
End With
End Sub



Private Sub ShowBtn_Click()
If Me.TxtModFlg(mIndex).Text <> "R" Then
FillGrid
End If
End Sub

' change id search
Private Sub TxtSerial1_Change(Index As Integer)
    Dim TxtMod As String
    TxtMod = Me.TxtModFlg(mIndex)
    Me.TxtModFlg(mIndex) = ""
    TxtModFlg(mIndex) = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click mIndex
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click(Index As Integer)
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click(Index As Integer)
    FindRec val(TxtSerial1(Index).Text)
    Me.TxtModFlg(mIndex).Text = "R"
    FiLLTXT
     BtnLast_Click mIndex
End Sub
' delet sub
Private Sub btnDelete_Click(Index As Integer)
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Long
    Dim ID As Double
    If TxtNoteSerial(mIndex).Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1(mIndex).Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
      Dim StrSQL As String
      If mIndex = 0 Then
            With FG
              For i = 1 To .Rows - 1
                    Cn.Execute "Update dbo.Transaction_Details set FlgReCost=Null,ReCostID=Null  where ID=" & val(.TextMatrix(i, .ColIndex("IDRef"))) & ""
              Next i
             End With
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 1

        Else
        
            With grdDet
                For i = 1 To .Rows - 1
                    Cn.Execute "Update dbo.Transaction_Details set FlgReCost=Null,ReCostID=Null,showPrice = " & IIf((.TextMatrix(i, .ColIndex("OldshowPrice"))) = "", ShowPrice, val(.TextMatrix(i, .ColIndex("OldshowPrice")))) & "  where ID=" & val(.TextMatrix(i, .ColIndex("IDRef"))) & ""
                Next i
            End With
            grdMaster.Clear flexClearScrollable, flexClearEverything
            grdDet.Clear flexClearScrollable, flexClearEverything
            StrSQL = "delete from   " & mTableName3 & "  where ReCostID =" & val(TxtSerial1(mIndex).Text) & ""
            Cn.Execute StrSQL
             
            StrSQL = "delete DOUBLE_ENTREY_VOUCHERS WHERE advanceID =" & val(TxtSerial1(mIndex).Text)
            Cn.Execute StrSQL
   
    
        End If
        
        StrSQL = "delete from   " & mTableName2 & "  where ReCostID =" & val(TxtSerial1(mIndex).Text) & ""
        Cn.Execute StrSQL
        
         RsSavRec.find "ID=" & val(TxtSerial1(Index).Text), , adSearchForward, 1
         RsSavRec.delete
     
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
        
     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
     LabCurrRec(mIndex).Caption = 0
     LabCountRec(mIndex).Caption = 0
       ' FillGridWithData
        BtnNext_Click mIndex
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           'Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg(mIndex).Text <> "R" Then
        Select Case Me.TxtModFlg(mIndex).Text
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
               btnSave_Click mIndex
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change(Index As Integer)
    If Me.TxtModFlg(mIndex) = "N" Then
    XPDtbTrans(mIndex).Enabled = True

        Me.btnNew(mIndex).Enabled = False
        btnSave(mIndex).Enabled = False
        btnDelete(mIndex).Enabled = False
        'Me.btnQuery.Enabled = False
      '  ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo(mIndex).Enabled = True
        Me.btnSave(mIndex).Enabled = True
      '  BtnUpdate.Enabled = False

    ElseIf Me.TxtModFlg(mIndex) = "R" Then
     XPDtbTrans(mIndex).Enabled = False
        btnSave(mIndex).Enabled = False
        btnDelete(mIndex).Enabled = False
        If TxtSerial1(mIndex).Text <> "" Then
            btnSave(mIndex).Enabled = True
            btnDelete(mIndex).Enabled = True
             btnModify(Index).Enabled = True
    End If
     '   BtnUpdate.Enabled = True
     '   Me.btnQuery.Enabled = True
        Me.btnNew(mIndex).Enabled = True
        BtnUndo(mIndex).Enabled = False
        Me.btnSave(mIndex).Enabled = False
    '    ISButton1.Enabled = True
        btnNext(mIndex).Enabled = True
        btnPrevious(mIndex).Enabled = True
        btnFirst(mIndex).Enabled = True
        btnLast(mIndex).Enabled = True
   ElseIf Me.TxtModFlg(mIndex) = "E" Then
    XPDtbTrans(mIndex).Enabled = True
        Me.btnNew(mIndex).Enabled = False
        btnSave(mIndex).Enabled = False
        btnDelete(mIndex).Enabled = False
       btnModify(mIndex).Enabled = False
'        BtnUpdate.Enabled = False
        BtnUndo(mIndex).Enabled = True
        Me.btnSave(mIndex).Enabled = True
    '    Grid.Enabled = False
        btnNext(mIndex).Enabled = False
        btnPrevious(mIndex).Enabled = False
        btnFirst(mIndex).Enabled = False
        btnLast(mIndex).Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(1).Text)
        Me.TxtModFlg(mIndex).Text = "R"
    End If
    TxtModFlg(mIndex) = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me

        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst

    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        Me.TxtModFlg(mIndex).Text = "R"
    End If
    TxtModFlg(mIndex) = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me

        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast

    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click(Index As Integer)
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtNoteSerial(mIndex).Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
    Else
    MsgBox "Please Delete Voucher"
    End If
    Exit Sub
    End If
    If TxtSerial1(mIndex).Text <> "" Then
        TxtModFlg(mIndex) = "E"

        Me.DCboUserName(mIndex).BoundText = user_id
        Me.Dcbranch(mIndex).SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"

            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click(Index As Integer)
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    Me.TxtModFlg(mIndex) = "N"
    FG.Clear flexClearScrollable, flexClearEverything
    grdMaster.Clear flexClearScrollable, flexClearEverything
    grdDet.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    grdDet.Rows = 2
    
    grdMaster.Rows = 2
    Me.DCboUserName(mIndex).BoundText = user_id
    Me.Dcbranch(mIndex).BoundText = Current_branch
    Dcbranch(mIndex).SetFocus
    XPDtbTrans(mIndex).value = Date
        ToDate(mIndex) = Date
ErrTrap:
End Sub
Private Sub BtnNext_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(1).Text)
        Me.TxtModFlg(mIndex).Text = "R"
    End If
    TxtModFlg(mIndex) = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me

        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        Me.TxtModFlg(mIndex).Text = "R"
    End If
    TxtModFlg(mIndex) = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me

        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub


'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg(mIndex).Text = "R" Then
            btnNew_Click mIndex
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew(mIndex).Enabled = False Then Exit Sub
        btnNew_Click mIndex
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnSave(mIndex).Enabled = False Then Exit Sub
        btnModify_Click mIndex
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave(mIndex).Enabled = False Then Exit Sub
        btnSave_Click mIndex
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo(mIndex).Enabled = False Then Exit Sub
        BtnUndo_Click mIndex
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete(mIndex).Enabled = False Then Exit Sub
        btnDelete_Click mIndex
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel(mIndex).Enabled = False Then Exit Sub
            BtnCancel_Click mIndex
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If Me.TxtModFlg(mIndex) = "R" Then
    'Move first --------------------------------------------
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
'        If btnFirst(mIndex).Enabled = False Then Exit Sub
'        BtnFirst_Click mIndex
'    End If
'    'Move Previous---------------------------------------------------------
'    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
'        If btnPrevious(mIndex).Enabled = False Then Exit Sub
'        BtnPrevious_Click mIndex
'    End If
'    'Move Next---------------------------------------------------------
'    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
'        If btnNext(mIndex).Enabled = False Then Exit Sub
'        BtnNext_Click mIndex
'    End If
'    'Move Last---------------------------------------------------------
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
'        If btnLast(mIndex).Enabled = False Then Exit Sub
'        BtnLast_Click mIndex
'    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
       Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst(mIndex).ButtonImage
    Set Me.btnFirst(mIndex).ButtonImage = Me.btnLast(mIndex).ButtonImage
    Set Me.btnLast(mIndex).ButtonImage = XPic
    Set XPic = Me.btnPrevious(mIndex).ButtonImage
    Set Me.btnPrevious(mIndex).ButtonImage = Me.btnNext(mIndex).ButtonImage
    Set Me.btnNext(mIndex).ButtonImage = XPic
   ''''''''''''''''''''////
       Me.Caption = "Recalculate The Cost"
      Label1(2).Caption = Me.Caption
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(21).Caption = "Remarks"
      Cmd(0).Caption = "Delete"
      Cmd(1).Caption = "Delete All"
    ISButton5(mIndex).Caption = "Print"
    ISButton8(mIndex).Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew(mIndex).Caption = "New"
    btnSave(mIndex).Caption = "Modify"
    btnSave(mIndex).Caption = "Save"
    BtnUndo(mIndex).Caption = "Undo"
    'BtnUpdate.Caption = "Refresh "
'    ISButton1.Caption = "Print"
'    btnQuery.Caption = "Search"
    btnDelete(mIndex).Caption = "Delete"
    btnCancel(mIndex).Caption = "Exit"

  With Me.FG
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Std")) = "Std "
  .TextMatrix(0, .ColIndex("BoxNo")) = "Box No"
  .TextMatrix(0, .ColIndex("C1")) = "1"
  .TextMatrix(0, .ColIndex("C2")) = "2"
  .TextMatrix(0, .ColIndex("C3")) = "3"
  .TextMatrix(0, .ColIndex("C4")) = "4"
  .TextMatrix(0, .ColIndex("C5")) = "5"
  .TextMatrix(0, .ColIndex("C6")) = "6"
  .TextMatrix(0, .ColIndex("C7")) = "7"
  .TextMatrix(0, .ColIndex("C8")) = "8"
  .TextMatrix(0, .ColIndex("C9")) = "9"
  End With

ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = mTableName
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).Text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end






Public Sub CreateIssueVoucher()

    On Error GoTo errortrap
    'DeleteTransactiomsVoucher Val(Text1.text)




    Dim i As Long
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    

    Dim RsTrans As ADODB.Recordset
    Dim rsTrans2 As ADODB.Recordset
    Dim rsTransDet As ADODB.Recordset
    Dim s As String
    Dim mOrderId As Long
    
    Dim sql As String
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    Dim TxtNoteSerialV As String
    Dim mDate As Date
    Dim mStoreId As Integer
    Dim mUserID As Long
    Dim xyeas As Boolean
    Dim Transaction_ID As Long
    
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
    Dim mNoteSerial1 As String
    Dim TxtNoteSerial1V As String
    Dim mTotalCost As Double
    Dim CurrentVoucherNo As String
ll:
Dim TransBegine As Boolean

Cn.BeginTrans
TransBegine = True
 For i = 1 To grdMaster2.Rows - 1
    mOrderId = val(grdMaster2.TextMatrix(i, grdMaster2.ColIndex("Transaction_ID")))
    s = " select * from Transactions WHERE Transaction_ID =  " & mOrderId
    Set RsTrans = New ADODB.Recordset
    RsTrans.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If Not RsTrans.EOF Then
        s = " select * from Transactions WHERE Transaction_ID  = " & val(RsTrans!nots & "")
        Set rsTrans2 = New ADODB.Recordset
        
        rsTrans2.Open s, Cn, adOpenKeyset, adLockOptimistic
        Do While Not rsTrans2.EOF
            Cn.Execute " Delete DOUBLE_ENTREY_VOUCHERS where Notes_ID = " & val(rsTrans2!NoteID & "")
            Cn.Execute " Delete Notes where NoteId = " & val(rsTrans2!NoteID & "")
            Cn.Execute " Delete Transaction_Details where Transaction_ID = " & val(RsTrans!nots & "")
            
            
            rsTrans2.MoveNext
            
        Loop
        Cn.Execute "Delete Transactions where Transaction_ID = " & val(RsTrans!nots & "") & " and Transaction_Type=19"
        
        
        
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=19"))
        
   
            
        my_branch = val(RsTrans!BranchID & "")
        mDate = (RsTrans!Transaction_Date & "")
        TxtNoteSerialV = Notes_coding(val(my_branch), mDate)
        mStoreId = val(RsTrans!StoreID & "")
        mUserID = val(RsTrans!UserID & "")
        mNoteSerial1 = Trim(RsTrans!NoteSerial1 & "")
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        TxtNoteSerial1V = Voucher_coding(val(my_branch), mDate, 10, 180, , 19, , mStoreId, , , , mUserID)
        CurrentVoucherNo = ""
        
        s = " select Sum(IsNull(CostPrice,0)* IsNull(ShowQty,0)) as CostPrice from Transaction_Details Where Transaction_ID  = " & mOrderId
        Set rsTransDet = New ADODB.Recordset
        rsTransDet.Open s, Cn, adOpenKeyset, adLockOptimistic
        If Not rsTransDet.EOF Then
            mTotalCost = val(rsTransDet!costPrice & "")
        End If
        
        With grdMaster2
            .TextMatrix(i, .ColIndex("StoreID")) = mStoreId
            .TextMatrix(i, .ColIndex("StoreID2")) = mStoreId
    
            'mItemCode = (.TextMatrix(mRow, .ColIndex("Item_ID")))
            .TextMatrix(i, .ColIndex("DiffTotal")) = mTotalCost
            
            .TextMatrix(i, .ColIndex("OldTotal")) = 0
            .TextMatrix(i, .ColIndex("Transaction_ID2")) = Transaction_ID
    '        mNoteSerial = val(.TextMatrix(mRow, .ColIndex("NoteSerial")))
            .TextMatrix(i, .ColIndex("NoteID")) = 0
            .TextMatrix(i, .ColIndex("Doctype")) = val(RsTrans!docType & "")
            .TextMatrix(i, .ColIndex("CusID")) = val(RsTrans!CusID & "")
            .TextMatrix(i, .ColIndex("NoteSerial12")) = TxtNoteSerial1V
            .TextMatrix(i, .ColIndex("FixesAssetsID")) = val(RsTrans!FixesAssetsID & "")
            
            'mDate = (.TextMatrix(mRow, .ColIndex("Transaction_Date")))
            .TextMatrix(i, .ColIndex("Emp_ID")) = val(RsTrans!Emp_id & "")
            .TextMatrix(i, .ColIndex("DepartementID")) = val(RsTrans!Departementid & "")
            .TextMatrix(i, .ColIndex("BranchId")) = my_branch
       ' mTransType = val(.TextMatrix(mRow, .ColIndex("Transaction_Type")))
        End With
        

        sql = " INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed,ManualNO,CashCustomerName,Ser,SessionD )SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & val(mOrderId) & ",nots2='" & mNoteSerial1 & "' ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1,ManualNO ,CashCustomerName,Ser = " & 1 & ",1 From Transactions Where  Transaction_ID =" & mOrderId & " And Transaction_Type = 21"
        Cn.Execute sql
        
        
        
        
        If SystemOptions.RawMaterMix = False Then
            Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost,MixNo,QtyFaqtors,MaxQty,MaxUnitID ,length,OUTR,INR,Height,Width,NoCount)SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost,MixNo,QtyFaqtors,MaxQty,MaxUnitID,length,OUTR,INR,Height,Width,NoCount From dbo.Transaction_Details Where SavedItemType=0 and   Transaction_ID = " & mOrderId
        Else
            FillMixToVoucher Transaction_ID, mStoreId, mOrderId
        End If
        
       ' Text1.Text = Transaction_ID
        UpdateTransactionsCost CStr(Transaction_ID)
           
        'TxtIssueSerial.text = TxtNoteSerial1V
 
        StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & ",SessionD = 1 WHERE Transaction_ID=" & mOrderId
        Cn.Execute StrSQL
        
        
        
           CreateVouc i, 1
            DoEvents
        
    End If
    
    
        CountNo2.Caption = "”ÿ— —ﬁ„ " & i
        Status2.Caption = "÷»ÿ «·Õ—ﬂ… —ﬁ„ " & mNoteSerial1
 
 Next i
TransBegine = False
Cn.CommitTrans
    Dim groupAccount  As String

'    If detect_inventory_work_type = 3 Then
'
'        With FG
'
'            For i = 1 To FG.Rows - 1
'
'                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
'
'                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
'                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), val(DCboStoreName.BoundText), 0)
'
'                    If groupAccount = "Error" Then
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
'                        Else
'                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
'                        End If
'
'                        Exit Sub
'                    End If
'                End If
'
'            Next i
'
'        End With
'
'    End If


 
    

      '  If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
      '      TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
      '      TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
      '  End If

        
  
'        Text1.text = Transaction_ID
'
'        If SystemOptions.TypicalProduction = True Then
'            Exit Sub
'        End If
'
'        'Create big notes
' Dim NoteID As Long
'  Dim NoteDate As Date
'    Dim NoteSerial As String
'    Dim Notevalue As Double
'    Dim des As String
'If CurrentVoucherNo <> "" Then
'NoteSerial = CurrentVoucherNo
'End If
'
'
''*****************************************************************
'    Dim TOTAL_COST As Double
'    With FG
'
'        For i = 1 To FG.Rows - 1
'
'            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
'                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
'                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
'
'                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
'
'                '           TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))
'                        'CostPrice
'                TOTAL_COST = TOTAL_COST + val(FG.TextMatrix(i, FG.ColIndex("ItemCostPrice"))) * FG.TextMatrix(i, FG.ColIndex("Count"))
'            End If
'
'        Next i
'
'    End With
'    '*****************************************************************
'
' CreateNotes NoteID, (XPDtbBill.value), val(Dcbranch.BoundText), 180, TOTAL_COST, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID, TxtNoteSerial1V, ToHijriDate(XPDtbBill.value)
'          ' TxtNoteID.text = NoteID
'           general_noteid = NoteID
'
'        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.Dcbranch.BoundText)
'
'    End If
'
    '
' End If
MsgBox " „ «‰‘«¡ ”‰œ«  «·’—›"
Exit Sub
errortrap:
    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

End Sub



Sub FillMixToVoucher(Optional Transaction_ID As Long, Optional ByRef StoreID3 As Integer, Optional Transaction_ID2 As Long)
  Dim Rs3 As ADODB.Recordset
  Dim RSTransDetails As ADODB.Recordset
  Dim RsUnitData As ADODB.Recordset
  Dim LngCurItemID As Long
  Dim LngUnitID As Long
  Dim DblQty As Double
  Dim ItemID2 As Double
  Dim StrSQL As String
  Dim RowNum As Integer
   StrSQL = "SELECT     ItemID, UnitId, Qty, Cost"
   StrSQL = StrSQL & " From dbo.TblSalesMixItems"
   StrSQL = StrSQL & " Where (TransectionID = " & Transaction_ID2 & ")"
   ' And (StoreID = " & StoreID3 & ")"
  ' StrSQL = StrSQL & " GROUP BY ItemID, UnitId, Qty, Cost"
   Set Rs3 = New ADODB.Recordset
   Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
     Set RSTransDetails = New ADODB.Recordset
     StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
     RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For RowNum = 1 To Rs3.RecordCount
        ItemID2 = IIf(IsNull(Rs3("ItemID").value), 0, Rs3("ItemID").value)
            If ItemID2 <> 0 Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = ItemID2
                RSTransDetails("Quantity").value = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
                RSTransDetails("SHOWQTY").value = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
                RSTransDetails("showPrice").value = IIf(IsNull(Rs3("Cost").value), 0, Rs3("Cost").value)
                RSTransDetails("UnitID").value = IIf(IsNull(Rs3("UnitId").value), 0, Rs3("UnitId").value)
               ' RSTransDetails("MixNo").value = IIf(IsNull(Rs3("MixNo").value), "", Rs3("MixNo").value)
               ' RSTransDetails("QtyFaqtors").value = IIf(IsNull(Rs3("QtyFaqtors").value), 0, Rs3("QtyFaqtors").value)
               ' RSTransDetails("MaxQty").value = IIf(IsNull(Rs3("MaxQty").value), 0, Rs3("MaxQty").value)
               ' RSTransDetails("MaxUnitID").value = IIf(IsNull(Rs3("MaxUnitID").value), 0, Rs3("MaxUnitID").value)
        
                RSTransDetails("StoreID2").value = StoreID3
                          '«·ÊÕœ« 
            LngCurItemID = ItemID2
            LngUnitID = IIf(IsNull(Rs3("UnitId").value), 0, Rs3("UnitId").value)
            DblQty = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value) / RSTransDetails("QtyBySmalltUnit").value
            
            End If
                RSTransDetails.update
            End If
     Rs3.MoveNext
        Next RowNum
    End If
End Sub


