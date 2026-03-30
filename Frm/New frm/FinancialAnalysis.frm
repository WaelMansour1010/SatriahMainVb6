VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FinancialAnalysis 
   Caption         =   "«⁄œ«œ  „⁄«œ·«  «·‰Õ·Ì· «·„«·Ì"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   16695
   Icon            =   "FinancialAnalysis.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   16695
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16695
      _cx             =   29448
      _cy             =   16748
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
      AutoSizeChildren=   8
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FinancialAnalysis.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8910
         Index           =   4
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   16665
         _cx             =   29395
         _cy             =   15716
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
            Height          =   8850
            Index           =   6
            Left            =   11475
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   30
            Width           =   5160
            _cx             =   9102
            _cy             =   15610
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
               Left            =   1920
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
                     Picture         =   "FinancialAnalysis.frx":040E
                     Key             =   "Expanded_Node"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FinancialAnalysis.frx":1260
                     Key             =   "Root"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FinancialAnalysis.frx":15FA
                     Key             =   "Open_Node"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FinancialAnalysis.frx":1994
                     Key             =   "Closed_Node"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FinancialAnalysis.frx":1D2E
                     Key             =   "Item"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.TreeView TrvAccounts 
               Height          =   6630
               HelpContextID   =   380
               Left            =   0
               TabIndex        =   5
               Top             =   30
               Width           =   5085
               _ExtentX        =   8969
               _ExtentY        =   11695
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   706
               LabelEdit       =   1
               Style           =   7
               Checkboxes      =   -1  'True
               ImageList       =   "ImgLstChartTree"
               Appearance      =   1
            End
            Begin VB.Label lbldown 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00FF0000&
               Height          =   615
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   8160
               Width           =   4575
            End
            Begin VB.Label lblUp 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   7320
               Width           =   4575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "‘ﬂ· «·„⁄«œ·Â"
               Height          =   615
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   6720
               Width           =   4575
            End
            Begin VB.Line Line1 
               BorderWidth     =   3
               X1              =   4920
               X2              =   240
               Y1              =   8040
               Y2              =   8040
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8850
            Index           =   3
            Left            =   30
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   11415
            _cx             =   20135
            _cy             =   15610
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
            AutoSizeChildren=   8
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
            GridRows        =   10
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FinancialAnalysis.frx":20C8
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Frame7 
               Height          =   870
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   30
               Width           =   11355
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   345
                  Index           =   0
                  Left            =   1065
                  TabIndex        =   73
                  Top             =   240
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
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
                  ButtonImage     =   "FinancialAnalysis.frx":2191
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   345
                  Index           =   2
                  Left            =   0
                  TabIndex        =   74
                  Top             =   240
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
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
                  ButtonImage     =   "FinancialAnalysis.frx":252B
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   345
                  Index           =   1
                  Left            =   1590
                  TabIndex        =   75
                  Top             =   240
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
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
                  ButtonImage     =   "FinancialAnalysis.frx":28C5
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   345
                  Index           =   3
                  Left            =   525
                  TabIndex        =   76
                  Top             =   240
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
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
                  ButtonImage     =   "FinancialAnalysis.frx":2C5F
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin VB.Label LblHeader 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "  «⁄œ«œ  „⁄«œ·«  «· Õ·Ì· «·„«·Ì"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   24
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404000&
                  Height          =   585
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   120
                  Width           =   11355
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1740
               Index           =   1
               Left            =   30
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   915
               Width           =   11355
               _cx             =   20029
               _cy             =   3069
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
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.TextBox XPTxtID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   8880
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   300
                  Width           =   1095
               End
               Begin VB.CheckBox ChKBlock 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   1860
                  Width           =   1215
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ»Ì⁄Â «·—’Ìœ"
                  ForeColor       =   &H000000C0&
                  Height          =   1335
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   2160
                  Width           =   5055
                  Begin VB.Frame Frame5 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   375
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   720
                     Width           =   3495
                     Begin VB.OptionButton Differenttype 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " Õ–Ì— ›ﬁÿ"
                        Height          =   195
                        Index           =   1
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   65
                        Top             =   120
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
                        TabIndex        =   64
                        Top             =   120
                        Width           =   1725
                     End
                  End
                  Begin VB.Frame Frame3 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " "
                     Height          =   375
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   240
                     Width           =   3495
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„œÌ‰"
                        Height          =   195
                        Index           =   0
                        Left            =   1800
                        RightToLeft     =   -1  'True
                        TabIndex        =   62
                        Top             =   120
                        Width           =   1365
                     End
                     Begin VB.OptionButton DepitOrCredit 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "œ«∆‰"
                        Height          =   195
                        Index           =   1
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   61
                        Top             =   120
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
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   720
                     Width           =   1215
                  End
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·Â „Ê«“‰‹‹‹Â"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   1860
                  Width           =   1695
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’·«ÕÌ… «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   1215
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   2280
                  Width           =   4935
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„” Œœ„"
                     Height          =   195
                     Index           =   2
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   840
                     Value           =   -1  'True
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ã„Ê⁄Â"
                     Height          =   195
                     Index           =   1
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   480
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ· «·„” Œœ„Ì‰"
                     Height          =   195
                     Index           =   0
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   240
                     Width           =   1845
                  End
                  Begin MSDataListLib.DataCombo DataCombo1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   55
                     Top             =   480
                     Width           =   3135
                     _ExtentX        =   5530
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   56
                     Top             =   840
                     Width           =   3135
                     _ExtentX        =   5530
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
                  Height          =   1095
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   3000
                  Width           =   4935
                  Begin VB.CheckBox Check2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·Â „—ﬂ“  ﬂ·›Â"
                     Height          =   255
                     Left            =   3480
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·„—ﬂ“"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   120
                     Width           =   3135
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
                     Left            =   120
                     TabIndex        =   44
                     Top             =   720
                     Width           =   3015
                     _ExtentX        =   5318
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
                     Left            =   3360
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   720
                     Width           =   1215
                  End
               End
               Begin VB.TextBox txtFinancialEquationsDes 
                  Alignment       =   1  'Right Justify
                  Height          =   675
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   35
                  Top             =   720
                  Width           =   9930
               End
               Begin MSDataListLib.DataCombo DboParentAccount 
                  Height          =   315
                  Left            =   5115
                  TabIndex        =   31
                  Top             =   2340
                  Width           =   3375
                  _ExtentX        =   5953
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
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   -210
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1095
                  Index           =   5
                  Left            =   120
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   4905
                  _cx             =   8652
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
                  ForeColor       =   192
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "‰Ê⁄ «·Õ”«»"
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
                  Begin VB.ComboBox Combo1 
                     Height          =   315
                     ItemData        =   "FinancialAnalysis.frx":2FF9
                     Left            =   2520
                     List            =   "FinancialAnalysis.frx":3006
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     ItemData        =   "FinancialAnalysis.frx":3023
                     Left            =   0
                     List            =   "FinancialAnalysis.frx":3036
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.CheckBox Check3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«»  Ã„Ì⁄Ì"
                     Height          =   210
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» ‰Â«∆Ï"
                     Height          =   210
                     Index           =   0
                     Left            =   1710
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   225
                     Value           =   -1  'True
                     Width           =   1215
                  End
                  Begin VB.OptionButton OptAccountType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» —∆Ì”ÌÏ"
                     Height          =   195
                     Index           =   1
                     Left            =   3270
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ»Ì⁄Â «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   300
                     Index           =   3
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   705
                     Width           =   1230
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " »ÊÌ» «·Õ”«»"
                     ForeColor       =   &H00000000&
                     Height          =   300
                     Index           =   9
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   705
                     Width           =   990
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   1
                     Left            =   4620
                     Picture         =   "FinancialAnalysis.frx":3062
                     Top             =   225
                     Width           =   240
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   0
                     Left            =   2940
                     Picture         =   "FinancialAnalysis.frx":33EC
                     Top             =   225
                     Width           =   240
                  End
               End
               Begin VB.TextBox XPTxtName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2910
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   300
                  Width           =   4650
               End
               Begin VB.TextBox TxtGroupCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin VB.TextBox TxtAccount_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   -360
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin MSDataListLib.DataCombo DCCURRENCY 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   34
                  Top             =   -930
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "  ·Õ–› «Ì ”ÿ— «÷›ÿ Delete "
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
                  Height          =   255
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   1440
                  Width           =   3735
               End
               Begin VB.Label lbl«”„«·„⁄«œ·Â 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰”» «·„À«·Ì…"
                  Height          =   285
                  Index           =   10
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   360
                  Width           =   1230
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
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„·…"
                  Height          =   195
                  Index           =   8
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   -480
                  Width           =   510
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Name En"
                  Height          =   285
                  Index           =   7
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   2070
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—Õ «·„⁄«œ·Â"
                  Height          =   435
                  Index           =   6
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   780
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„⁄«œ·Â"
                  Height          =   285
                  Index           =   2
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   300
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·„⁄«œ·Â"
                  Height          =   345
                  Index           =   0
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   300
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ”«»"
                  Height          =   345
                  Index           =   1
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   -120
                  Visible         =   0   'False
                  Width           =   930
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   870
               Index           =   7
               Left            =   4605
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   7950
               Width           =   3390
               _cx             =   5980
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
                  Top             =   630
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
                  ButtonImage     =   "FinancialAnalysis.frx":3776
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   255
                  Index           =   1
                  Left            =   -360
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
                  ButtonImage     =   "FinancialAnalysis.frx":3B10
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdN 
                  Height          =   285
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   28
                  Top             =   30
                  Visible         =   0   'False
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
                  ButtonImage     =   "FinancialAnalysis.frx":3EAA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Image Image6 
                  Height          =   720
                  Left            =   0
                  Picture         =   "FinancialAnalysis.frx":4244
                  Top             =   0
                  Width           =   765
               End
               Begin VB.Image Image5 
                  Height          =   720
                  Left            =   840
                  Picture         =   "FinancialAnalysis.frx":49B0
                  Top             =   0
                  Width           =   810
               End
               Begin VB.Image Image4 
                  Height          =   720
                  Left            =   1680
                  Picture         =   "FinancialAnalysis.frx":5270
                  Top             =   0
                  Width           =   810
               End
               Begin VB.Image Image3 
                  Height          =   720
                  Left            =   2520
                  Picture         =   "FinancialAnalysis.frx":584E
                  Top             =   0
                  Width           =   705
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   870
               Index           =   2
               Left            =   30
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   7950
               Width           =   11355
               _cx             =   20029
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
               Begin VB.TextBox TxtFinancialEquationsOpr 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   7680
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1200
               End
               Begin VB.TextBox TxtGeneralValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   0
                  Width           =   1200
               End
               Begin VB.TextBox TxtModflg 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   10830
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   8
                  Left            =   -210
                  TabIndex        =   79
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ—"
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   750
                  Index           =   8
                  Left            =   3120
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1590
                  _cx             =   2805
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
                     Index           =   3
                     Left            =   30
                     TabIndex        =   83
                     Top             =   1350
                     Visible         =   0   'False
                     Width           =   525
                     _ExtentX        =   926
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
                     ButtonImage     =   "FinancialAnalysis.frx":5F3B
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdN 
                     Height          =   255
                     Index           =   4
                     Left            =   -1470
                     TabIndex        =   84
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
                     ButtonImage     =   "FinancialAnalysis.frx":62D5
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmdN 
                     Height          =   285
                     Index           =   5
                     Left            =   1320
                     TabIndex        =   85
                     Top             =   30
                     Visible         =   0   'False
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
                     ButtonImage     =   "FinancialAnalysis.frx":666F
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Image Image7 
                     Height          =   720
                     Left            =   120
                     Picture         =   "FinancialAnalysis.frx":6A09
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   690
                  End
               End
               Begin VB.Image Image8 
                  Height          =   720
                  Left            =   360
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Index           =   100
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   0
                  Width           =   600
               End
               Begin VB.Image Image1 
                  Height          =   240
                  Left            =   10785
                  Picture         =   "FinancialAnalysis.frx":6FE7
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   240
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ã“¡ «·À«» "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   255
                  Index           =   4
                  Left            =   7755
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   30
                  Width           =   1320
               End
               Begin VB.Label lbl 
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
                  Height          =   555
                  Index           =   5
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   -1485
                  Width           =   11295
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts 
               Height          =   2625
               Left            =   30
               TabIndex        =   50
               Top             =   2670
               Width           =   7965
               _cx             =   14049
               _cy             =   4630
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
               BackColorBkg    =   16777215
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
               Rows            =   3
               Cols            =   10
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FinancialAnalysis.frx":7371
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2625
               Index           =   9
               Left            =   8010
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   2670
               Width           =   3375
               _cx             =   5953
               _cy             =   4630
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
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   855
                  Index           =   1
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   120
                  Width           =   975
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   975
                  Index           =   3
                  Left            =   1455
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   1680
                  Width           =   960
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   1215
                  Index           =   2
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   840
                  Width           =   1110
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   975
                  Index           =   4
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   960
                  Width           =   990
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   1095
                  Index           =   0
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   0
                  Width           =   975
               End
               Begin VB.Image Image2 
                  Height          =   2475
                  Index           =   0
                  Left            =   120
                  Picture         =   "FinancialAnalysis.frx":7530
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   3090
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2625
               Index           =   10
               Left            =   8010
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   5310
               Width           =   3375
               _cx             =   5953
               _cy             =   4630
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
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   1095
                  Index           =   9
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   0
                  Width           =   975
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   975
                  Index           =   8
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   960
                  Width           =   990
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   1215
                  Index           =   7
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   840
                  Width           =   1110
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   975
                  Index           =   6
                  Left            =   1455
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   1680
                  Width           =   960
               End
               Begin VB.Label Opr 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   855
                  Index           =   5
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   120
                  Width           =   975
               End
               Begin VB.Image Image2 
                  Height          =   2475
                  Index           =   1
                  Left            =   120
                  Picture         =   "FinancialAnalysis.frx":97D8
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   3090
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2625
               Left            =   30
               TabIndex        =   99
               Top             =   30
               Width           =   7965
               _cx             =   14049
               _cy             =   4630
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
               Rows            =   3
               Cols            =   10
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FinancialAnalysis.frx":BA80
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
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts1 
               Height          =   2625
               Left            =   30
               TabIndex        =   100
               Top             =   5310
               Width           =   7965
               _cx             =   14049
               _cy             =   4630
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
               BackColorBkg    =   16777215
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
               Rows            =   3
               Cols            =   10
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FinancialAnalysis.frx":BC44
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Currency"
               Height          =   1740
               Left            =   4605
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   915
               Width           =   3390
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„·… «·Õ”«»"
               Height          =   870
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   30
               Width           =   1980
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   0
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   8940
         Width           =   16665
         _cx             =   29395
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
            Left            =   14790
            TabIndex        =   7
            Top             =   90
            Width           =   1560
            _ExtentX        =   2752
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
            Left            =   12930
            TabIndex        =   8
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
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
            Left            =   11055
            TabIndex        =   9
            Top             =   90
            Width           =   1710
            _ExtentX        =   3016
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
            Left            =   9570
            TabIndex        =   10
            Top             =   90
            Width           =   1380
            _ExtentX        =   2434
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
            Left            =   7125
            TabIndex        =   11
            Top             =   90
            Width           =   2355
            _ExtentX        =   4154
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
            Left            =   5625
            TabIndex        =   12
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
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
            Left            =   255
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   90
            Width           =   1590
            _ExtentX        =   2805
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
            Left            =   3750
            TabIndex        =   14
            Top             =   90
            Width           =   1680
            _ExtentX        =   2963
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
            Left            =   1905
            TabIndex        =   15
            Top             =   90
            Width           =   1695
            _ExtentX        =   2990
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
      End
   End
End
Attribute VB_Name = "FinancialAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Dim CurrentNode As MSComctlLib.Node
Dim AccountCode As String
Dim AccountName As String

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim XNode As MSComctlLib.Node

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        If XPTxtName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„⁄«œ·Â"
            Else
                Msg = "plz enter group name firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtName.SetFocus
            Exit Sub
        End If
   
        Select Case TxtModFlg.text

            Case "N"
                StrSQL = "select * From FinancialEquations where FinancialEquationsName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÊÃœ „⁄«œ·Â „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„⁄«œ·Â"
                
                    Else
                        Msg = "This group Name Already Exisi" & Chr(13)
            
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            
                XPTxtID.text = CStr(new_id("FinancialEquations", "FinancialEquationsId", "", True))
                Me.TxtGroupCode.text = XPTxtID.text

            Case "E"
                StrSQL = "select * From FinancialEquations where FinancialEquationsName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("FinancialEquationsId").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÊÃœ „⁄«œ·Â „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„⁄«œ·Â"
                        Else
                            Msg = "This group Name Already Exisi" & Chr(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If

        End Select
     
        Select Case TxtModFlg.text

            Case "N"
                Cn.BeginTrans
                BeginTrans = True
            
                rs.AddNew
                rs("FinancialEquationsId").value = IIf(XPTxtID.text = "", "", val(XPTxtID.text))
            
            Case "E"
 
                Cn.BeginTrans
                BeginTrans = True
        End Select

        rs("FinancialEquationsCode").value = IIf(TxtGroupCode.text = "", "", Trim(TxtGroupCode.text))
        rs("FinancialEquationsName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
        rs("FinancialEquationsDes").value = IIf(Me.txtFinancialEquationsDes.text = "", "", Trim(txtFinancialEquationsDes.text))
        rs("FinancialEquationsUp").value = ""
        rs("FinancialEquationsDown").value = ""
        rs("FinancialEquationsOpr").value = IIf(TxtFinancialEquationsOpr.text = "", "", Trim(TxtFinancialEquationsOpr.text))
        rs("GeneralValue").value = IIf(TxtGeneralValue.text = "", 0, Trim(TxtGeneralValue.text))
         
    End If
        
    rs.update
    Dim sql As String
    sql = "Delete  From  FinancialEquationsData where FinancialEquationsId= " & Me.XPTxtID.text
    Cn.Execute sql
        
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "FinancialEquationsData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    With Me.FgAccounts

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
         
                RsDev.AddNew
                RsDev("FinancialEquationsId").value = Me.XPTxtID.text
                RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                RsDev("Opr").value = .TextMatrix(i, .ColIndex("Operator"))
                RsDev("UpOrDown").value = 0
                RsDev.update
            End If
       
        Next i

    End With

    With Me.FgAccounts1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
         
                RsDev.AddNew
                RsDev("FinancialEquationsId").value = Me.XPTxtID.text
                RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                RsDev("Opr").value = .TextMatrix(i, .ColIndex("Operator"))
                RsDev("UpOrDown").value = 1
                RsDev.update
            End If
       
        Next i

    End With

    Cn.CommitTrans
    BeginTrans = False
 
    Me.retrive (XPTxtID.text)

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„Ì“«‰ÌÂ" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data was Saved , do you want to enter another data y/n" & Chr(13)
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
            End If

    End Select

    TxtModFlg.text = "R"
 
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Can't Save error in entered data " & Chr(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« " & Chr(13)
        Else
            Msg = "Can't save Data , Reasons: Data integrity " & Chr(13)
 
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        rs.CancelUpdate
        Exit Sub
    End If

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Else
                Msg = "Sorry...... Error During Saving Data " & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «· ⁄œÌ·«  " & Chr(13)
            Else
                Msg = "Sorry...... Error During Saving cahanges" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End Select

End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
       
            '        XPTxtID.text = CStr(new_id("Groups", "GroupID", "", True))
            FgAccounts.Rows = 3
            FgAccounts.FixedRows = 2
            FgAccounts1.Rows = 3
            FgAccounts1.FixedRows = 2
lblUp.Caption = ""
lbldown.Caption = ""

            XPTxtName.SetFocus

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            FgAccounts.Rows = FgAccounts.Rows + 1
            FgAccounts1.Rows = FgAccounts1.Rows + 1

        Case 2
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If
       
            Del_Trans
       
        Case 5
     
        Case 6
            Unload Me

        Case 7
 
        Case 8
            RemoveLine
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String

    'On Error GoTo ErrTrap
    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„⁄«œ·… —ﬁ„ " & Chr(13)
        Msg = Msg + (Me.XPTxtID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
     
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '                XPTxtCurrent.Caption = 0
                    '                XPTxtCount.Caption = 0
                Else
                    retrive
                    'GetBoxData
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function RemoveLine1()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Function
    Dim sql As String
    
    If FgAccounts1.Rows > 1 Then
        If FgAccounts1.Rows = 2 Then
            Me.FgAccounts1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.FgAccounts1.Rows > 1 Then
                If Me.FgAccounts1.Row <> Me.FgAccounts1.FixedRows - 1 Then
                    Me.FgAccounts1.RemoveItem (Me.FgAccounts1.Row)
                End If
            End If
        End If
    End If
            
    drawEquationText
    ' ReLineGrid

End Function

Function RemoveLine()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Function
    Dim sql As String
    
    If FgAccounts.Rows > 1 Then
        If FgAccounts.Rows = 2 Then
            Me.FgAccounts.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.FgAccounts.Rows > 1 Then
                If Me.FgAccounts.Row <> Me.FgAccounts.FixedRows - 1 Then
                    Me.FgAccounts.RemoveItem (Me.FgAccounts.Row)
                End If
            End If
        End If
    End If
            
    ' ReLineGrid
    drawEquationText

End Function

Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
      
            retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub FgAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = 46 Then
        RemoveLine
    End If

End Sub

Private Sub FgAccounts1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = 46 Then
        RemoveLine1
    End If

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim My_SQL As String
    My_SQL = "  select id,code from currency"

    fill_combo Dccurrency, My_SQL

    Dim GrdBack As ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
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
        Dcombos.GetAccountingCodesENg Me.DboParentAccount
    Else
        Dcombos.GetAccountingCodes Me.DboParentAccount
    End If
  
    Me.Height = 8605
    Me.Width = 16000
    Resize_Form Me
    Set GrdBack = New ClsBackGroundPic
  
    With Me.TrvAccounts
        .Appearance = ccFlat
        .Checkboxes = False
        .BorderStyle = ccNone
        .LineStyle = tvwRootLines
        .SingleSel = False
    End With

    LoadData

    With FgAccounts
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
     
        .MergeCol(.ColIndex("AccountName")) = True
        .MergeCol(.ColIndex("Operator")) = True
        .Cell(flexcpText, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = "    "
        .Cell(flexcpAlignment, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = flexAlignCenterCenter
    End With

    With FgAccounts1
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
     
        .MergeCol(.ColIndex("AccountName")) = True
        .MergeCol(.ColIndex("Operator")) = True
        .Cell(flexcpText, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = "    "
        .Cell(flexcpAlignment, 0, .ColIndex("Account_Name"), 0, .ColIndex("Operator")) = flexAlignCenterCenter
    End With

    Set rs = New ADODB.Recordset
    StrSQL = "select * From FinancialEquations"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    XPBtnMove_Click 2
 
    Me.TxtModFlg.text = "R"
 
    Me.TrvAccounts.Nodes("r").EnsureVisible
    Me.TrvAccounts.Nodes("r").Expanded = True
    Me.TrvAccounts.Nodes("r").Selected = True

    'If OPEN_NEW_SCREEN = True Then
    'Cmd_Click (0)
    'End If

End Sub

Private Sub LoadData()
    ModTree.LoadTreeAccountBalanceSheet Me.TrvAccounts, True
End Sub

Private Sub Image3_Click()
    Image7.Picture = Image3.Picture
    TxtFinancialEquationsOpr.text = "+"

End Sub

Private Sub Image4_Click()
    Image7.Picture = Image4.Picture
    TxtFinancialEquationsOpr.text = "-"

End Sub

Private Sub Image5_Click()
    Image7.Picture = Image5.Picture
    TxtFinancialEquationsOpr.text = "*"

End Sub

Private Sub Image6_Click()
    Image7.Picture = Image6.Picture

    TxtFinancialEquationsOpr.text = "/"
End Sub

Private Sub Opr_Click(Index As Integer)

    If AccountCode = "" Then Exit Sub

    Select Case Index

        Case 0
            AddNewRow AccountCode, AccountName, "+"

        Case 1
            AddNewRow AccountCode, AccountName, "-"

        Case 2
            AddNewRow AccountCode, AccountName, "*"

        Case 3
            AddNewRow AccountCode, AccountName, "/"

        Case 4
            AddNewRow AccountCode, AccountName, "="
 
        Case 9
            AddNewRow1 AccountCode, AccountName, "+"

        Case 5
            AddNewRow1 AccountCode, AccountName, "-"

        Case 7
            AddNewRow1 AccountCode, AccountName, "*"

        Case 6
            AddNewRow1 AccountCode, AccountName, "/"

        Case 8
            AddNewRow1 AccountCode, AccountName, "="

    End Select

End Sub

Private Sub AddNewRow(Optional Account_Code As String, _
                      Optional account_name As String, _
                      Optional Operator As String)
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
  
    'On Error Resume Next
    With Me.FgAccounts

        LngRow = .Rows - 1
        .TextMatrix(LngRow, .ColIndex("Account_Code")) = Account_Code
        .TextMatrix(LngRow, .ColIndex("account_name")) = account_name
        .TextMatrix(LngRow, .ColIndex("Operator")) = Operator
    
        .TextMatrix(LngRow, .ColIndex("account_serial")) = 1
  
        '.AutoSize 0, .Cols - 1, False
        .Rows = .Rows + 1
    End With
 
    drawEquationText
 
End Sub

Function drawEquationText()
    On Error Resume Next
    Dim equationtext As String
    Dim account_name As String
    Dim Opr As String

    With Me.FgAccounts
 
        equationtext = ""

        For i = 2 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
                account_name = .TextMatrix(i, .ColIndex("account_name"))
                Opr = .TextMatrix(i, .ColIndex("Operator"))
                equationtext = equationtext & account_name & Opr
            End If

        Next i
 
    End With
If equationtext <> "" Then
    lblUp.Caption = Mid(equationtext, 1, Len(equationtext) - 1)
Else
lblUp.Caption = ""
End If
    With Me.FgAccounts1
 
        equationtext = ""

        For i = 2 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
                account_name = .TextMatrix(i, .ColIndex("account_name"))
                Opr = .TextMatrix(i, .ColIndex("Operator"))
                equationtext = equationtext & account_name & Opr
            End If

        Next i
 
    End With
If equationtext <> "" Then
    Me.lbldown.Caption = Mid(equationtext, 1, Len(equationtext) - 1)
Else
   Me.lbldown.Caption = ""
End If
End Function

Private Sub AddNewRow1(Optional Account_Code As String, _
                       Optional account_name As String, _
                       Optional Operator As String)
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
  
    'On Error Resume Next
    With Me.FgAccounts1

        LngRow = .Rows - 1
        .TextMatrix(LngRow, .ColIndex("Account_Code")) = Account_Code
        .TextMatrix(LngRow, .ColIndex("account_name")) = account_name
        .TextMatrix(LngRow, .ColIndex("Operator")) = Operator
    
        .TextMatrix(LngRow, .ColIndex("account_serial")) = 1
  
        '.AutoSize 0, .Cols - 1, False
        .Rows = .Rows + 1
    End With
 
    drawEquationText
 
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Financial Analysis Equation Prepare"
    LblHeader.Caption = Me.Caption

    lbl(0).Caption = "Eq Code"
    lbl(2).Caption = "Eq Name"
    lbl(6).Caption = "Des"

    With FgAccounts
 
        .TextMatrix(1, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(1, .ColIndex("Operator")) = "Operator"
 
    End With

    With FgAccounts1
 
        .TextMatrix(1, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(1, .ColIndex("Operator")) = "Operator"
 
    End With

    lbl(4).Caption = "Result"
    lbl(100).Caption = "Value"
    Cmd(8).Caption = "Remove"
    Label7.Caption = "Equation"

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

Public Sub retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        '   XPTxtCurrent.Caption = 0
        '   XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "FinancialEquationsId=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("FinancialEquationsId").value), "", val(rs("FinancialEquationsId").value))
    Me.TxtGroupCode.text = IIf(IsNull(rs("FinancialEquationsCode").value), "", Trim(rs("FinancialEquationsCode").value))
    XPTxtName.text = IIf(IsNull(rs("FinancialEquationsName").value), "", Trim(rs("FinancialEquationsName").value))

    TxtGeneralValue.text = IIf(Not IsNumeric(rs("GeneralValue").value), "", val(rs("GeneralValue").value))
    TxtFinancialEquationsOpr.text = IIf(IsNull(rs("FinancialEquationsOpr").value), "", Trim(rs("FinancialEquationsOpr").value))
    txtFinancialEquationsDes.text = IIf(IsNull(rs("FinancialEquationsDes").value), "", Trim(rs("FinancialEquationsDes").value))

    Select Case TxtFinancialEquationsOpr.text

        Case ""
            Image7.Picture = Image8.Picture

        Case "+"
            Image7.Picture = Image3.Picture

        Case "-"
            Image7.Picture = Image4.Picture

        Case "*"
            Image7.Picture = Image5.Picture

        Case "/"
            Image7.Picture = Image6.Picture
    End Select
 
    StrSQL = " SELECT     dbo.ACCOUNTS.Account_Name, dbo.FinancialEquationsData.Opr, dbo.FinancialEquationsData.UpOrDown, dbo.FinancialEquationsData.Account_Code, "
    StrSQL = StrSQL & "                dbo.FinancialEquationsData.FinancialEquationsId"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.FinancialEquationsData ON dbo.ACCOUNTS.Account_Code = dbo.FinancialEquationsData.Account_Code"
    StrSQL = StrSQL & " where FinancialEquationsId = " & val(Me.XPTxtID.text)
    StrSQL = StrSQL & "  AND UpOrDown =0"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.FgAccounts
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
            
                .TextMatrix(i, .ColIndex("Operator")) = IIf(IsNull(RsDev("Opr").value), "", RsDev("Opr").value)
         
                '              .AutoSize 0, .Cols - 1, False
                RsDev.MoveNext
            Next i
 
        End With
Else
FgAccounts2.Rows = 2
    End If

    RsDev.Close

    StrSQL = " SELECT     dbo.ACCOUNTS.Account_Name, dbo.FinancialEquationsData.Opr, dbo.FinancialEquationsData.UpOrDown, dbo.FinancialEquationsData.Account_Code, "
    StrSQL = StrSQL & "                dbo.FinancialEquationsData.FinancialEquationsId"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.FinancialEquationsData ON dbo.ACCOUNTS.Account_Code = dbo.FinancialEquationsData.Account_Code"
    StrSQL = StrSQL & " where FinancialEquationsId = " & val(Me.XPTxtID.text)
    StrSQL = StrSQL & "  AND UpOrDown =1"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.FgAccounts1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
            
                .TextMatrix(i, .ColIndex("Operator")) = IIf(IsNull(RsDev("Opr").value), "", RsDev("Opr").value)
         
                '              .AutoSize 0, .Cols - 1, False
                RsDev.MoveNext
            Next i
 
        End With
Else
FgAccounts1.Rows = 2
    End If

    drawEquationText
    Exit Sub
ErrTrap:
End Sub

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)

    If Not Node Is Nothing Then

        If InStr(1, Node.key, "G", vbTextCompare) <> 0 Then
            StrTemp = Node.key
            StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
       
        Else
            StrTemp = Node.key
        End If
 
        AccountCode = StrTemp
        AccountName = Node.text
    End If
 
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
    ElseIf Me.TxtModFlg.text = "E" Then
 
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
 
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    retrive
    Exit Sub
ErrTrap:

End Sub
