VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BaklanceSheet 
   Caption         =   "«⁄œ«œ ‘ﬂ· «·„Ì“«‰Ì…"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   12795
   Icon            =   "BalanceSheet.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12795
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
      Width           =   12795
      _cx             =   22569
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
      _GridInfo       =   $"BalanceSheet.frx":038A
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
         Width           =   12765
         _cx             =   22516
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
            Height          =   7815
            Index           =   6
            Left            =   10515
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   30
            Width           =   2220
            _cx             =   3916
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
                     Picture         =   "BalanceSheet.frx":040D
                     Key             =   "Expanded_Node"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "BalanceSheet.frx":125F
                     Key             =   "Root"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "BalanceSheet.frx":15F9
                     Key             =   "Open_Node"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "BalanceSheet.frx":1993
                     Key             =   "Closed_Node"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "BalanceSheet.frx":1D2D
                     Key             =   "Item"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.TreeView TrvAccounts 
               Height          =   7755
               HelpContextID   =   380
               Left            =   0
               TabIndex        =   5
               Top             =   30
               Width           =   2220
               _ExtentX        =   3916
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
            Left            =   30
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   10455
            _cx             =   18441
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
            _GridInfo       =   $"BalanceSheet.frx":20C7
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Frame7 
               Height          =   765
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   30
               Width           =   10395
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
                  ButtonImage     =   "BalanceSheet.frx":2190
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
                  ButtonImage     =   "BalanceSheet.frx":252A
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
                  ButtonImage     =   "BalanceSheet.frx":28C4
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
                  ButtonImage     =   "BalanceSheet.frx":2C5E
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
                  Caption         =   "«⁄œ«œ ‘ﬂ· «·„Ì“«‰Ì…"
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
                  Width           =   10395
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   1
               Left            =   30
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   810
               Width           =   10395
               _cx             =   18336
               _cy             =   1349
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
               Begin VB.TextBox XPTxtID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   7920
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Text            =   "Text1"
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
                  Top             =   1200
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
               Begin VB.TextBox TxtAccount_NameE 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   960
                  Width           =   3330
               End
               Begin MSDataListLib.DataCombo DboParentAccount 
                  Height          =   315
                  Left            =   5115
                  TabIndex        =   31
                  Top             =   1380
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
                     ItemData        =   "BalanceSheet.frx":2FF8
                     Left            =   2520
                     List            =   "BalanceSheet.frx":3005
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   675
                     Width           =   1215
                  End
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     ItemData        =   "BalanceSheet.frx":3022
                     Left            =   0
                     List            =   "BalanceSheet.frx":3035
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
                     Picture         =   "BalanceSheet.frx":3061
                     Top             =   225
                     Width           =   240
                  End
                  Begin VB.Image Img 
                     Height          =   240
                     Index           =   0
                     Left            =   2940
                     Picture         =   "BalanceSheet.frx":33EB
                     Top             =   225
                     Width           =   240
                  End
               End
               Begin VB.TextBox XPTxtName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   300
                  Width           =   5730
               End
               Begin VB.TextBox TxtGroupCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   540
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
                  Top             =   990
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «·Õ”«» «·—∆Ì”Ì   "
                  Height          =   435
                  Index           =   6
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   1380
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Ì“«‰ÌÂ"
                  Height          =   285
                  Index           =   2
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   300
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·„Ì“«‰ÌÂ"
                  Height          =   345
                  Index           =   0
                  Left            =   8760
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
               Height          =   765
               Index           =   7
               Left            =   7395
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   3135
               Visible         =   0   'False
               Width           =   3030
               _cx             =   5345
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
                  ButtonImage     =   "BalanceSheet.frx":3775
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
                  ButtonImage     =   "BalanceSheet.frx":3B0F
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
                  ButtonImage     =   "BalanceSheet.frx":3EA9
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
               Top             =   7020
               Width           =   10395
               _cx             =   18336
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
                  Height          =   330
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   825
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   8
                  Left            =   240
                  TabIndex        =   79
                  Top             =   0
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   661
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
               Begin VB.Image Image1 
                  Height          =   240
                  Left            =   9705
                  Picture         =   "BalanceSheet.frx":4243
                  Top             =   30
                  Width           =   240
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Left            =   7650
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   30
                  Width           =   1920
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
                  Height          =   465
                  Index           =   5
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   300
                  Width           =   10335
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FgAccounts 
               Height          =   5415
               Left            =   30
               TabIndex        =   50
               Top             =   1590
               Width           =   10395
               _cx             =   18336
               _cy             =   9551
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
               Rows            =   2
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"BalanceSheet.frx":45CD
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
               Height          =   1530
               Left            =   4335
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   810
               Width           =   3045
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„·… «·Õ”«»"
               Height          =   765
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
         Top             =   7905
         Width           =   12765
         _cx             =   22516
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
            Left            =   11340
            TabIndex        =   7
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
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
            Left            =   9870
            TabIndex        =   8
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
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
            Left            =   8505
            TabIndex        =   9
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   7380
            TabIndex        =   10
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
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
            Left            =   5415
            TabIndex        =   11
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
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
            Left            =   4275
            TabIndex        =   12
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
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
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
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
            Left            =   2895
            TabIndex        =   14
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
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
            Left            =   1455
            TabIndex        =   15
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
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
Attribute VB_Name = "BaklanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset

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
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ì“«‰ÌÂ"
            Else
                Msg = "plz enter group name firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtName.SetFocus
            Exit Sub
        End If
   
        Select Case TxtModFlg.text

            Case "N"
                StrSQL = "select * From BalanceSheetView where BalanceSheetName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÊÃœ „Ì“«‰ÌÂ „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                
                    Else
                        Msg = "This group Name Already Exisi" & Chr(13)
            
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            
                XPTxtID.text = CStr(new_id("BalanceSheetView", "BalanceSheetId", "", True))
                Me.TxtGroupCode.text = XPTxtID.text

            Case "E"
                StrSQL = "select * From BalanceSheetView where BalanceSheetName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("BalanceSheetId").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÊÃœ „Ì“«‰ÌÂ „”Ã·… „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                        Else
                            Msg = "This group Name Already Exisi" & Chr(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If

        End Select
     
        Select Case TxtModFlg.text

            Case "N"
                Cn.BeginTrans
                BeginTrans = True
            
                rs.AddNew
                rs("BalanceSheetId").value = IIf(XPTxtID.text = "", "", val(XPTxtID.text))
            
            Case "E"
 
                Cn.BeginTrans
                BeginTrans = True
        End Select

        rs("BalanceSheetCode").value = IIf(TxtGroupCode.text = "", "", Trim(TxtGroupCode.text))
        rs("BalanceSheetName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
  
    End If
        
    rs.update
    Dim sql As String
    sql = "Delete  From  BalanceSheetViewAccounts where BalanceSheetId= " & Me.XPTxtID.text
    Cn.Execute sql
        
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "BalanceSheetViewAccounts", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    With Me.FgAccounts

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Account_Code")) <> "" Then
         
                RsDev.AddNew
                RsDev("BalanceSheetId").value = Me.XPTxtID.text
                RsDev("Account_Code").value = .TextMatrix(i, .ColIndex("Account_Code"))
                RsDev.update
            End If
       
        Next i

    End With

    Cn.CommitTrans
    BeginTrans = False
 
    Me.Retrive (XPTxtID.text)

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„Ì“«‰ÌÂ" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data was Saved , do you want to enter another data y/n" & Chr(13)
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
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

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« " & Chr(13)
        Else
            Msg = "Can't save Data , Reasons: Data integrity " & Chr(13)
 
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «· ⁄œÌ·«  " & Chr(13)
            Else
                Msg = "Sorry...... Error During Saving cahanges" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
 
            XPTxtName.SetFocus

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"

        Case 2
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If
       
        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            FrmGroupSearch.Show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

        Case 8
            RemoveLine
    End Select

    Exit Sub
ErrTrap:
End Sub

Function RemoveLine()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Function
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
    With FgAccounts
        '           Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
        '      Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
        '          Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
        '            Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
    End With

End Function

Function Undo()

End Function

Private Sub Form_Load()
    Dim Msg As String
    Dim My_SQL As String
    My_SQL = "  select id,code from currency"

    fill_combo DcCurrency, My_SQL

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

    Set rs = New ADODB.Recordset
    StrSQL = "select * From BalanceSheetView"
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

Private Sub TrvAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim AccountCode As String
    Dim AccountName As String

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

    AddNewRow AccountCode, AccountName
End Sub

Private Sub AddNewRow(Optional Account_Code As String, _
                      Optional account_name As String)
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
  
    On Error Resume Next

    With Me.FgAccounts

        LngRow = .Rows - 1
        .TextMatrix(LngRow, .ColIndex("Account_Code")) = Account_Code
        .TextMatrix(LngRow, .ColIndex("account_name")) = account_name
    
        .TextMatrix(LngRow, .ColIndex("account_serial")) = 1
  
        .AutoSize 0, .Cols - 1, False
        .Rows = .Rows + 1
    End With
 
End Sub

Private Sub ChangeLang()
    ChKBlock.Caption = "Block"
    Frame6.Caption = "Balance Type"
    DepitOrCredit(0).Caption = "Depit"
    DepitOrCredit(1).Caption = "Credit"
    Label4.Caption = "In Different Case"
    Differenttype(0).Caption = "Acess Deny"
    Differenttype(1).Caption = "Alarm Only"
    lbl(3).Caption = "Acc. Type"
    lbl(9).Caption = "Acc. Class."
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
    lbl(1).Caption = "Account#"
    lbl(0).Caption = "Acc. Code"
    lbl(6).Caption = "Parent Account"
    lbl(2).Caption = "Name A"
    lbl(7).Caption = "Name E"
    'lbl(3).Caption = "Derived Account From This Acc"
    lbl(4).Caption = "Note"
    lbl(5).Visible = False
    Check1.Caption = " have A Budget"
    Check2.Caption = "Cost Center"
    Check3.Caption = "Sum account"
    lbl(8).Caption = "Curr"

    Ele(5).Caption = "Account Type"

    With FgAccounts
        .TextMatrix(0, .ColIndex("Account_ID")) = "Account ID"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code"
        .TextMatrix(0, .ColIndex("Account_name")) = "Account Name"
        .TextMatrix(0, .ColIndex("OpenAccount")) = "Opening Balance"
        .TextMatrix(0, .ColIndex("AccountState")) = "Account State"
        .TextMatrix(0, .ColIndex("DateCreated")) = "DateCreated"
        .TextMatrix(0, .ColIndex("CurrentAccount")) = "Current Account"
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

Public Sub Retrive(Optional Lngid As Long = 0)

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
        rs.find "BalanceSheetId=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("BalanceSheetId").value), "", val(rs("BalanceSheetId").value))
    Me.TxtGroupCode.text = IIf(IsNull(rs("BalanceSheetCode").value), "", Trim(rs("BalanceSheetCode").value))
    XPTxtName.text = IIf(IsNull(rs("BalanceSheetName").value), "", Trim(rs("BalanceSheetName").value))
    StrSQL = "SELECT     dbo.BalanceSheetViewAccounts.Account_Code, dbo.ACCOUNTS.Account_Name"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & "  dbo.BalanceSheetViewAccounts ON dbo.ACCOUNTS.Account_Code = dbo.BalanceSheetViewAccounts.Account_Code"
    StrSQL = StrSQL & " where BalanceSheetId = " & val(Me.XPTxtID.text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.FgAccounts
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
            
                .AutoSize 0, .Cols - 1, False
                RsDev.MoveNext
            Next i
 
        End With

    End If

    Exit Sub
ErrTrap:
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

    Retrive
    Exit Sub
ErrTrap:

End Sub
