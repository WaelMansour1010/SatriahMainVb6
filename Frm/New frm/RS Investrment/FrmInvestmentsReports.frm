VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmInvestmentsReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   Icon            =   "FrmInvestmentsReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   5328
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   12192
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   8520
         TabIndex        =   4
         Top             =   120
         Width           =   3972
         Begin VB.Image Image1 
            Height          =   3672
            Left            =   0
            Picture         =   "FrmInvestmentsReports.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3912
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
            Height          =   1092
            Left            =   600
            TabIndex        =   5
            Top             =   3840
            Width           =   2892
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   3612
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8388
         _cx             =   14795
         _cy             =   6371
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
         Caption         =   "›Ê« Ì— «·„»Ì⁄« | ’›Ì… „”«Â„…| Ê“Ì⁄ «·«—»«Õ|«· ‰«“·|»Ì«‰«  «·„”«Â„« |»Ì«‰«  «·«—÷|„’—Ê›«  «· ÿÊÌ—|«· ﬁ”Ì„|«·„”«Â„Ì‰"
         Align           =   0
         CurrTab         =   8
         FirstTab        =   1
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3195
            Left            =   -10740
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   2532
               Left            =   720
               TabIndex        =   41
               Top             =   240
               Width           =   6372
               Begin VB.TextBox TxtDcbEmploSearch 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3528
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   720
                  Width           =   1068
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   528
                  Index           =   1
                  Left            =   240
                  TabIndex        =   43
                  Top             =   1560
                  Width           =   4284
                  _ExtentX        =   7567
                  _ExtentY        =   926
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
               Begin MSDataListLib.DataCombo DcbInvest 
                  Bindings        =   "FrmInvestmentsReports.frx":28E2
                  Height          =   288
                  Left            =   144
                  TabIndex        =   44
                  Top             =   360
                  Width           =   4452
                  _ExtentX        =   7832
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
               Begin MSDataListLib.DataCombo DcbEmployee 
                  Bindings        =   "FrmInvestmentsReports.frx":28F7
                  Height          =   288
                  Left            =   144
                  TabIndex        =   45
                  Top             =   720
                  Width           =   3252
                  _ExtentX        =   5715
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
                  Caption         =   "«·ﬁ«∆„ »«· ’›Ì…"
                  Height          =   312
                  Index           =   22
                  Left            =   4380
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   720
                  Width           =   1956
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„”«Â„…"
                  Height          =   312
                  Index           =   16
                  Left            =   4380
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   360
                  Width           =   1956
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3195
            Index           =   2
            Left            =   -11040
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
               Caption         =   "„‰ «·› —Â"
               Height          =   735
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   1680
               Width           =   5676
               Begin MSComCtl2.DTPicker DtpDateFrom 
                  Height          =   330
                  Left            =   2640
                  TabIndex        =   19
                  Top             =   270
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93847555
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DtpDateTo 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   93847555
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   3
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
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
                  TabIndex        =   21
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   1692
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   0
               Width           =   5676
               Begin MSDataListLib.DataCombo DcbSeller 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   13
                  Top             =   360
                  Width           =   4176
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcLand 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   14
                  Top             =   720
                  Width           =   4176
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcShare 
                  Height          =   288
                  Left            =   120
                  TabIndex        =   23
                  Top             =   1080
                  Width           =   4164
                  _ExtentX        =   7355
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«—÷"
                  Height          =   288
                  Index           =   0
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   720
                  Width           =   1092
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„”«Â„…"
                  Height          =   288
                  Index           =   38
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   1080
                  Width           =   1092
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»«∆⁄"
                  Height          =   288
                  Index           =   37
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   360
                  Width           =   1092
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   492
               Index           =   0
               Left            =   948
               TabIndex        =   24
               Top             =   2520
               Width           =   5700
               _ExtentX        =   10054
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3195
            Left            =   -10440
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Height          =   2892
               Left            =   1200
               TabIndex        =   48
               Top             =   240
               Width           =   5892
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   3
                  Left            =   240
                  TabIndex        =   49
                  Top             =   1680
                  Width           =   3420
                  _ExtentX        =   6033
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
               Begin MSDataListLib.DataCombo dcShare3 
                  Bindings        =   "FrmInvestmentsReports.frx":290C
                  Height          =   288
                  Left            =   120
                  TabIndex        =   50
                  Top             =   480
                  Width           =   4008
                  _ExtentX        =   7064
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
               Begin MSDataListLib.DataCombo dcEmpShare 
                  Bindings        =   "FrmInvestmentsReports.frx":2921
                  Height          =   288
                  Left            =   120
                  TabIndex        =   51
                  Top             =   840
                  Width           =   4008
                  _ExtentX        =   7064
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
                  Caption         =   "«·„”«Â„…"
                  Height          =   288
                  Index           =   2
                  Left            =   3876
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   480
                  Width           =   1776
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”«Â„"
                  Height          =   288
                  Index           =   5
                  Left            =   3876
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   840
                  Width           =   1776
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3195
            Left            =   -10140
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Height          =   3012
               Left            =   1800
               TabIndex        =   54
               Top             =   120
               Width           =   5652
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   4
                  Left            =   240
                  TabIndex        =   55
                  Top             =   1800
                  Width           =   3900
                  _ExtentX        =   6879
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
               Begin MSDataListLib.DataCombo dcShare4 
                  Bindings        =   "FrmInvestmentsReports.frx":2936
                  Height          =   288
                  Left            =   120
                  TabIndex        =   56
                  Top             =   600
                  Width           =   4008
                  _ExtentX        =   7064
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
               Begin MSDataListLib.DataCombo dcEmpShare4 
                  Bindings        =   "FrmInvestmentsReports.frx":294B
                  Height          =   288
                  Left            =   120
                  TabIndex        =   57
                  Top             =   960
                  Width           =   4008
                  _ExtentX        =   7064
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
                  Caption         =   "„”«Â„"
                  Height          =   288
                  Index           =   6
                  Left            =   3876
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   960
                  Width           =   2256
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„”«Â„…"
                  Height          =   288
                  Index           =   7
                  Left            =   3876
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   600
                  Width           =   2256
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3195
            Left            =   -9840
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Height          =   2892
               Left            =   1680
               TabIndex        =   60
               Top             =   240
               Width           =   5892
               Begin VB.ComboBox DcbType 
                  Height          =   288
                  ItemData        =   "FrmInvestmentsReports.frx":2960
                  Left            =   120
                  List            =   "FrmInvestmentsReports.frx":2962
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1440
                  Width           =   4008
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   5
                  Left            =   120
                  TabIndex        =   62
                  Top             =   2040
                  Width           =   4020
                  _ExtentX        =   7091
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
               Begin MSDataListLib.DataCombo dcShare5 
                  Bindings        =   "FrmInvestmentsReports.frx":2964
                  Height          =   288
                  Left            =   120
                  TabIndex        =   63
                  Top             =   480
                  Width           =   4008
                  _ExtentX        =   7064
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
               Begin MSDataListLib.DataCombo DcbGroupInvs 
                  Bindings        =   "FrmInvestmentsReports.frx":2979
                  Height          =   288
                  Left            =   120
                  TabIndex        =   64
                  Top             =   960
                  Width           =   4008
                  _ExtentX        =   7064
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
                  Caption         =   "«·„”«Â„…"
                  Height          =   288
                  Index           =   8
                  Left            =   3876
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   480
                  Width           =   2256
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ã„Ê⁄… «·„”«Â„…"
                  Height          =   288
                  Index           =   10
                  Left            =   4128
                  TabIndex        =   66
                  Top             =   960
                  Width           =   1332
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·⁄‰’—"
                  Height          =   288
                  Index           =   11
                  Left            =   4224
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   1476
                  Width           =   1344
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   3195
            Left            =   -9540
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2E9E9&
               Height          =   3012
               Left            =   1080
               TabIndex        =   68
               Top             =   120
               Width           =   6492
               Begin VB.TextBox TxtFullCode 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   240
                  TabIndex        =   70
                  Top             =   1320
                  Width           =   3960
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   240
                  TabIndex        =   69
                  Top             =   1800
                  Width           =   3960
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   6
                  Left            =   240
                  TabIndex        =   71
                  Top             =   2280
                  Width           =   3960
                  _ExtentX        =   6985
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
               Begin MSDataListLib.DataCombo dcsupplier 
                  Bindings        =   "FrmInvestmentsReports.frx":298E
                  Height          =   288
                  Left            =   240
                  TabIndex        =   72
                  Top             =   360
                  Width           =   3960
                  _ExtentX        =   6985
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
               Begin MSDataListLib.DataCombo dcBranch6 
                  Bindings        =   "FrmInvestmentsReports.frx":29A3
                  Height          =   288
                  Left            =   240
                  TabIndex        =   73
                  Top             =   840
                  Width           =   3960
                  _ExtentX        =   6985
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
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂÊœ"
                  Height          =   288
                  Index           =   9
                  Left            =   4296
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   1356
                  Width           =   1524
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   288
                  Index           =   12
                  Left            =   4296
                  TabIndex        =   76
                  Top             =   840
                  Width           =   1524
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„«·ﬂ"
                  Height          =   288
                  Index           =   13
                  Left            =   3288
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   360
                  Width           =   2532
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„"
                  Height          =   288
                  Index           =   14
                  Left            =   4296
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1836
                  Width           =   1524
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   3195
            Left            =   -9240
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame10 
               BackColor       =   &H00E2E9E9&
               Height          =   2772
               Left            =   1440
               TabIndex        =   78
               Top             =   120
               Width           =   6372
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   7
                  Left            =   840
                  TabIndex        =   79
                  Top             =   1800
                  Width           =   3420
                  _ExtentX        =   6033
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
               Begin MSDataListLib.DataCombo DcbInvise 
                  Height          =   288
                  Left            =   720
                  TabIndex        =   80
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   360
                  Width           =   3444
                  _ExtentX        =   6085
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbLand 
                  Height          =   288
                  Left            =   720
                  TabIndex        =   81
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   720
                  Width           =   3444
                  _ExtentX        =   6085
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «·„”«Â„…"
                  Height          =   288
                  Index           =   7
                  Left            =   4464
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«—÷ „„·Êﬂ…"
                  Height          =   288
                  Index           =   6
                  Left            =   4464
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   720
                  Width           =   1440
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   3195
            Left            =   -8940
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   3012
               Left            =   360
               TabIndex        =   32
               Top             =   120
               Width           =   6132
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   8
                  Left            =   480
                  TabIndex        =   33
                  Top             =   2280
                  Width           =   4056
                  _ExtentX        =   7144
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
               Begin MSDataListLib.DataCombo dcInvestment8 
                  Height          =   288
                  Left            =   600
                  TabIndex        =   34
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   240
                  Width           =   3072
                  _ExtentX        =   5424
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbLand8 
                  Height          =   288
                  Left            =   600
                  TabIndex        =   35
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   600
                  Width           =   3072
                  _ExtentX        =   5424
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   252
                  Index           =   0
                  Left            =   2664
                  TabIndex        =   36
                  Top             =   1080
                  Width           =   948
                  _Version        =   786432
                  _ExtentX        =   1672
                  _ExtentY        =   444
                  _StockProps     =   79
                  Caption         =   "„”«Â„…"
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   252
                  Index           =   1
                  Left            =   768
                  TabIndex        =   37
                  Top             =   1080
                  Width           =   1548
                  _Version        =   786432
                  _ExtentX        =   2730
                  _ExtentY        =   444
                  _StockProps     =   79
                  Caption         =   "«—÷ „„·Êﬂ…"
                  ForeColor       =   8388608
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«—÷ „„·Êﬂ…"
                  Height          =   288
                  Index           =   0
                  Left            =   4488
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   600
                  Width           =   1368
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «·„”«Â„…"
                  Height          =   288
                  Index           =   1
                  Left            =   4488
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   240
                  Width           =   1368
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ „⁄Ì‰"
                  Height          =   288
                  Index           =   2
                  Left            =   4488
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   1368
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   3195
            Left            =   45
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   45
            Width           =   8295
            _cx             =   14631
            _cy             =   5636
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
            Begin VB.Frame Frame11 
               BackColor       =   &H00E2E9E9&
               Height          =   3012
               Left            =   1200
               TabIndex        =   85
               Top             =   0
               Width           =   6492
               Begin VB.ComboBox dcInvestmentType9 
                  Height          =   288
                  ItemData        =   "FrmInvestmentsReports.frx":29B8
                  Left            =   240
                  List            =   "FrmInvestmentsReports.frx":29C2
                  TabIndex        =   94
                  Top             =   360
                  Width           =   3972
               End
               Begin VB.TextBox txtName9 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   240
                  TabIndex        =   87
                  Top             =   1800
                  Width           =   3960
               End
               Begin VB.TextBox txtCode9 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   240
                  TabIndex        =   86
                  Top             =   1320
                  Width           =   3960
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   492
                  Index           =   9
                  Left            =   240
                  TabIndex        =   88
                  Top             =   2280
                  Width           =   3960
                  _ExtentX        =   6985
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
               Begin MSDataListLib.DataCombo DcBranch9 
                  Bindings        =   "FrmInvestmentsReports.frx":29D3
                  Height          =   288
                  Left            =   240
                  TabIndex        =   89
                  Top             =   840
                  Width           =   3960
                  _ExtentX        =   6985
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
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„"
                  Height          =   288
                  Index           =   19
                  Left            =   4296
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   1836
                  Width           =   1524
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·„”«Â„"
                  Height          =   288
                  Index           =   18
                  Left            =   3288
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   360
                  Width           =   2532
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   288
                  Index           =   17
                  Left            =   4296
                  TabIndex        =   91
                  Top             =   840
                  Width           =   1524
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂÊœ"
                  Height          =   288
                  Index           =   15
                  Left            =   4296
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   1356
                  Width           =   1524
               End
            End
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "‘«‘…  ﬁ«—Ì— «·«” À„«— "
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
         Height          =   1020
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   4080
         Width           =   5772
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1092
         Left            =   120
         Top             =   4080
         Width           =   5772
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   492
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   6120
      Width           =   1212
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
      Caption         =   "‘«‘…  ﬁ«—Ì— «·«” À„«—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12264
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
Attribute VB_Name = "FrmInvestmentsReports"
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
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

  
End Sub
Private Sub btnClear_Click()
clear_all Me
'DcbTypeMain.Enabled = False
'TxtSearchCode.Enabled = False
'DcbEmp.Enabled = False
'DcbDept.Enabled = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub

Private Sub Chorder_Click()
 End Sub

Private Sub CHReq_Click()
 End Sub

Private Sub Cmd_Click(Index As Integer)

        Select Case Index
        
        Case 0
                print_Invoice
        Case 1
                print_Liquidation
        Case 2
                Unload Me
        Case 3
               print_ProfitDist
               
       Case 4
                print_BuyBill
        Case 5
                Print_Investment
       Case 6
                Print_BuyLandReal
        Case 7
                Print_Expenses
        Case 8
                Print_LandDivid
        Case 9
                Print_InvestmentCustomers
        End Select

End Sub




Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub

Private Sub DcbEmp_Click(Area As Integer)
     
End Sub

Private Sub dcEmp_Change()
 End Sub

Private Sub DcbEmployee_Change()
If val(DcbEmployee.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmployee.BoundText, EmpCode
    TxtDcbEmploSearch.Text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    
End Sub

Private Sub LoadCombos()
Dim Dcombos As ClsDataCombos
Dim str As String
Set Dcombos = New ClsDataCombos
    
    
str = " select ID , Name  from TblBuyLanReEst "
fill_combo dcLand, str

'str = " select ID , Name  from Tblinvestment "
str = " SElect id ,Name from Tblinvestment where id  in (select InvesID from TblInvesOpenSales )   "
fill_combo dcShare, str

Dcombos.GetEmployees DcbSeller


'//////////////////

Dcombos.GetEmployees Me.DcbEmployee
Dcombos.GetInvestmentActive Me.DcbInvest
    
    '/////////////////////////
str = " select  cusID , cusName  from TblCustemers where type = 20  "
fill_combo dcEmpShare, str

Dcombos.GetInvestmentActive dcShare3


'////////////////////////////////
str = " select  cusID , cusName  from TblCustemers where type = 20  "
fill_combo dcEmpShare4, str

Dcombos.GetInvestmentActive dcShare4


'////////////////////////////////
Dcombos.GetInvestmentGroup Me.DcbGroupInvs
  If SystemOptions.UserInterface = ArabicInterface Then
     With DcbType
       .Clear
     .AddItem "«—«÷Ì"
       .AddItem "⁄ﬁ«—"
       
    End With
 Else
    With DcbType
      .Clear
     .AddItem "Land"
      .AddItem "Estate"
       
   End With
End If


fill_combo dcShare5, " SELECT ID,Name From dbo.Tblinvestment "

'/////////////////////////////////////////////
   Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
   Dcombos.GetBranches dcBranch6
   
'////////////////////
 Dcombos.GetInvestmentActive Me.DcbInvise
 Dcombos.GetBuyLandRealEstate DcbLand
 
 
'////////////////////////////////////////
Dcombos.GetInvestmentActive dcInvestment8
Dcombos.GetBuyLandRealEstate DcbLand8
 '   Dcombos.GetTblSpreadingInvestment DcbDivMain, 1
 
  '//////////////////////////////////////
  Dcombos.GetBranches DcBranch9
  
  
  
End Sub


Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim str As String
    
    LoadCombos
    
    DtpDateFrom.value = ""
    DtpDateTo.value = ""

    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If
    
    DtpDateFrom.value = Now
    DtpDateTo.value = Now
    
    DtpDateFrom.value = Null
    DtpDateTo.value = Null

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Function Selection_Query() As String
 
End Function


Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 
    StrSQL = Selection_Query

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        rs.MoveFirst
        print_report StrSQL
    End If


End Sub

Private Sub oneEmp_Code_Change()

 
End Sub

Private Sub OneEmployee_Click(Area As Integer)
 End Sub



Function print_Invoice(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""


 MySQL = MySQL & "   SELECT dbo.TblSaleBilllInvestment.BranchID, dbo.TblSaleBilllInvestment.SellerID, dbo.TblBranchesData.branch_name, dbo.TblBuyLanReEst.Name AS LandName, Client.CusID,"
 MySQL = MySQL & "   Client.CusName, Client.Fullcode, dbo.TblSaleBilllInvestment.ID, dbo.TblSaleBilllInvestment.RecordDate, dbo.TblSaleBilllInvestment.SellerType,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.UserID, dbo.TblSaleBilllInvestment.InvesID, dbo.TblSaleBilllInvestment.LandID, dbo.TblSaleBilllInvestment.commission,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.DesLocation, dbo.TblSaleBilllInvestment.Remarks, dbo.TblSaleBilllInvestment.PropertyDeed, dbo.TblSaleBilllInvestment.NorthlengthStr,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.SouthlengthStr, dbo.TblSaleBilllInvestment.EastlengthStr, dbo.TblSaleBilllInvestment.WestlengthStr, dbo.TblSaleBilllInvestment.Northlength,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.Southlength, dbo.TblSaleBilllInvestment.Eastlength, dbo.TblSaleBilllInvestment.Westlength, dbo.TblSaleBilllInvestment.Cus_Tpe,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.Cus_ID, dbo.TblSaleBilllInvestment.Payment, dbo.TblSaleBilllInvestment.RecordNo, dbo.TblSaleBilllInvestment.CusID AS Expr1,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.PaymentNo, dbo.TblSaleBilllInvestment.Period, dbo.TblSaleBilllInvestment.PeriodType, dbo.TblSaleBilllInvestment.RemarkPay,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.FristDate, dbo.TblSaleBilllInvestment.Typepartial, dbo.TblSaleBilllInvestment.TotaPtofit, dbo.TblSaleBilllInvestment.FlgReturn,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.TypeCom, dbo.TblSaleBilllInvestment.BillNo, dbo.TblSaleBilllInvestment.NetComm, dbo.TblSaleBilllInvestment.TypeRetSal,"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.TypDiv, dbo.TblSaleBilllInvestment.NoteSerial, dbo.TblSaleBilllInvestment.NoteID, dbo.TblSaleBilllInvestment.CusComm,"
 MySQL = MySQL & "   dbo.TblEmployee.Emp_Name AS SellerName"
 MySQL = MySQL & "   FROM     dbo.TblEmployee RIGHT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment ON dbo.TblEmployee.Emp_ID = dbo.TblSaleBilllInvestment.SellerID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.Tblinvestment RIGHT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestmentDet ON dbo.Tblinvestment.ID = dbo.TblSaleBilllInvestmentDet.InvesID ON"
 MySQL = MySQL & "   dbo.TblSaleBilllInvestment.ID = dbo.TblSaleBilllInvestmentDet.SBINVID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblCustemers AS Client ON dbo.TblSaleBilllInvestment.Cus_ID = Client.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblBuyLanReEst ON dbo.TblSaleBilllInvestment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblSaleBilllInvestment.BranchID = dbo.TblBranchesData.branch_id"

 
 MySQL = MySQL & "    where 1 = 1"
 
 
If dcShare.BoundText <> "" Then
        MySQL = MySQL & " and dbo.Tblinvestment.ID  =  " & val(dcShare.BoundText)
End If
 
If DcbSeller.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblSaleBilllInvestment.SellerID  =  " & val(DcbSeller.BoundText)
End If

If dcLand.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblSaleBilllInvestment.LandID  =  " & val(dcLand.BoundText)
End If


If Not IsNull(DtpDateFrom.value) Then
         MySQL = MySQL & "  and TblSaleBilllInvestment.RecordDate >= " & SQLDate(DtpDateFrom.value, True) & ""
End If

If Not IsNull(DtpDateTo.value) Then
         MySQL = MySQL & "  and TblSaleBilllInvestment.RecordDate <= " & SQLDate(DtpDateTo.value, True) & ""
End If

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_SalesBillinvestment.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_SalesBillinvestment.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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


 
Function print_Liquidation(Optional NoteSerial As String)
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""


MySQL = MySQL & "     SELECT dbo.TblInvestliquidation.EmpID, dbo.TblInvestliquidation.InvesID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.Tblinvestment.Name AS InvestmentName,"
MySQL = MySQL & "     dbo.TblInvestliquidation.RecordDate, dbo.TblInvestliquidation.BranchID, dbo.TblBranchesData.branch_name, dbo.TblInvestliquidation.LiqDate, dbo.TblInvestliquidation.ID,"
MySQL = MySQL & "     dbo.TblInvestliquidation.Remarks, dbo.TblInvestliquidation.TotalSalLand, dbo.TblInvestliquidation.TotalExpens, dbo.TblInvestliquidation.TotalCost,"
MySQL = MySQL & "     dbo.TblInvestliquidation.TotalDiff, dbo.TblInvestliquidation.TotalShare, dbo.TblInvestliquidation.WritSalLand, dbo.TblInvestliquidation.WritExpens,"
MySQL = MySQL & "     dbo.TblInvestliquidation.WritCost , dbo.TblInvestliquidation.WritShare, dbo.TblInvestliquidation.WritDiff"
MySQL = MySQL & "     FROM     dbo.TblInvestliquidation INNER JOIN"
MySQL = MySQL & "     dbo.TblEmployee ON dbo.TblInvestliquidation.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
MySQL = MySQL & "     dbo.Tblinvestment ON dbo.TblInvestliquidation.InvesID = dbo.Tblinvestment.ID INNER JOIN"
MySQL = MySQL & "     dbo.TblBranchesData ON dbo.TblInvestliquidation.BranchID = dbo.TblBranchesData.branch_id"
 
 
If DcbInvest.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblInvestliquidation.InvesID  =  " & val(DcbInvest.BoundText)
End If
 
If DcbEmployee.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblInvestliquidation.EmpID  =  " & val(DcbEmployee.BoundText)
End If

If dcLand.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblSaleBilllInvestment.LandID  =  " & val(dcLand.BoundText)
End If


If Not IsNull(DtpDateFrom.value) Then
         MySQL = MySQL & "  and TblInvestliquidation.RecordDate >= " & SQLDate(DtpDateFrom.value, True) & ""
End If

If Not IsNull(DtpDateTo.value) Then
         MySQL = MySQL & "  and TblInvestliquidation.RecordDate <= " & SQLDate(DtpDateTo.value, True) & ""
End If

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Liquidation.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Liquidation.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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

Private Sub TxtDcbEmploSearch_Change()

Dim EmpID As Integer
GetEmployeeIDFromCode TxtDcbEmploSearch.Text, EmpID
DcbEmployee.BoundText = EmpID

End Sub

Function print_ProfitDist(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""


MySQL = MySQL & "  SELECT dbo.TblInvestProfitDistri.BranchID, dbo.TblInvestProfitDistri.RecordDate, dbo.TblInvestProfitDistri.ID, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "  dbo.Tblinvestment.Name AS InvestmentName, dbo.TblInvestProfitDistri.InvesID, dbo.TblInvestProfitDistri.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
MySQL = MySQL & "  dbo.TblInvestProfitDistri.UserID, dbo.TblInvestProfitDistri.Remarks, dbo.TblInvestProfitDistri.Comm, dbo.TblInvestProfitDistri.InvestValue, dbo.TblInvestProfitDistri.SalValue,"
MySQL = MySQL & "  dbo.TblInvestProfitDistri.PorfetValue, dbo.TblInvestProfitDistri.ShareID, dbo.TblInvestProfitDistri.TypeShere, dbo.TblInvestProfitDistri.SharNo,"
MySQL = MySQL & "  dbo.TblInvestProfitDistri.TotalShare , dbo.TblInvestProfitDistri.NetComm, dbo.TblInvestProfitDistri.TypeCom, dbo.TblInvestProfitDistri.NeProfit"
MySQL = MySQL & "  FROM     dbo.TblInvestProfitDistri INNER JOIN"
MySQL = MySQL & "  dbo.TblInvestProfitDistriDet ON dbo.TblInvestProfitDistri.ID = dbo.TblInvestProfitDistriDet.InvProID INNER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblInvestProfitDistri.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
MySQL = MySQL & "  dbo.Tblinvestment ON dbo.TblInvestProfitDistri.InvesID = dbo.Tblinvestment.ID INNER JOIN"
MySQL = MySQL & "  dbo.TblEmployee ON dbo.TblInvestProfitDistri.EmpID = dbo.TblEmployee.Emp_ID"

 MySQL = MySQL & "    where 1 = 1"
 
 
If dcShare3.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblInvestProfitDistri.InvesID  =  " & val(dcShare3.BoundText)
End If
 
If dcEmpShare.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblInvestProfitDistriDet.ShareID  =  " & val(dcEmpShare.BoundText)
End If


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ProfitDist.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ProfitDist.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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


Function print_BuyBill(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""
    
 MySQL = MySQL & "       SELECT dbo.TblBuyBilllInvestment.BranchID, dbo.TblBuyBilllInvestment.ID, dbo.TblBuyBilllInvestment.RecordDate, dbo.TblBuyBilllInvestment.Payment,"
MySQL = MySQL & "        dbo.TblBuyBilllInvestment.RecordNo, dbo.TblBuyBilllInvestment.Cus_Type, dbo.TblBuyBilllInvestment.CusID, dbo.TblBuyBilllInvestment.InvesID,"
MySQL = MySQL & "        dbo.TblBuyBilllInvestment.Cus_ID, dbo.TblBuyBilllInvestment.SharNo, dbo.TblBuyBilllInvestment.SharValue, dbo.TblBuyBilllInvestment.Remarks,"
MySQL = MySQL & "        dbo.TblBuyBilllInvestment.NetShare, dbo.TblBuyBilllInvestment.ShrNetValue, dbo.TblBuyBilllInvestment.Comm1, dbo.TblBuyBilllInvestment.Comm2,"
MySQL = MySQL & "        dbo.TblBuyBilllInvestment.ValueCom1, dbo.TblBuyBilllInvestment.ValueCom2, dbo.TblBuyBilllInvestment.TypeCom1, dbo.TblBuyBilllInvestment.TypeCom2,"
MySQL = MySQL & "        dbo.TblBuyBilllInvestment.SellerID, dbo.TblCustemers.CusName AS SellerName, Client.CusName AS BuyerName, dbo.TblBranchesData.branch_name"
MySQL = MySQL & "        FROM     dbo.TblBuyBilllInvestment INNER JOIN"
MySQL = MySQL & "        dbo.TblBranchesData ON dbo.TblBuyBilllInvestment.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
MySQL = MySQL & "        dbo.TblCustemers ON dbo.TblBuyBilllInvestment.SellerID = dbo.TblCustemers.CusID INNER JOIN"
MySQL = MySQL & "        dbo.TblCustemers AS Client ON dbo.TblBuyBilllInvestment.Cus_ID = Client.CusID INNER JOIN"
MySQL = MySQL & "        dbo.TblBuyBilllInvestmentDet ON dbo.TblBuyBilllInvestment.ID = dbo.TblBuyBilllInvestmentDet.BuyBilInvsID"

 MySQL = MySQL & "    where 1 = 1"
 
 
If dcShare4.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblBuyBilllInvestmentDet.InvesID  =  " & val(dcShare4.BoundText)
End If
 
If dcEmpShare4.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblBuyBilllInvestment.SellerID  =  " & val(dcEmpShare4.BoundText)
End If


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BuyBillInvestment.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BuyBillInvestment.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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

Function Print_Investment(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""
    
MySQL = MySQL & "  SELECT dbo.TblSharesGroup.Name AS GroupName, dbo.TblSharesGroup.Code AS GroupCode, dbo.Tblinvestment.ID, dbo.Tblinvestment.TypwInvse, dbo.Tblinvestment.GroupInvs,"
MySQL = MySQL & "  dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.Tblinvestment.InvsValue, dbo.Tblinvestment.DevlpValue, dbo.Tblinvestment.TotalInDe,"
MySQL = MySQL & "  dbo.Tblinvestment.AllInvsValue, dbo.Tblinvestment.warrantValue, dbo.Tblinvestment.Remark, dbo.Tblinvestment.BranchID, dbo.Tblinvestment.UserID,"
MySQL = MySQL & "  dbo.Tblinvestment.RecorDate, dbo.Tblinvestment.EmpID, dbo.Tblinvestment.StatusIPO, dbo.Tblinvestment.Typ, dbo.Tblinvestment.AccounCode,"
MySQL = MySQL & "  dbo.Tblinvestment.ParentAccount, dbo.Tblinvestment.BankID, dbo.Tblinvestment.BanckName, dbo.Tblinvestment.Account_Code, dbo.Tblinvestment.Account_Code1,"
MySQL = MySQL & "  dbo.Tblinvestment.Account_Code2, dbo.Tblinvestment.Account_Code3, dbo.Tblinvestment.Account_Code4, dbo.Tblinvestment.ParetnAccount,"
MySQL = MySQL & "  dbo.Tblinvestment.ParetnAccount1, dbo.Tblinvestment.RootAccount, dbo.Tblinvestment.ParentAccountSub, dbo.Tblinvestment.ParentAccount1,"
MySQL = MySQL & "  dbo.Tblinvestment.RootAccount1, dbo.Tblinvestment.Account_Code5, dbo.Tblinvestment.Account_Code6, dbo.Tblinvestment.Account_Code7, dbo.Tblinvestment.FlagActive,"
MySQL = MySQL & "  dbo.Tblinvestment.FlagExpenses, dbo.Tblinvestment.CostMeterExp, dbo.Tblinvestment.InvesValueExp, dbo.Tblinvestment.ExpenseValueExp,"
MySQL = MySQL & "  dbo.TblBranchesData.branch_name , dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
MySQL = MySQL & "  FROM     dbo.Tblinvestment LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblSharesGroup ON dbo.Tblinvestment.GroupInvs = dbo.TblSharesGroup.ID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.Tblinvestment.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblCustemers ON dbo.Tblinvestment.EmpID = dbo.TblCustemers.CusID"

 MySQL = MySQL & "    where 1 = 1"
 
 
If dcShare5.BoundText <> "" Then
        MySQL = MySQL & " and dbo.Tblinvestment.ID  =  " & val(dcShare5.BoundText)
End If
 
If DcbGroupInvs.BoundText <> "" Then
        MySQL = MySQL & " and dbo.Tblinvestment.GroupInvs  =  " & val(DcbGroupInvs.BoundText)
End If

If DcbType.ListIndex <> -1 Then
        MySQL = MySQL & " and dbo.Tblinvestment.GroupInvs  =  " & val(DcbType.ListIndex)
End If


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Investment.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Investment.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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



Function Print_BuyLandReal(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""

MySQL = MySQL & "   SELECT dbo.TblBuyLanReEst.ID, dbo.TblBuyLanReEst.RecordDate, dbo.TblBuyLanReEst.Name, dbo.TblBuyLanReEst.NameE, dbo.TblBuyLanReEst.FullCode,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.BranchID, dbo.TblBuyLanReEst.UserID, dbo.TblBuyLanReEst.No_planned, dbo.TblBuyLanReEst.Area, dbo.TblBuyLanReEst.MeterPrice,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.Total, dbo.TblBuyLanReEst.TitledeedNo, dbo.TblBuyLanReEst.OwnerID, dbo.TblBuyLanReEst.PayType, dbo.TblBuyLanReEst.InstalNo,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.FristDate, dbo.TblBuyLanReEst.Period, dbo.TblBuyLanReEst.PeriodType, dbo.TblBuyLanReEst.Remarks, dbo.TblBuyLanReEst.CountryID,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.CityID, dbo.TblBuyLanReEst.HyID, dbo.TblBuyLanReEst.SchemeID, dbo.TblBuyLanReEst.DesLocation, dbo.TblBuyLanReEst.Street,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.DatePropertyDeed, dbo.TblBuyLanReEst.Unit, dbo.TblBuyLanReEst.Block, dbo.TblBuyLanReEst.PlateNo, dbo.TblBuyLanReEst.StreetNo,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.NorthlengthStr, dbo.TblBuyLanReEst.SouthlengthStr, dbo.TblBuyLanReEst.EastlengthStr, dbo.TblBuyLanReEst.WestlengthStr,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.Northlength, dbo.TblBuyLanReEst.Southlength, dbo.TblBuyLanReEst.Eastlength, dbo.TblBuyLanReEst.Westlength, dbo.TblBuyLanReEst.Typepartial,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.SchemName, dbo.TblBuyLanReEst.ActivID, dbo.TblBuyLanReEst.NewLand, dbo.TblBuyLanReEst.FlagActive, dbo.TblBuyLanReEst.Googlemap,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.FlgReturn, dbo.TblBuyLanReEst.BuySal, dbo.TblBuyLanReEst.BillNo, dbo.TblBuyLanReEst.Debt_Credit, dbo.TblBuyLanReEst.OpenBalance,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.OpenDate, dbo.TblBuyLanReEst.Account_Code, dbo.TblBuyLanReEst.OpenBalanceDate, dbo.TblBuyLanReEst.OpenBalanceType,"
MySQL = MySQL & "   dbo.TblBuyLanReEst.opening_balance_voucher_id, dbo.TblBuyLanReEst.NoteSerial, dbo.TblBuyLanReEst.NoteID, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "   dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee"
MySQL = MySQL & "   FROM     dbo.TblBuyLanReEst INNER JOIN"
MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblBuyLanReEst.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
MySQL = MySQL & "   dbo.TblCustemers ON dbo.TblBuyLanReEst.OwnerID = dbo.TblCustemers.CusID"

 MySQL = MySQL & "    where 1 = 1"
 
 
If dcBranch6.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblBuyLanReEst.BranchID  =  " & val(dcBranch6.BoundText)
End If
 
If dcsupplier.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblBuyLanReEst.OwnerID  =  " & val(dcsupplier.BoundText)
End If

If NameTxt.Text <> "" Then
        MySQL = MySQL & " and  TblBuyLanReEst.Name  like '%" & NameTxt.Text & "%'"
End If

If TxtFullcode.Text <> "" Then
        MySQL = MySQL & " and  TblBuyLanReEst.FullCode  like '%" & TxtFullcode.Text & "%'"
End If

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BuyLandReal.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BuyLandReal.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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


Function Print_Expenses(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""
    
 MySQL = MySQL & "   SELECT dbo.TblExpensesInvesment.ID, dbo.TblExpensesInvesment.RecordDate, dbo.TblExpensesInvesment.BranchID, dbo.TblExpensesInvesment.UserID,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.InvesID, dbo.TblExpensesInvesment.CurrValue, dbo.TblExpensesInvesment.DevlopValue, dbo.TblExpensesInvesment.AfterDevlopValue,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.ShareValue, dbo.TblExpensesInvesment.SharNo, dbo.TblExpensesInvesment.Remarks, dbo.TblExpensesInvesment.Total,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.ShareValueNew, dbo.TblExpensesInvesment.LandID, dbo.TblExpensesInvesment.DivPayed, dbo.TblExpensesInvesment.NoteSerial,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.NoteID, dbo.TblBranchesData.branch_name, dbo.Tblinvestment.Name AS InvestmentName, dbo.TblBuyLanReEst.Name AS LandName,"
 MySQL = MySQL & "   dbo.TblBuyLanReEst.ID AS LandCode, dbo.Tblinvestment.ID AS InvestmentCode, SUM(dbo.TblExpensesInvesmentDet.Valu) AS value"
 MySQL = MySQL & "   FROM     dbo.TblExpensesInvesment LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblExpensesInvesmentDet ON dbo.TblExpensesInvesment.ID = dbo.TblExpensesInvesmentDet.ExpInvID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblBuyLanReEst ON dbo.TblExpensesInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.Tblinvestment ON dbo.TblExpensesInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
 MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblExpensesInvesment.BranchID = dbo.TblBranchesData.branch_id"
 


 MySQL = MySQL & "    where 1 = 1"
 
 
If DcbInvise.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblExpensesInvesment.InvesID  =  " & val(DcbInvise.BoundText)
End If
 
If DcbLand.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblExpensesInvesment.LandID  =  " & val(DcbLand.BoundText)
End If


 MySQL = MySQL & "   GROUP BY dbo.TblExpensesInvesment.ID, dbo.TblExpensesInvesment.RecordDate, dbo.TblExpensesInvesment.BranchID, dbo.TblExpensesInvesment.UserID,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.InvesID, dbo.TblExpensesInvesment.CurrValue, dbo.TblExpensesInvesment.DevlopValue, dbo.TblExpensesInvesment.AfterDevlopValue,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.ShareValue, dbo.TblExpensesInvesment.SharNo, dbo.TblExpensesInvesment.Remarks, dbo.TblExpensesInvesment.Total,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.ShareValueNew, dbo.TblExpensesInvesment.LandID, dbo.TblExpensesInvesment.DivPayed, dbo.TblExpensesInvesment.NoteSerial,"
 MySQL = MySQL & "   dbo.TblExpensesInvesment.NoteID , dbo.TblBranchesData.branch_name, dbo.Tblinvestment.name, dbo.TblBuyLanReEst.name, dbo.TblBuyLanReEst.ID, dbo.Tblinvestment.ID"



 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Expenses.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Expenses.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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

Function Print_LandDivid(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""

MySQL = MySQL & "   SELECT dbo.TblDivInvesment.*, dbo.Tblinvestment.Name AS InvestmentName, dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBranchesData.branch_name"
MySQL = MySQL & " , dbo.Tblinvestment.ID AS InvestmentCode, dbo.TblBuyLanReEst.ID AS LandCode"
MySQL = MySQL & "   FROM     dbo.TblDivInvesment LEFT OUTER JOIN"
MySQL = MySQL & "   dbo.TblBuyLanReEst ON dbo.TblDivInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
MySQL = MySQL & "   dbo.Tblinvestment ON dbo.TblDivInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblDivInvesment.BranchID = dbo.TblBranchesData.branch_id"

MySQL = MySQL & "    where 1 = 1"
 
 
If dcInvestment8.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblDivInvesment.InvesID  =  " & val(dcInvestment8.BoundText)
End If
 
If DcbLand8.BoundText <> "" Then
        MySQL = MySQL & " and dbo.TblDivInvesment.LandID  =  " & val(DcbLand8.BoundText)
End If

If RdType(0).value = True Then
        MySQL = MySQL & " and dbo.TblDivInvesment.TypDiv = 0 "
ElseIf RdType(1).value = True Then
        MySQL = MySQL & " and dbo.TblDivInvesment.TypDiv = 1 "
End If

 

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_LandDivid.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_LandDivid.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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

Function Print_InvestmentCustomers()
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""


MySQL = MySQL & "  SELECT dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_Phone,"
MySQL = MySQL & "  dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Type, dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType,"
MySQL = MySQL & "  dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.CreditLimit, dbo.TblCustemers.Account_Code_As_Client, dbo.TblCustemers.Account_Code_As_Supplier,"
MySQL = MySQL & "  dbo.TblCustemers.CreditlimitCredit, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.SaleType, dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "  dbo.TblCustemers.Trans_Discount, dbo.TblCustemers.Trans_DiscountType, dbo.TblCustemers.CountryID, dbo.TblCustemers.GovernmentID, dbo.TblCustemers.CityID,"
MySQL = MySQL & "  dbo.TblCustemers.Address, dbo.TblCustemers.Trans_DiscountPur, dbo.TblCustemers.Trans_DiscountTypePur, dbo.TblCustemers.CountEmp, dbo.TblCustemers.ToTal,"
MySQL = MySQL & "  dbo.TblCustemers.c1, dbo.TblCustemers.c2, dbo.TblCustemers.Remark2, dbo.TblCustemers.locked, dbo.TblCustemers.parent_account,"
MySQL = MySQL & "  dbo.TblCustemers.opening_balance_voucher_id, dbo.TblCustemers.DepitInterval, dbo.TblCustemers.CreditInterval, dbo.TblCustemers.DepitIntervalID,"
MySQL = MySQL & "  dbo.TblCustemers.CreditIntervalID, dbo.TblCustemers.EmpId, dbo.TblCustemers.prifix, dbo.TblCustemers.code, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "  dbo.TblCustemers.CustomerandVendor, dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.BranchId, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH,"
MySQL = MySQL & "  dbo.TblCustemers.Company, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Salary, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel,"
MySQL = MySQL & "  dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.CountryID2, dbo.TblCustemers.Sex,"
MySQL = MySQL & "  dbo.TblCustemers.Account_Code1, dbo.TblCustemers.Account_Code2, dbo.TblCustemers.ParentAccount, dbo.TblCustemers.OpenBalanceType1,"
MySQL = MySQL & "  dbo.TblCustemers.OpenBalance1, dbo.TblCustemers.OpenBalanceType2, dbo.TblCustemers.OpenBalance2, dbo.TblCustemers.ShowQty1, dbo.TblCustemers.showPrice1,"
MySQL = MySQL & "  dbo.TblCustemers.showPrice2, dbo.TblCustemers.Salaries1, dbo.TblCustemers.Salaries2, dbo.TblCustemers.ShowQty1c, dbo.TblCustemers.showPrice1c,"
MySQL = MySQL & "  dbo.TblCustemers.showPrice2c, dbo.TblCustemers.Salaries1c, dbo.TblCustemers.Salaries2c, dbo.TblCustemers.Totald, dbo.TblCustemers.Totalc,"
MySQL = MySQL & "  dbo.TblCustemers.balanced, dbo.TblCustemers.balancec, dbo.TblCustemers.TypeCustomer, dbo.TblCustemers.BoxMil, dbo.TblCustemers.ZipCode,"
MySQL = MySQL & "  dbo.TblCustemers.RecordDate, dbo.TblCustemers.Entry, dbo.TblCustemers.JobName, dbo.TblCustemers.Map, dbo.TblCustemers.authorizationname,"
MySQL = MySQL & "  dbo.TblCustemers.authorizationNo, dbo.TblCustemers.InsuranceAccount, dbo.TblCustemers.Owner, dbo.TblCustemers.creditlocked, dbo.TblCustemers.BoxNo,"
MySQL = MySQL & "  dbo.TblCustemers.PostalCode, dbo.TblCustemers.RsID, dbo.TblCustemers.RsDegree, dbo.TblCustemers.RsIDDate, dbo.TblCustemers.BankAccount,"
MySQL = MySQL & "  dbo.TblCustemers.BankName, dbo.TblCustemers.RecordNo, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.Category,"
MySQL = MySQL & "  dbo.TblCustemers.BankID, dbo.TblCustemers.Flg, dbo.TblCustemers.TypeInvestor, dbo.TblCustemers.GroupInvestor, dbo.TblCustemers.IBAN, dbo.TblCustemers.Typ,"
MySQL = MySQL & "  dbo.TblCustemers.BankCode , dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankAddress, dbo.TblCustemers.CustGIDPlace, dbo.TblBranchesData.branch_name"
MySQL = MySQL & "  FROM     dbo.TblCustemers INNER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "  Where (dbo.TblCustemers.Type = 20)"

 
If dcInvestmentType9.ListIndex <> -1 Then
        MySQL = MySQL & " and dbo.TblCustemers.typ  =  " & val(dcInvestmentType9.ListIndex)
End If
 
If txtName9.Text <> "" Then
        MySQL = MySQL & " and dbo.TblCustemers.CusName  like  '%" & txtName9.Text & "%'"
End If

If txtCode9.Text <> "" Then
        MySQL = MySQL & " and dbo.TblCustemers.FullCode  like  '%" & txtCode9.Text & "%'"
End If

If DcBranch9.BoundText <> "" Then
        MySQL = MySQL & " and TblCustemers.BranchId =  " & val(DcBranch9.BoundText)
End If
 

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_InvestmentCustomer.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_InvestmentCustomer.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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

