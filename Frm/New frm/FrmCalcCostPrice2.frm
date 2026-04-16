VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCalcCostPrice2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ”«» «· ﬂ«·»› €Ì— «·„»«‘—… Ê Ê“Ì⁄Â«"
   ClientHeight    =   8985
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10290
   HelpContextID   =   580
   Icon            =   "FrmCalcCostPrice2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   10290
   Visible         =   0   'False
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
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10245
      _cx             =   18071
      _cy             =   15849
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
      _GridInfo       =   $"FrmCalcCostPrice2.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7950
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10185
         _cx             =   17965
         _cy             =   14023
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
         Caption         =   "."
         Align           =   0
         CurrTab         =   0
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
            Height          =   7530
            Index           =   2
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   10095
            _cx             =   17806
            _cy             =   13282
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
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
               Height          =   720
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   6720
               Width           =   5775
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox txtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   240
                  Width           =   2160
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   11
                  Left            =   120
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   315
                  Index           =   18
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   240
                  Width           =   720
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   " ﬁ«—Ì— «·«‰ «Ã"
               Height          =   705
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   6720
               Visible         =   0   'False
               Width           =   3975
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   7
                  Left            =   1680
                  TabIndex        =   20
                  Top             =   240
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   " ﬁ—Ì— „·Œ’ «·«‰ «Ã «·‰„ÿÌ"
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
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   8
                  Left            =   -480
                  TabIndex        =   21
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   " ﬁ—Ì— «—»«Õ «·„»Ì⁄« "
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
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   0
               Width           =   10035
               _cx             =   17701
               _cy             =   1349
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   24
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
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Picture         =   "FrmCalcCostPrice2.frx":040F
               Caption         =   "Õ”«» «· ﬂ«·»› €Ì— «·„»«‘—… Ê Ê“Ì⁄Â«"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   0
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
               PicturePos      =   0
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
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   28
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmCalcCostPrice2.frx":10E9
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
                  Height          =   375
                  Index           =   2
                  Left            =   630
                  TabIndex        =   29
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmCalcCostPrice2.frx":1483
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
                  Height          =   375
                  Index           =   1
                  Left            =   2220
                  TabIndex        =   30
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmCalcCostPrice2.frx":181D
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
                  Height          =   375
                  Index           =   3
                  Left            =   1155
                  TabIndex        =   31
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmCalcCostPrice2.frx":1BB7
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6675
               Index           =   1
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   120
               Width           =   10425
               _cx             =   18389
               _cy             =   11774
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
                  Caption         =   "„»Ì⁄«  «·Œœ„« "
                  Height          =   615
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   5040
                  Visible         =   0   'False
                  Width           =   4215
                  Begin VB.CheckBox chkProfitService 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check1"
                     Height          =   255
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Tag             =   "Ì „ «Œ Ì«—Â« ›Ì Õ«·… «·—€»… ›Ì «÷«› Â« ·›Ì„… «·„»Ì⁄«  ·“Ì«œ… «·—»Õ"
                     Top             =   240
                     Width           =   255
                  End
                  Begin VB.TextBox TxtServicesValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   240
                     Width           =   2160
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ì—«œ«  «·Œœ„« "
                     ForeColor       =   &H00000000&
                     Height          =   405
                     Index           =   2
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   240
                     Width           =   1320
                  End
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4395
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   2220
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   -4110
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   10530
                  Width           =   2265
               End
               Begin VB.TextBox TxtTypicalProductionId 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   8415
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   375
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.TextBox txtRemarks 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   4260
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   61
                  Top             =   1290
                  Width           =   5205
               End
               Begin VB.TextBox TxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7920
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   840
                  Width           =   1530
               End
               Begin VB.Frame Frame2 
                  Caption         =   "«·„’—Ê›«  Œ·«· «·› —…"
                  Enabled         =   0   'False
                  Height          =   3045
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   1650
                  Width           =   5685
                  Begin VB.TextBox TxtExpenses 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   360
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtSalaryVouchersTotals 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   720
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtAccDep 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   1800
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtMaterialIssueVoucherTotals 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   2160
                     Visible         =   0   'False
                     Width           =   2160
                  End
                  Begin VB.TextBox Txttotal 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   2640
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtAllocations 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   1080
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtAllocations1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   1440
                     Width           =   2160
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›«    Ê «·›Ê« Ì— «·„«·Ì…"
                     Height          =   405
                     Index           =   22
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   360
                     Width           =   2880
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì ﬁÌ„… ”‰œ«  «·—« » ··› —…"
                     Height          =   420
                     Index           =   24
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   720
                     Width           =   3000
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì ﬁÌ„… ”‰œ«  «·«Â·«ﬂ ··› —…"
                     Height          =   420
                     Index           =   25
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   1800
                     Width           =   2280
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬁÌ„… ’—› «·„Ê«œ «·Œ«„ ··› —… ··› —…"
                     Height          =   405
                     Index           =   26
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   2160
                     Visible         =   0   'False
                     Width           =   2640
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì    ﬂ·›… «·«‰ «Ã ⁄‰ «·› —…"
                     ForeColor       =   &H00FF0000&
                     Height          =   405
                     Index           =   15
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   2640
                     Width           =   2160
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H000000FF&
                     X1              =   120
                     X2              =   5400
                     Y1              =   2520
                     Y2              =   2520
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì ﬁÌ„…    „Œ’’«  «·«Ã«“…"
                     Height          =   420
                     Index           =   29
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   1080
                     Width           =   3000
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì ﬁÌ„…     „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
                     Height          =   420
                     Index           =   30
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   1440
                     Width           =   3000
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "”⁄— «· ﬂ·›… ··ÊÕœ… „"
                  Enabled         =   0   'False
                  Height          =   930
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   4695
                  Width           =   5685
                  Begin VB.TextBox TxtTotalProductionQty 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   240
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtUnitValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   600
                     Width           =   2160
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·ﬂ„Ì«  «·„‰ Ã… Œ·«· «·› —…"
                     ForeColor       =   &H00000000&
                     Height          =   225
                     Index           =   13
                     Left            =   3105
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   360
                     Width           =   2400
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "‰’Ì» «·ÊÕœ… „‰ «·„’—Ê›«  €Ì— «·„»«‘—…"
                     Height          =   315
                     Index           =   17
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   600
                     Width           =   3000
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "«·„»Ì⁄« "
                  Enabled         =   0   'False
                  Height          =   960
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   5610
                  Visible         =   0   'False
                  Width           =   9945
                  Begin VB.TextBox TxtSaLePayValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   240
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtSaleValue 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   600
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtTotalsalesQty 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   120
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtProfit 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   600
                     Width           =   2160
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬁÌ„… „»Ì⁄«  «·› —…"
                     ForeColor       =   &H00000000&
                     Height          =   405
                     Index           =   12
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   240
                     Width           =   1320
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   " ﬂ·›… «·„»Ì⁄«  Œ·«· «·› —…"
                     ForeColor       =   &H00FF0000&
                     Height          =   315
                     Index           =   19
                     Left            =   7800
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   600
                     Width           =   1800
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì    «·ﬂ„Ì… «·„»«⁄Â ⁄‰ «·› —…"
                     ForeColor       =   &H00000000&
                     Height          =   405
                     Index           =   27
                     Left            =   6840
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   240
                     Width           =   2760
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«—»«Õ «·› —…"
                     ForeColor       =   &H00000000&
                     Height          =   405
                     Index           =   28
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   600
                     Width           =   1320
                  End
               End
               Begin MSComCtl2.DTPicker dbRecordDate 
                  Height          =   285
                  Left            =   4260
                  TabIndex        =   65
                  Top             =   810
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   102957057
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DCIntervals 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1200
                  Width           =   3180
                  _ExtentX        =   5609
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTStartdate 
                  Height          =   285
                  Left            =   2010
                  TabIndex        =   67
                  Top             =   1560
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   102957057
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTEndDate 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1560
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   102957057
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   69
                  Top             =   780
                  Width           =   3180
                  _ExtentX        =   5609
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Â–… «·‘«‘…  Ê›— Œ«’Ì…  Õ”«» ‰’Ì» «·«’‰«› „‰ «· ﬂ«·Ì› €Ì— «·„»«‘—… À„  Ê“Ì⁄ Â–… «· ﬂ·›… «·Ï  ﬂ«·Ì› «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   2490
                  Index           =   0
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   2160
                  Width           =   3735
               End
               Begin VB.Shape Shape1 
                  BorderWidth     =   2
                  FillColor       =   &H00C0FFFF&
                  FillStyle       =   0  'Solid
                  Height          =   2700
                  Left            =   0
                  Top             =   2130
                  Width           =   4155
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
                  Height          =   465
                  Left            =   14415
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   1050
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   480
                  Index           =   7
                  Left            =   8085
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   810
                  Width           =   1905
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   5
                  Left            =   7260
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   810
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   3
                  Left            =   9315
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1290
                  Width           =   750
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·› —…"
                  Height          =   285
                  Index           =   14
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1320
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   45
                  Index           =   20
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1560
                  Width           =   225
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   285
                  Index           =   21
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   1560
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·›—⁄"
                  Height          =   285
                  Index           =   16
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   810
                  Width           =   825
               End
               Begin VB.Label lbl 
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
                  Height          =   255
                  Index           =   37
                  Left            =   2745
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   1830
                  Width           =   1380
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7995
         Width           =   10185
         _cx             =   17965
         _cy             =   1693
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "FrmCalcCostPrice2.frx":1F51
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            ButtonImage     =   "FrmCalcCostPrice2.frx":22EB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCalcCostPrice2.frx":2685
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7140
            TabIndex        =   9
            Top             =   510
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   6240
            TabIndex        =   10
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            Height          =   495
            Index           =   2
            Left            =   5400
            TabIndex        =   11
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   4395
            TabIndex        =   12
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   4
            Left            =   3360
            TabIndex        =   13
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            CausesValidation=   0   'False
            Height          =   495
            Index           =   6
            Left            =   480
            TabIndex        =   14
            Top             =   510
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            Height          =   495
            Index           =   5
            Left            =   2430
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   9120
            TabIndex        =   16
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCalcCostPrice2.frx":2A1F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   9
            Left            =   1080
            TabIndex        =   17
            Top             =   600
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
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
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
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
      ButtonImage     =   "FrmCalcCostPrice2.frx":2A3B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmCalcCostPrice2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim strSQL  As String
Dim rs As ADODB.Recordset
Dim SalesDifferentValue As Double

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal x As Long, _
                                  ByVal y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Public Sub YearMonth()

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Private Sub CboPayMentType_Click()
 
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        If val(Me.DcBranch.BoundText) = 0 Then
            Msg = "ÌÃ» ≈Œ Ì«— «·›—⁄..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcBranch.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
        If val(Me.DCIntervals.BoundText) = 0 Then
            Msg = "ÌÃ» ≈Œ Ì«— «·› —…..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCIntervals.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
    End If

    '-------------------------------------------------------------------------------------------
             
    If txtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), dbRecordDate.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), dbRecordDate.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                txtNoteSerial.text = Notes_coding(val(my_branch), dbRecordDate.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1.text = "" Then
        If Voucher_coding(val(my_branch), dbRecordDate.value, 24, 103) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ﬂ«·Ì› «‰ «Ã ‰„ÿÌ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Voucher_coding(val(my_branch), dbRecordDate.value, 24, 103) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ   ﬂ«·Ì› «‰ «Ã ‰„ÿÌ  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), dbRecordDate.value, 24, 103)
            End If
        End If
    End If
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
        TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
    ElseIf Me.TxtModFlg.text = "E" Then
     
        strSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
        Cn.Execute strSQL, , adExecuteNoRecords
   
        strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.text)
        Cn.Execute strSQL, , adExecuteNoRecords
   
    End If
    
    rs("id").value = TxtTypicalProductionId.text
    
    rs("intervalid").value = IIf(Me.DCIntervals.BoundText = "", Null, Me.DCIntervals.BoundText)
    rs("branch_no").value = IIf(Me.DcBranch.BoundText = "", Null, Me.DcBranch.BoundText)
    
    rs("RecordDate").value = dbRecordDate.value
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
    rs("Expenses").value = IIf(val(Me.TxtExpenses.text) = 0, 0, val(Me.TxtExpenses.text))
 
    rs("SalaryVouchersTotals").value = IIf(val(Me.TxtSalaryVouchersTotals.text) = 0, 0, val(Me.TxtSalaryVouchersTotals.text))
    rs("Allocations").value = IIf(val(Me.TxtAllocations.text) = 0, 0, val(Me.TxtAllocations.text))
    rs("Allocations1").value = IIf(val(Me.TxtAllocations1.text) = 0, 0, val(Me.TxtAllocations1.text))
 
    rs("MaterialIssueVoucherTotals").value = IIf(val(Me.TxtMaterialIssueVoucherTotals.text) = 0, 0, val(Me.TxtMaterialIssueVoucherTotals.text))
    rs("AccDep").value = IIf(val(Me.TxtAccDep.text) = 0, 0, val(Me.TxtAccDep.text))
    rs("total").value = IIf(val(Me.Txttotal.text) = 0, 0, val(Me.Txttotal.text))
    rs("UnitValue").value = IIf(val(Me.TxtUnitValue.text) = 0, 0, val(Me.TxtUnitValue.text))
  
    rs("SaleValue").value = IIf(val(Me.TxtSaleValue.text) = 0, 0, val(Me.TxtSaleValue.text))
    rs("TotalProductionQty").value = IIf(val(Me.TxtTotalProductionQty.text) = 0, 0, val(Me.TxtTotalProductionQty.text))
    rs("TotalsalesQty").value = IIf(val(Me.TxtTotalsalesQty.text) = 0, 0, val(Me.TxtTotalsalesQty.text))
    rs("NoteID").value = IIf(val(Me.TxtNoteID.text) = 0, 0, val(Me.TxtNoteID.text))
    rs("NoteSerial").value = IIf(Me.txtNoteSerial.text = "", "", Me.txtNoteSerial.text)
 
    rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1.text = "", "", Me.TxtNoteSerial1.text)
 
    rs("SaLePayValue").value = IIf(val(Me.TxtSaLePayValue.text) = 0, 0, val(Me.TxtSaLePayValue.text))
    '  rs("SalesValue1").value = IIf(Val(Me.TxtServicesValue.text) = 0, 0, Val(Me.TxtServicesValue.text))
  
    If chkProfitService.value = vbChecked Then
        '  rs("ProfitService").value = 1
    Else
        'rs("ProfitService").value = 0
    End If

    rs("Profit").value = IIf(val(Me.TxtProfit.text) = 0, 0, val(Me.TxtProfit.text))
 
    rs.update
 
    createVoucher
    
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function createVoucher()
 
    Dim bankDes As String
    Dim AccountCode As String
 
    Dim NoteID As String
    Dim sql As String
    Dim Msg As String
    Dim ProductionStoreId As Long
    ProductionStoreId = GetProductionInventoryId(val(Me.DcBranch.BoundText))

    If ProductionStoreId = 0 Then
    
        Msg = " ·« ÌÊÃœ „Œ“‰ «‰ «Ã  «„ ·œÌﬂ..!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcBranch.SetFocus
        SendKeys "{F4}"
        
        Exit Function
    End If
        
    Dim Account_Code_dynamic2 As String
    Account_Code_dynamic2 = get_account_code_branch(1, my_branch)
        
    If Account_Code_dynamic2 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic2 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ﬂ·›… «·„»Ì⁄«              ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic3 As String
    Account_Code_dynamic3 = get_account_code_branch(55, my_branch)
        
    If Account_Code_dynamic3 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic3 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    „Œ’’ «·«Ã«“…             ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic4 As String
    Account_Code_dynamic4 = get_account_code_branch(56, my_branch)
        
    If Account_Code_dynamic4 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic4 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    „Œ’’ „ﬂ«›√… ‰Â«Ì… «·Œœ„…             ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic7 As String '–„„ «·„ÊŸ›Ì‰
    Account_Code_dynamic7 = get_account_code_branch(7, my_branch)
        
    If Account_Code_dynamic7 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic7 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    –„„ «·„ÊŸ›Ì‰          ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic29 As String '«·«ÃÊ— «·„” Õﬁ… «·„ÊŸ›Ì‰
    Account_Code_dynamic29 = get_account_code_branch(29, my_branch)
        
    If Account_Code_dynamic29 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic29 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·«ÃÊ—  «·„” Õﬁ… «·„ÊŸ›Ì‰          ", vbCritical
        
            Exit Function
        End If
    End If
  
    If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = "  Ê“Ì⁄ «·„’—Ê›«  €Ì— «·„»«‘—…      " & Me.DCIntervals.text & "   „‰ " & DTStartdate.value & "  «·Ï " & DTEndDate.value
    Else
        bankDes = " Calc Indirect Cost For Period  " & Me.DCIntervals.text
  
    End If

    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
  '  RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     strSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Me.TxtModFlg.text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
     
    End If

    RsNotes.AddNew
    NoteID = CStr(TxtNoteID.text)
    RsNotes("NoteID").value = CStr(TxtNoteID.text)
    RsNotes("NoteType").value = 103
    RsNotes("NoteDate").value = dbRecordDate.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.txtNoteSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”·   ”‰œ  ﬂ«·Ì› «·«‰ «Ã «·‰„ﬂÌ
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(24) '‰Ê⁄  —ﬁÌ„ ”‰œ «·«Ìœ«⁄
    RsNotes("sanad_year").value = year(dbRecordDate.value)
    RsNotes("sanad_month").value = Month(dbRecordDate.value)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(Txttotal.text) + val(TxtSaleValue.text), "0.00"), 0, True, ".")
    RsNotes("remark").value = txtRemarks.text & bankDes
    RsNotes("Branch_no").value = val(Me.DcBranch.BoundText)
                
    RsNotes.update
                
    line_no = 1
 
    Dim RsDev  As ADODB.Recordset
    Set RsDev = New ADODB.Recordset
    'RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                  strSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    '«·ÿ—› «·„œÌ‰     «·„Œ“Ê‰
    AccountCode = ModAccounts.GetMyAccountCode("TblStore", "StoreID", ProductionStoreId)

    If val(Txttotal.text) > 0 Then
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(Me.Txttotal.text)
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    '
    '”‰œ«  «·—« »
    Dim i  As Integer
    Dim LngDevID  As Long
    
    Dim SQLSalaryExpenses    As String
    Dim RsSalaryV As ADODB.Recordset
    Set RsSalaryV = New ADODB.Recordset
    
    SQLSalaryExpenses = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit ,dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    SQLSalaryExpenses = SQLSalaryExpenses & " FROM         dbo.Notes INNER JOIN"
    SQLSalaryExpenses = SQLSalaryExpenses & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    SQLSalaryExpenses = SQLSalaryExpenses & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    SQLSalaryExpenses = SQLSalaryExpenses & " WHERE     (dbo.Notes.NoteType = 66) AND (dbo.Notes.NoteDate >= " & SQLDate(Me.DTStartdate.value, True) & ") AND (dbo.Notes.NoteDate <=" & SQLDate(Me.DTEndDate.value, True) & ") AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    SQLSalaryExpenses = SQLSalaryExpenses & " AND (branch_no = " & val(DcBranch.BoundText) & ")"
    SQLSalaryExpenses = SQLSalaryExpenses & "   ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit DESC"
 
    Dim Credit_Or_Debit As Integer
  
    RsSalaryV.Open SQLSalaryExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For i = 1 To RsSalaryV.RecordCount

        If Not IsNull(RsSalaryV("Account_Code").value) And (RsSalaryV("Value").value) > 0 Then
               
            AccountCode = (RsSalaryV("Account_Code").value)
            line_no = line_no + 1

            If RsSalaryV("Credit_Or_Debit").value = 0 Then
                Credit_Or_Debit = 1
            Else
                Credit_Or_Debit = 0
            End If
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(RsSalaryV("Value").value), Credit_Or_Debit, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsSalaryV("Value").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        End If

        RsSalaryV.MoveNext
    Next i
 
    Set RsSalaryV = Nothing

    '      «·„’—Ê›«  Ê «·›Ê« Ì— «·„«·Ì…
 
    Dim SQLExpenses As String
    Dim RsExpenseV As ADODB.Recordset
    Set RsExpenseV = New ADODB.Recordset
  
    SQLExpenses = "SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.ACCOUNTS.Account_Code"
    SQLExpenses = SQLExpenses & "  FROM         dbo.ACCOUNTS INNER JOIN"
    SQLExpenses = SQLExpenses & "   dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    SQLExpenses = SQLExpenses & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    SQLExpenses = SQLExpenses & " WHERE     (dbo.ExpensesType.IndirectCosts = 1) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    SQLExpenses = SQLExpenses & "       RecordDate >= " & SQLDate(Me.DTStartdate.value, True)
    SQLExpenses = SQLExpenses & "  AND RecordDate <= " & SQLDate(Me.DTEndDate.value, True)
    SQLExpenses = SQLExpenses & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
                     
    SQLExpenses = SQLExpenses & "  GROUP BY dbo.ACCOUNTS.Account_Code"
  
    RsExpenseV.Open SQLExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For i = 1 To RsExpenseV.RecordCount

        If Not IsNull(RsExpenseV("Account_Code").value) And (RsExpenseV("Total").value) > 0 Then
               
            AccountCode = (RsExpenseV("Account_Code").value)
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(RsExpenseV("Total").value), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsExpenseV("Total").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        End If

        RsExpenseV.MoveNext
    Next i
 
    Set RsExpenseV = Nothing
            
    ' „’«—Ì› «·«Â·«ﬂ
    SQLExpenses = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS total"
    SQLExpenses = SQLExpenses & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    SQLExpenses = SQLExpenses & " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
    SQLExpenses = SQLExpenses & "  WHERE      (dbo.Notes.NoteType = 90) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    SQLExpenses = SQLExpenses & "       RecordDate >= " & SQLDate(Me.DTStartdate.value, True)
    SQLExpenses = SQLExpenses & "  AND RecordDate <= " & SQLDate(Me.DTEndDate.value, True)
    SQLExpenses = SQLExpenses & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
                     
    SQLExpenses = SQLExpenses & "  GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
  
    Set RsExpenseV = New ADODB.Recordset
    RsExpenseV.Open SQLExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For i = 1 To RsExpenseV.RecordCount

        If Not IsNull(RsExpenseV("Account_Code").value) And (RsExpenseV("Total").value) > 0 Then
               
            AccountCode = (RsExpenseV("Account_Code").value)
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(RsExpenseV("Total").value), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsExpenseV("Total").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        End If

        RsExpenseV.MoveNext
    Next i
    
    ' ﬁÌœ ' ﬂ·›… «·„»Ì⁄« 
    '„œÌ‰
    If SalesDifferentValue > 0 Then
        AccountCode = Account_Code_dynamic2
        line_no = line_no + 1
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = SalesDifferentValue
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = " ﬁÌœ  ﬂ·›… «·„»Ì⁄«  " & bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = " Sales Cost Vchr" & bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    'œ«∆‰
    If SalesDifferentValue > 0 Then
        AccountCode = ModAccounts.GetMyAccountCode("TblStore", "StoreID", ProductionStoreId)
        line_no = line_no + 1
          
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = SalesDifferentValue
        RsDev("Credit_Or_Debit").value = 1
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = " ﬁÌœ  ﬂ·›… «·„»Ì⁄«  " & bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = " Sales Cost Vchr" & bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
                    
    End If
 
    '  «·«Ã«“… „’«—Ì›       «·„Œ’’« 
    If val(TxtAllocations.text) > 0 Then
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic3, val(TxtAllocations.text), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtSalaryVouchersTotals), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
    End If

    '  ‰Â«Ì… «·Œœ„… „’«—Ì›       «·„Œ’’« 
    If val(TxtAllocations1.text) > 0 Then
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic4, val(TxtAllocations1.text), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtSalaryVouchersTotals), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
    End If

ErrTrap:
End Function

Function IndirectCostAddition(IndirectCost As Double, FromDate As Date, todate As Date, Transaction_Type As Integer)
    'On Error GoTo ErrTrap
    Dim strSQL As String

    ' StrSQL = "update dbo.Transaction_Details  set Price=" & cost & ", CostPrice=" & cost
    If IndirectCost = 0 Then '›Ì Õ«·Â Õ–› «·”‰œ
        strSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",Price=OldPrice,ShowPrice=OldShowPrice"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transactions.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL
    Else
        strSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",OldPrice=Price,OldShowPrice=ShowPrice,Price=Price+" & IndirectCost & ",ShowPrice=ShowPrice+(" & IndirectCost & "*QtyBySmalltUnit)"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transactions.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL
    End If
 
    '  StrSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

    '     StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(fromdate, True)
    '      StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
    '           StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
    ' StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & Val(DcBranch.BoundText) & ")"
 
    ' Cn.Execute StrSQL
 
    'If
    Dim FirstPeriodDateInthisYear  As Date
 
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    Dim fromdateS As Variant
    Dim todateS As Variant
    fromdateS = Replace(Format$(FirstPeriodDateInthisYear, "MM/DD/yyyy"), "-", "/")
    todateS = Replace(Format$(DTEndDate.value, "MM/DD/yyyy"), "-", "/")

    Transaction_Type = 19 '”‰œ«  «·’—›
 
    If IndirectCost = 0 Then
        ' StrSQL = "update dbo.Transaction_Details  set Price= 0"
        'StrSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",OldPrice=Price,OldShowPrice=ShowPrice,Price=Price+" & IndirectCost & "ShowPrice=ShowPrice+(" & IndirectCost & "*QtyBySmalltUnit)"
        strSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",Price=OldPrice,ShowPrice=OldShowPrice"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 
        Cn.Execute strSQL
 
        strSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL
 
    Else
        'StrSQL = "update dbo.Transaction_Details  set Price= dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID) , CostPrice=dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID)"
        strSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",OldPrice=Price,OldShowPrice=ShowPrice, Price= dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID) "

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 
        Cn.Execute strSQL
 
        strSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL

    End If

    Transaction_Type = 21 '  ⁄œÌ· „ Ê”ÿ «· ﬂ·›… ·Õ”«» «·—»Õ '  ›Ê« Ì— «·„»Ì⁄« 
 
    If IndirectCost = 0 Then
        ' StrSQL = "update dbo.Transaction_Details  set Price= 0"
        'StrSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",OldPrice=Price,OldShowPrice=ShowPrice,Price=Price+" & IndirectCost & "ShowPrice=ShowPrice+(" & IndirectCost & "*QtyBySmalltUnit)"
        strSQL = "update dbo.Transaction_Details  set IndirectCost=" & IndirectCost & ",CostPrice=OldCostPrice"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 
        Cn.Execute strSQL
 
        strSQL = "update dbo.Transaction_Details  set ItemProfit=(ShowPrice-CostPrice)*Showqty"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL
 
    Else
        'StrSQL = "update dbo.Transaction_Details  set Price= dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID) , CostPrice=dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID)"
        strSQL = "update dbo.Transaction_Details  set OldCostPrice=CostPrice , CostPrice= dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID)*QtyBySmalltUnit "

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 
        Cn.Execute strSQL
 
        strSQL = "update dbo.Transaction_Details  set ItemProfit=(ShowPrice-CostPrice)*Showqty"

        strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
        strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        strSQL = strSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
        strSQL = strSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute strSQL

    End If

    'End If
ErrTrap:
End Function

Private Sub Cmd_Click(Index As Integer)
 
    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
 
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.text = "N"
            clear_all Me
            Me.TxtTypicalProductionId.text = CStr(new_id("TblTypicalProduction2", "id", "", True))
        
            Me.DcBranch.BoundText = branch_id
            chkProfitService.value = vbChecked
         
        Case 1
                    If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
         
        Case 2
                     If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            my_branch = Me.DcBranch.BoundText
  
            Dim StartDate As Date
            Dim EndDate As Date

            If val(DCIntervals.BoundText) = 0 Then Exit Sub

            GetIntervalsFullData val(DCIntervals.BoundText), StartDate, EndDate
            DTStartdate.value = StartDate
            DTEndDate.value = EndDate
            GetAllTotals StartDate, EndDate
            SaveData
           
        Case 3
            Undo

        Case 4
                    If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            PrintReport

            '   ViewDataList
        Case 9
    
        Case 20
     
        Case 21
            RemoveGridRow

        Case 11

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.txtNoteSerial.text, , 200
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap
    Dim ItemReport As ClsItemsReport

    If TxtTypicalProductionId.text <> "" Then
        Set ItemReport = New ClsItemsReport
        ItemReport.TypicalProduction val(TxtTypicalProductionId.text), Me.DCIntervals.text, Me.DcBranch.text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim strSQL As String
    On Error GoTo ErrTrap

    If Me.TxtTypicalProductionId.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (TxtNoteSerial1.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
      
            strSQL = "Delete From notes Where NoteID=" & val(TxtNoteID.text)
            Cn.Execute strSQL, , adExecuteNoRecords
 
            If Not rs.RecordCount < 1 Then
                IndirectCostAddition 0, DTStartdate.value, DTEndDate.value, 28

                DoEvents
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
        
                    clear_all Me
                
                    ' XPTxtCurrent.Caption = 0
                    '          XPTxtCount.Caption = 0
                          
                    TxtModFlg_Change
           
                Else
                    Retrive
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

Private Sub RemoveGridRow()
 
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub dbRecordDate_Change()
    TxtNoteSerial1.text = ""
    txtNoteSerial.text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    DCIntervals_Click 0
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub
 
Private Sub Dcbranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches DcBranch
    End If

End Sub

Private Sub DCIntervals_Change()

    If Me.TxtModFlg = "R" Or Me.TxtModFlg = "" Then
        DCIntervals_Click 0
        dbRecordDate_Change

    End If

End Sub

Private Sub DCIntervals_Click(Area As Integer)
    Dim StartDate As Date
    Dim EndDate As Date
    GetIntervalsFullData val(DCIntervals.BoundText), StartDate, EndDate
    DTStartdate.value = StartDate
    DTEndDate.value = EndDate
    dbRecordDate.value = EndDate
    GetAllTotals StartDate, EndDate, True
End Sub

Function GetNetsalaryVouchers(NoteType As Integer, FromDate As Date, todate As Date)
    Dim strSQL  As String
    Dim DepitValue As Double
    Dim CreditValue As Double
        
    Dim Account_Code_dynamic7 As String '–„„ «·„ÊŸ›Ì‰
    Account_Code_dynamic7 = get_account_code_branch(7, my_branch)
        
    If Account_Code_dynamic7 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic7 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    –„„ «·„ÊŸ›Ì‰          ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic29 As String '«·«ÃÊ— «·„” Õﬁ… «·„ÊŸ›Ì‰
    Account_Code_dynamic29 = get_account_code_branch(29, my_branch)
        
    If Account_Code_dynamic29 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic29 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·«ÃÊ—  «·„” Õﬁ… «·„ÊŸ›Ì‰          ", vbCritical
        
            Exit Function
        End If
    End If
        
    strSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    strSQL = strSQL & " FROM         dbo.Notes INNER JOIN"
    strSQL = strSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    strSQL = strSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    strSQL = strSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
    'StrSQL = StrSQL & " AND (branch_no = " & Val(DcBranch.BoundText) & ")"
    strSQL = strSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    strSQL = strSQL & "  AND (branch_no = " & val(DcBranch.BoundText) & ")"

    strSQL = strSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    strSQL = strSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
            
    Dim RsUnitData As New ADODB.Recordset
            
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        DepitValue = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        DepitValue = 0
               
    End If

    RsUnitData.Close
       
    strSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    strSQL = strSQL & " FROM         dbo.Notes INNER JOIN"
    strSQL = strSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    strSQL = strSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    strSQL = strSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
    'StrSQL = StrSQL & " AND (branch_no = " & Val(DcBranch.BoundText) & ")"
    strSQL = strSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    strSQL = strSQL & "  AND (branch_no = " & val(DcBranch.BoundText) & ")"

    strSQL = strSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    strSQL = strSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"
              
    Dim RsUnitData1 As New ADODB.Recordset
            
    RsUnitData1.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData1.RecordCount) > 0 Then
                 
        CreditValue = IIf(IsNull(RsUnitData1("Total").value), 0, (RsUnitData1("Total").value))
    Else
        CreditValue = 0
               
    End If

    RsUnitData1.Close

    GetNetsalaryVouchers = Abs(DepitValue - CreditValue)

End Function

Function GetAllTotals(FromDate As Date, todate As Date, Optional jump As Boolean = False) As Double
    'TxtMaterialIssueVoucherTotals.text = Round(gettotal(240, fromdate, todate),2) '”‰œ«  ’—› «·„Ê«œ «·Œ«„
    Dim SalesCostBefroreProcess As Double
    Dim SalesCostAfterProcess As Double

    TxtMaterialIssueVoucherTotals.text = 0

    TxtAccDep.text = (gettotal(90, FromDate, todate))  '”‰œ«  «·«Â·«ﬂ
    '66 ﬁÌœ «·«” Õﬁ«ﬁ
    '555 ﬁÌœ «·”œ«œ

    'TxtSalaryVouchersTotals.text = Round(gettotal(66, fromdate, todate),2)
    TxtSalaryVouchersTotals.text = Round(GetNetsalaryVouchers(66, FromDate, todate), 2)
    TxtAllocations.text = Round(gettotal(8023, FromDate, todate, 0), 2)
    TxtAllocations1.text = Round(gettotal(8023, FromDate, todate, 1), 2)

    TxtExpenses.text = Round(GetExpensestotal(FromDate, todate), 2)
    'Txttotal.text = Val(TxtMaterialIssueVoucherTotals.text) +
    Txttotal.text = val(TxtAccDep.text) + val(TxtSalaryVouchersTotals.text) + val(TxtExpenses.text) + val(TxtAllocations.text) + val(TxtAllocations1.text)
    Txttotal.text = Round(Txttotal.text, 2)

    TxtTotalProductionQty.text = Round(GetÛQTY(28, FromDate, todate), SystemOptions.SysDefQuantityDecimal) 'ﬂ„Ì«  «·«‰ «Ã «· «„
    TxtTotalsalesQty.text = 0 ' Round(GetÛQTY(21, fromdate, todate), SystemOptions.SysDefQuantityDecimal) '    ﬂ„Ì«  «·„»Ì⁄« 

    If val(TxtTotalProductionQty.text) <> 0 Then
        TxtUnitValue.text = val(Txttotal.text) / val(TxtTotalProductionQty.text)
    Else
        TxtUnitValue.text = 0
    End If

    TxtUnitValue.text = Round((TxtUnitValue.text), 2)
    TxtSaLePayValue.text = 0 ' Round(GetSalesValue(fromdate, todate, 0), 2) 'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtServicesValue.text = 0 'Round(GetSalesValue(fromdate, todate, 1), 2) 'ﬁÌ„… Œœ„«  «·› —…

    If jump = True Then
        TxtSaleValue.text = 0
        TxtProfit.text = 0
        Exit Function
    End If

    'TxtSaleValue.text = Val(TxtUnitValue.text) * Val(TxtTotalsalesQty.text)
    'TxtSaleValue.text = Round(Val(TxtSaleValue.text),2)
    SalesCostBefroreProcess = Round(GetSalesCost(FromDate, todate), 2)

    IndirectCostAddition val(TxtUnitValue.text), DTStartdate.value, DTEndDate.value, 28
    SalesCostAfterProcess = Round(GetSalesCost(FromDate, todate), 2)

    SalesDifferentValue = SalesCostAfterProcess - SalesCostBefroreProcess
    TxtSaleValue.text = 0 'Round(GetSalesCost(fromdate, todate), 2) '   ﬂ·›… «·„»Ì⁄« 

    'TxtSaLePayValue.text = Round(gettotal(170, fromdate, todate),2) 'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtSaLePayValue.text = 0 ' Round(GetSalesValue(fromdate, todate, 0), 2) 'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtServicesValue.text = 0 ' Round(GetSalesValue(fromdate, todate, 1), 2) 'ﬁÌ„… Œœ„«  «·› —…

    If chkProfitService.value = vbChecked Then
        TxtProfit.text = 0 ' Val(TxtServicesValue) + Val(TxtSaLePayValue.text) - Val(TxtSaleValue.text)
    Else
        TxtProfit.text = 0 ' Val(TxtSaLePayValue.text) - Val(TxtSaleValue.text)
    End If

End Function

Function GetSalesValue(FromDate As Date, todate As Date, ItemType As Integer) As Double
    Dim strSQL  As String
    strSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
    strSQL = strSQL & " FROM         dbo.QryItemsSalesTotal(21, DEFAULT, DEFAULT, " & SQLDate(FromDate, True) & ", " & SQLDate(todate, True) & "," & ItemType & ") QryItemsSalesTotal"
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetSalesValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
    Else
        GetSalesValue = 0
               
    End If

    RsUnitData.Close

End Function

Function GetISSueVoucherForProductionValue(FromDate As Date, todate As Date, ItemType As Integer) As Double
    Dim strSQL  As String
    strSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
    strSQL = strSQL & " FROM         dbo.QryItemsSalesTotal(27, DEFAULT, DEFAULT, " & SQLDate(FromDate, True) & ", " & SQLDate(todate, True) & "," & ItemType & ") QryItemsSalesTotal"
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetISSueVoucherForProductionValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
    Else
        GetISSueVoucherForProductionValue = 0
               
    End If

    RsUnitData.Close

End Function

Function GetExpensestotal(FromDate As Date, todate As Date) As Double
    Dim strSQL  As String
  
    strSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total"
    strSQL = strSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    strSQL = strSQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    strSQL = strSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    strSQL = strSQL & " WHERE     (dbo.ExpensesType.IndirectCosts = 1) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    strSQL = strSQL & "       RecordDate >= " & SQLDate(FromDate, True)
    strSQL = strSQL & "  AND RecordDate <= " & SQLDate(todate, True)
    strSQL = strSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
    Debug.Print strSQL
    
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetExpensestotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        GetExpensestotal = 0
               
    End If

    RsUnitData.Close
End Function

Function GetÛQTY(Transaction_Type As Integer, FromDate As Date, todate As Date) As Double
    Dim strSQL  As String

    strSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity) AS TotalQty"
    strSQL = strSQL & " FROM         dbo.Transactions INNER JOIN "
    strSQL = strSQL & "dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    strSQL = strSQL & " WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
    strSQL = strSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
    strSQL = strSQL & " AND (Transaction_Type = " & Transaction_Type & ")"
    strSQL = strSQL & " AND (Transactions.BranchId = " & val(DcBranch.BoundText) & ")"
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetÛQTY = IIf(IsNull(RsUnitData("TotalQty").value), 0, (RsUnitData("TotalQty").value))
    Else
        GetÛQTY = 0
               
    End If

    RsUnitData.Close
End Function

Function gettotal(NoteType As Integer, FromDate As Date, todate As Date, Optional AllocationType As Integer = -1) As Double
    Dim strSQL  As String
        
    strSQL = "  SELECT     SUM(Note_Value) AS Total from dbo.Notes"

    strSQL = strSQL & " WHERE      NoteDate >= " & SQLDate(FromDate, True)
    strSQL = strSQL & "  AND NoteDate <= " & SQLDate(todate, True)
    strSQL = strSQL & " AND (NoteType = " & NoteType & ")"
    strSQL = strSQL & " AND (branch_no = " & val(DcBranch.BoundText) & ")"
         
    If AllocationType <> -1 Then
        strSQL = strSQL & " AND  AllocationType=" & AllocationType
    End If
            
    Dim RsUnitData As New ADODB.Recordset
            
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        gettotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        gettotal = 0
               
    End If

    RsUnitData.Close
End Function

Function GetSalesCost(FromDate As Date, todate As Date) As Double
    Dim strSQL  As String
        
    strSQL = "  SELECT     SUM(dbo.Transaction_Details.SHOWQTY * dbo.Transaction_Details.SHOWPrice) AS TotalCost"
    strSQL = strSQL & "  FROM         dbo.Transactions INNER JOIN"
    strSQL = strSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    strSQL = strSQL & " WHERE     (dbo.Transactions.Transaction_Type = 19) and    (dbo.Transactions.Transaction_Date  >= " & SQLDate(FromDate, True)
    strSQL = strSQL & "  AND (dbo.Transactions.Transaction_Date  <= " & SQLDate(todate, True)
    strSQL = strSQL & " ))"
    strSQL = strSQL & " AND (dbo.Transaction_Details.BranchId  = " & val(DcBranch.BoundText) & ") and Doctype is null"
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetSalesCost = IIf(IsNull(RsUnitData("TotalCost").value), 0, (RsUnitData("TotalCost").value))
    Else
        GetSalesCost = 0
               
    End If

    RsUnitData.Close
End Function

Private Sub DCIntervals_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetIntervalsData Me.DCIntervals
    End If

End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic
 
    Dcombos.GetBranches Me.DcBranch

    Dcombos.GetIntervalsData Me.DCIntervals
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    strSQL = "select * From TblTypicalProduction2  "
    rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Frame1.Caption = "Services Revenue"
    lbl(2).Caption = "Value"

    Me.Caption = "Typical Production Cost Calc."
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Date"
    lbl(16).Caption = "Branch"
    lbl(3).Caption = "Notes"
  
    lbl(14).Caption = "Period"
    Frame2.Caption = "Expenses"
   
    lbl(22).Caption = "Expenses And Fin. Inv."
    lbl(24).Caption = "Salaries"
    
    lbl(29).Caption = "Vacation Alloc"
    lbl(30).Caption = "End Of Service Alloc"
     
    lbl(25).Caption = "Depreciation Cost"
    lbl(26).Caption = "Materials Cost"
    lbl(15).Caption = "Total Expenses"
    Frame3.Caption = "Unit Cost"
    lbl(20).Caption = "From"
    lbl(21).Caption = "To"
    lbl(13).Caption = "total quantities produced"
    lbl(17).Caption = "Unit Cost"
    Frame4.Caption = "Sales Data"
          
    lbl(27).Caption = "Total quantities sold"
    lbl(19).Caption = "Cost of Sales"
    lbl(37).Caption = "Notes"
             
    lbl(28).Caption = "Profit"
    lbl(0).Caption = "This screen calculates the cost of items produced for the production of typical"
    lbl(12).Caption = "Value of sales"
    Frame5.Caption = "Accounting data"
                
    lbl(18).Caption = "GE no."
    Cmd(11).Caption = "Print GE"
    Frame6.Caption = "Reports"
    Cmd(7).Caption = "Summary RPT"
        
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
 
    lbl(5).Caption = "Project"

    CmdRemove.Caption = "Remove Line"

End Sub

Public Sub get_all_employee()
 
End Sub

Public Sub FillGridWithData()
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTypicalProductionId.text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    dbRecordDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)

    Me.DCIntervals.BoundText = IIf(IsNull(rs("intervalid").value), "", rs("intervalid").value)
    Me.DcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    txtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)

    TxtExpenses.text = IIf(IsNull(rs("Expenses").value), 0, rs("Expenses").value)
    TxtSalaryVouchersTotals.text = IIf(IsNull(rs("SalaryVouchersTotals").value), 0, rs("SalaryVouchersTotals").value)
    TxtAllocations.text = IIf(IsNull(rs("Allocations").value), 0, rs("Allocations").value)
    TxtAllocations1.text = IIf(IsNull(rs("Allocations1").value), 0, rs("Allocations1").value)

    TxtMaterialIssueVoucherTotals.text = IIf(IsNull(rs("MaterialIssueVoucherTotals").value), 0, rs("MaterialIssueVoucherTotals").value)
    TxtAccDep.text = IIf(IsNull(rs("AccDep").value), 0, rs("AccDep").value)
    Txttotal.text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    TxtUnitValue.text = IIf(IsNull(rs("UnitValue").value), 0, rs("UnitValue").value)

    TxtSaleValue.text = IIf(IsNull(rs("SaleValue").value), 0, rs("SaleValue").value)
    TxtTotalProductionQty.text = IIf(IsNull(rs("TotalProductionQty").value), 0, rs("TotalProductionQty").value)
    TxtTotalsalesQty.text = IIf(IsNull(rs("TotalsalesQty").value), 0, rs("TotalsalesQty").value)

    TxtNoteID.text = IIf(IsNull(rs("NoteID").value), 0, rs("NoteID").value)

    txtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
 
    TxtSaLePayValue.text = IIf(IsNull(rs("SaLePayValue").value), 0, rs("SaLePayValue").value)
    TxtServicesValue.text = IIf(IsNull(rs("SalesValue1").value), 0, rs("SalesValue1").value)
    TxtProfit.text = IIf(IsNull(rs("Profit").value), 0, rs("Profit").value)
  
    If IsNull(rs("ProfitService").value) Then
        chkProfitService.value = vbUnchecked
    Else

        If (rs("ProfitService").value) = 0 Then
            chkProfitService.value = vbUnchecked
        Else
            chkProfitService.value = vbChecked
        End If
 
    End If

    Me.TxtModFlg = "R"
    Exit Sub
ErrTrap:
End Sub
 
Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

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
