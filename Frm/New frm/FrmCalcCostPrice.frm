VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCalcCostPrice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ”«»  ﬂ·›… «·«‰ «Ã «·‰„ÿÌ ›Ì › —… „ÕœœÂ "
   ClientHeight    =   9510
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10275
   HelpContextID   =   580
   Icon            =   "FrmCalcCostPrice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   10275
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
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10245
      _cx             =   18071
      _cy             =   16695
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
      _GridInfo       =   $"FrmCalcCostPrice.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8430
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10185
         _cx             =   17965
         _cy             =   14870
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
            Height          =   8010
            Index           =   2
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   10095
            _cx             =   17806
            _cy             =   14129
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
               Top             =   7200
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
               Top             =   7200
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
               Left            =   120
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
               Picture         =   "FrmCalcCostPrice.frx":040F
               Caption         =   "Õ”«»  ﬂ·›… «·«‰ «Ã «·‰„ÿÌ ›Ì › —… „ÕœœÂ  "
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
                  ButtonImage     =   "FrmCalcCostPrice.frx":10E9
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
                  ButtonImage     =   "FrmCalcCostPrice.frx":1483
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
                  ButtonImage     =   "FrmCalcCostPrice.frx":181D
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
                  ButtonImage     =   "FrmCalcCostPrice.frx":1BB7
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
               Height          =   7035
               Index           =   1
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   120
               Width           =   10425
               _cx             =   18389
               _cy             =   12409
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
                  Top             =   5520
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
                  Height          =   3525
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   1650
                  Width           =   5685
                  Begin VB.TextBox TxtAdvancedPayments 
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
                     TabIndex        =   88
                     Top             =   1800
                     Width           =   2160
                  End
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
                     Top             =   2160
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
                     Top             =   2520
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
                     Top             =   3120
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
                     Caption         =   "«·„œ›Ê⁄«  «·„ﬁœ„Â ··„ÊŸ›Ì‰"
                     ForeColor       =   &H000000FF&
                     Height          =   405
                     Index           =   4
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   1800
                     Width           =   2640
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
                     Top             =   2160
                     Width           =   2280
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬁÌ„… ’—› «·„Ê«œ «·Œ«„ ··› —… "
                     Height          =   405
                     Index           =   26
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   2520
                     Width           =   2640
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     ForeColor       =   &H00FF0000&
                     Height          =   405
                     Index           =   15
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   3120
                     Width           =   2160
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H000000FF&
                     X1              =   120
                     X2              =   5400
                     Y1              =   3000
                     Y2              =   3000
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
                     Caption         =   "«Ã„«·Ì ﬁÌ„…   „Œ’’«  ‰Â«Ì… «·Œœ„…"
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
                  Top             =   5175
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
                     Caption         =   "‰’Ì» «·ÊÕœ… „‰ «·„’—Ê›« "
                     Height          =   315
                     Index           =   17
                     Left            =   3360
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   600
                     Width           =   2040
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "«·„»Ì⁄« "
                  Enabled         =   0   'False
                  Height          =   960
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   6090
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
                  Format          =   94240769
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
                  Format          =   94240769
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
                  Format          =   94240769
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
                  Caption         =   $"FrmCalcCostPrice.frx":1F51
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
                  Left            =   9435
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1290
                  Width           =   630
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
                  Height          =   285
                  Index           =   20
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1560
                  Width           =   465
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
         Top             =   8475
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
            ButtonImage     =   "FrmCalcCostPrice.frx":2127
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
            ButtonImage     =   "FrmCalcCostPrice.frx":24C1
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
            ButtonImage     =   "FrmCalcCostPrice.frx":285B
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
            MICON           =   "FrmCalcCostPrice.frx":2BF5
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
      ButtonImage     =   "FrmCalcCostPrice.frx":2C11
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmCalcCostPrice"
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
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
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
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
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
    Dim Vchr_result As String
    Dim notes_result As String

    If txtNoteSerial1.Text = "" Then
        Vchr_result = Voucher_coding(val(my_branch), dbRecordDate.value, 21, 102)

        If Vchr_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ﬂ«·Ì› «‰ «Ã ‰„ÿÌ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Vchr_result = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ   ﬂ«·Ì› «‰ «Ã ‰„ÿÌ  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                '  txtNoteSerial1.text = Voucher_coding(val(my_branch), dbRecordDate.value, 21, 102)
            End If
        End If
    End If
             
    If TxtNoteSerial.Text = "" Then
        notes_result = Notes_coding(val(my_branch), dbRecordDate.value)

        If notes_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If notes_result = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                ' TxtNoteSerial.text = Notes_coding(val(my_branch), dbRecordDate.value)
            End If
        End If
    End If
          
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
        TxtNoteID.Text = CStr(new_id("Notes", "NoteID", "", True))
    ElseIf Me.TxtModFlg.Text = "E" Then
     
        StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
    End If
    
    rs("id").value = TxtTypicalProductionId.Text
    
    rs("intervalid").value = IIf(Me.DCIntervals.BoundText = "", Null, Me.DCIntervals.BoundText)
    rs("branch_no").value = IIf(Me.DcBranch.BoundText = "", Null, Me.DcBranch.BoundText)
    
    rs("RecordDate").value = dbRecordDate.value
    rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)
    rs("Expenses").value = IIf(val(Me.txtExpenses.Text) = 0, 0, val(Me.txtExpenses.Text))
 
    rs("SalaryVouchersTotals").value = IIf(val(Me.TxtSalaryVouchersTotals.Text) = 0, 0, val(Me.TxtSalaryVouchersTotals.Text))
    rs("Allocations").value = IIf(val(Me.TxtAllocations.Text) = 0, 0, val(Me.TxtAllocations.Text))
    rs("Allocations1").value = IIf(val(Me.TxtAllocations1.Text) = 0, 0, val(Me.TxtAllocations1.Text))
 
    rs("MaterialIssueVoucherTotals").value = IIf(val(Me.TxtMaterialIssueVoucherTotals.Text) = 0, 0, val(Me.TxtMaterialIssueVoucherTotals.Text))
    rs("AccDep").value = IIf(val(Me.TxtAccDep.Text) = 0, 0, val(Me.TxtAccDep.Text))
    rs("total").value = IIf(val(Me.TxtTotal.Text) = 0, 0, val(Me.TxtTotal.Text))
    rs("UnitValue").value = IIf(val(Me.TxtUnitValue.Text) = 0, 0, val(Me.TxtUnitValue.Text))
  
    rs("SaleValue").value = IIf(val(Me.TxtSaleValue.Text) = 0, 0, val(Me.TxtSaleValue.Text))
    rs("TotalProductionQty").value = IIf(val(Me.TxtTotalProductionQty.Text) = 0, 0, val(Me.TxtTotalProductionQty.Text))
    rs("TotalsalesQty").value = IIf(val(Me.TxtTotalsalesQty.Text) = 0, 0, val(Me.TxtTotalsalesQty.Text))
    rs("NoteID").value = IIf(val(Me.TxtNoteID.Text) = 0, 0, val(Me.TxtNoteID.Text))
    rs("NoteSerial").value = IIf(Me.TxtNoteSerial.Text = "", "", Me.TxtNoteSerial.Text)
 
    rs("NoteSerial1").value = IIf(Me.txtNoteSerial1.Text = "", "", Me.txtNoteSerial1.Text)
 
    rs("SaLePayValue").value = IIf(val(Me.TxtSaLePayValue.Text) = 0, 0, val(Me.TxtSaLePayValue.Text))
    rs("SalesValue1").value = IIf(val(Me.TxtServicesValue.Text) = 0, 0, val(Me.TxtServicesValue.Text))
  
    If chkProfitService.value = vbChecked Then
        rs("ProfitService").value = 1
    Else
        rs("ProfitService").value = 0
    End If

    rs("Profit").value = IIf(val(Me.TxtProfit.Text) = 0, 0, val(Me.TxtProfit.Text))
 
    rs.update
 
    createVoucher
    
    Cn.CommitTrans
    BeginTrans = False
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
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

    TxtModFlg.Text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
 
    Dim Account_Code_dynamic As String
    Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«» ·„’—›«  „Ê«œ «·«‰ «Ã «· «„          ", vbCritical
        
            Exit Function
        End If
    End If
        
    Dim Account_Code_dynamic1 As String
    Account_Code_dynamic1 = get_account_code_branch(16, my_branch)
        
    If Account_Code_dynamic1 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic1 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«» ·„’—›«  «·«ÃÊ—              ", vbCritical
        
            Exit Function
        End If
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
        
    Dim Account_Code_dynamic64 As String
    Account_Code_dynamic64 = get_account_code_branch(64, my_branch)
        
    If Account_Code_dynamic64 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic64 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»      „›—œ«  ”‰ÊÌ…  ··„ÊŸ›Ì‰           ", vbCritical
        
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
        bankDes = " Õ”«»  ﬂ·›… «·«‰ «Ã «·‰„ÿÌ ⁄‰ «·› —…     " & Me.DCIntervals.Text & "   „‰ " & DTStartDate.value & "  «·Ï " & DTEnddate.value
    Else
        bankDes = " Calc Production Cost For Period  " & Me.DCIntervals.Text
  
    End If

    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
    'RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Me.TxtModFlg.Text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TxtNoteID.Text)
     
    End If

    RsNotes.AddNew
    NoteID = CStr(TxtNoteID.Text)
    RsNotes("NoteID").value = CStr(TxtNoteID.Text)
    RsNotes("NoteType").value = 102
    RsNotes("NoteDate").value = dbRecordDate.value
    RsNotes("UserID").value = user_id

    If txtNoteSerial1.Text = "" Then
        txtNoteSerial1.Text = Voucher_coding(val(my_branch), dbRecordDate.value, 21, 102)
    End If
          
    RsNotes("NoteSerial1").value = Trim$(Me.txtNoteSerial1.Text) '„”·”·   ”‰œ  ﬂ«·Ì› «·«‰ «Ã «·‰„ﬂÌ
          
    If TxtNoteSerial.Text = "" Then
        TxtNoteSerial.Text = Notes_coding(val(my_branch), dbRecordDate.value)
    End If

    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.Text) '„”·”· «·ﬁÌœ
                
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(21) '‰Ê⁄  —ﬁÌ„ ”‰œ «·«Ìœ«⁄
    RsNotes("sanad_year").value = year(dbRecordDate.value)
    RsNotes("sanad_month").value = Month(dbRecordDate.value)
    '   RsNotes("note_value_by_characters").value = WriteNo(Format(Val(Txttotal.text) + Val(TxtSaleValue.text), "0.00"), 0, True, ".")
    RsNotes("remark").value = TxtRemarks.Text & bankDes
    RsNotes("Branch_no").value = val(Me.DcBranch.BoundText)
                
    RsNotes.update
                
    line_no = 1
 
    Dim RsDev  As ADODB.Recordset
    Set RsDev = New ADODB.Recordset
'    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    '«·ÿ—› «·„œÌ‰     «·„Œ“Ê‰
    AccountCode = ModAccounts.GetMyAccountCode("TblStore", "StoreID", ProductionStoreId)

    If val(TxtTotal.Text) > 0 Then
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(Me.TxtTotal.Text)
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    '
    '”‰œ«  «·—« »
    Dim I  As Integer
    Dim LngDevID  As Long
    
    Dim SQLSalaryExpenses    As String
    Dim RsSalaryV As ADODB.Recordset
    Set RsSalaryV = New ADODB.Recordset
    
    SQLSalaryExpenses = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit ,dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    SQLSalaryExpenses = SQLSalaryExpenses & " FROM         dbo.Notes INNER JOIN"
    SQLSalaryExpenses = SQLSalaryExpenses & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    SQLSalaryExpenses = SQLSalaryExpenses & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    SQLSalaryExpenses = SQLSalaryExpenses & " WHERE     (dbo.Notes.NoteType = 66) AND (dbo.Notes.NoteDate >= " & SQLDate(Me.DTStartDate.value, True) & ") AND (dbo.Notes.NoteDate <=" & SQLDate(Me.DTEnddate.value, True) & ") AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    'SQLSalaryExpenses = SQLSalaryExpenses & " AND (branch_no = " & val(DcBranch.BoundText) & ")"
    SQLSalaryExpenses = SQLSalaryExpenses & "   ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit DESC"
 
    Dim Credit_Or_Debit As Integer
  
    RsSalaryV.Open SQLSalaryExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For I = 1 To RsSalaryV.RecordCount

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
    Next I
 
    Set RsSalaryV = Nothing

    '      «·„’—Ê›«  Ê «·›Ê« Ì— «·„«·Ì…
 
    Dim SQLExpenses As String
    Dim RsExpenseV As ADODB.Recordset
    Set RsExpenseV = New ADODB.Recordset
  
    '  SQLExpenses = "SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.ACCOUNTS.Account_Code"
    'SQLExpenses = SQLExpenses & "  FROM         dbo.ACCOUNTS INNER JOIN"
    'SQLExpenses = SQLExpenses & "   dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    'SQLExpenses = SQLExpenses & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    '  SQLExpenses = SQLExpenses & " WHERE     (dbo.ExpensesType.TypicalProduction = 1) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    '  SQLExpenses = SQLExpenses & "       RecordDate >= " & SQLDate(Me.DTStartdate.value, True)
    '  SQLExpenses = SQLExpenses & "  AND RecordDate <= " & SQLDate(Me.DTEndDate.value, True)
    '  SQLExpenses = SQLExpenses & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & Val(Dcbranch.BoundText) & ")"
    '
                     
    'SQLExpenses = SQLExpenses & "  GROUP BY dbo.ACCOUNTS.Account_Code"
   
    '18 12 2012
    SQLExpenses = "SELECT     SUM(Case"
    SQLExpenses = SQLExpenses & "     When Credit_Or_Debit=0 Then Value*1"
    SQLExpenses = SQLExpenses & " When Credit_Or_Debit=1 Then Value*-1"
    SQLExpenses = SQLExpenses & " Else  0"
    SQLExpenses = SQLExpenses & " End) AS Total, dbo.ACCOUNTS.Account_Code"
    SQLExpenses = SQLExpenses & "  FROM         dbo.ACCOUNTS INNER JOIN"
    SQLExpenses = SQLExpenses & "   dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    SQLExpenses = SQLExpenses & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    SQLExpenses = SQLExpenses & " WHERE     (dbo.ExpensesType.TypicalProduction = 1)  AND"
    SQLExpenses = SQLExpenses & "       RecordDate >= " & SQLDate(Me.DTStartDate.value, True)
    SQLExpenses = SQLExpenses & "  AND RecordDate <= " & SQLDate(Me.DTEnddate.value, True)
    'SQLExpenses = SQLExpenses & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
                     
    SQLExpenses = SQLExpenses & "  GROUP BY dbo.ACCOUNTS.Account_Code"
   
    RsExpenseV.Open SQLExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For I = 1 To RsExpenseV.RecordCount

        If Not IsNull(RsExpenseV("Account_Code").value) And (RsExpenseV("Total").value) > 0 Then
               
            AccountCode = (RsExpenseV("Account_Code").value)
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(RsExpenseV("Total").value), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsExpenseV("Total").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        ElseIf Not IsNull(RsExpenseV("Account_Code").value) And (RsExpenseV("Total").value) < 0 Then
               
            AccountCode = (RsExpenseV("Account_Code").value)
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Abs(val(RsExpenseV("Total").value)), 0, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsExpenseV("Total").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
        
        End If

        RsExpenseV.MoveNext
    Next I
 
    Set RsExpenseV = Nothing
 
    '„’«—Ì› ’—› „Ê«œ «‰ «Ã „Ê«œ
    If val(TxtMaterialIssueVoucherTotals) > 0 Then
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, val(TxtMaterialIssueVoucherTotals), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtMaterialIssueVoucherTotals), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
            
    End If
 
    '„’«—Ì›       «·«ÃÊ—
    'If Val(TxtSalaryVouchersTotals) > 0 Then
    '   line_no = line_no + 1
    '
    '  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
    '            If ModAccounts.AddNewDev(LngDevID, line_no, _
    '                Account_Code_dynamic1, Val(TxtSalaryVouchersTotals), 1, _
    '                 bankDes, Val(NoteID), , , _
    '                SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , Val(TxtSalaryVouchersTotals), , , , bankDes, , , , , , , , , , Val(Me.DcBranch.BoundText)) = False Then
    '                    GoTo ErrTrap
    '
    '            End If
    'End If
            
    ' „’«—Ì› «·«Â·«ﬂ
    SQLExpenses = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS total"
    SQLExpenses = SQLExpenses & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    SQLExpenses = SQLExpenses & " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
    SQLExpenses = SQLExpenses & "  WHERE      (dbo.Notes.NoteType = 90) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    SQLExpenses = SQLExpenses & "       RecordDate >= " & SQLDate(Me.DTStartDate.value, True)
    SQLExpenses = SQLExpenses & "  AND RecordDate <= " & SQLDate(Me.DTEnddate.value, True)
    'SQLExpenses = SQLExpenses & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
                     
    SQLExpenses = SQLExpenses & "  GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
  
    Set RsExpenseV = New ADODB.Recordset
    RsExpenseV.Open SQLExpenses, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For I = 1 To RsExpenseV.RecordCount

        If Not IsNull(RsExpenseV("Account_Code").value) And (RsExpenseV("Total").value) > 0 Then
               
            AccountCode = (RsExpenseV("Account_Code").value)
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(RsExpenseV("Total").value), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(RsExpenseV("Total").value), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        End If

        RsExpenseV.MoveNext
    Next I
    
    ' ﬁÌœ ' ﬂ·›… «·„»Ì⁄« 
    '„œÌ‰
    If val(Me.TxtSaleValue.Text) > 0 Then
        AccountCode = Account_Code_dynamic2
        line_no = line_no + 1
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(Me.TxtSaleValue.Text)
        RsDev("Credit_Or_Debit").value = 0
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = " ﬁÌœ  ﬂ·›… «·„»Ì⁄«  " & bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = " Sales Cost Vchr" & bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    'œ«∆‰
    Dim SQLCostPerInv As String
    If val(Me.TxtSaleValue.Text) > 0 Then
        
        '*********************************************************************
        SQLCostPerInv = "SELECT     SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS TotalCost, dbo.Transactions.StoreID"
SQLCostPerInv = SQLCostPerInv & " FROM         dbo.Transactions INNER JOIN"
SQLCostPerInv = SQLCostPerInv & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
SQLCostPerInv = SQLCostPerInv & " WHERE     (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Date >= " & SQLDate(Me.DTStartDate.value, True) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DTEnddate.value, True) & ")"
SQLCostPerInv = SQLCostPerInv & "                      AND (dbo.Transactions.Doctype IS NULL OR"
SQLCostPerInv = SQLCostPerInv & "                      dbo.Transactions.Doctype IN"
SQLCostPerInv = SQLCostPerInv & "                          (SELECT     id"
SQLCostPerInv = SQLCostPerInv & "                             From dbo.TblDoCumentsTypes"
SQLCostPerInv = SQLCostPerInv & "                             WHERE     (WorkWithProducction = 1)))"
SQLCostPerInv = SQLCostPerInv & "GROUP BY dbo.Transactions.StoreID"

  
    Dim RsCostPerInv As New ADODB.Recordset
     
    Dim Costvalue As Double
    Dim CostStoreId As Integer
    RsCostPerInv.Open SQLCostPerInv, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For I = 1 To RsCostPerInv.RecordCount
Costvalue = IIf(IsNull(RsCostPerInv("TotalCost").value), 0, RsCostPerInv("TotalCost").value)
CostStoreId = IIf(IsNull(RsCostPerInv("StoreId").value), 0, RsCostPerInv("StoreId").value)

        If Costvalue > 0 Then
               
            AccountCode = ModAccounts.GetMyAccountCode("TblStore", "StoreID", CLng(CostStoreId))
            line_no = line_no + 1
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(Costvalue), 1, " ﬁÌœ  ﬂ·›… «·„»Ì⁄«  " & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(Costvalue), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
                GoTo ErrTrap
                    
            End If
         
        End If

        RsCostPerInv.MoveNext
    Next I
    
        '**********************************************************************
        'AccountCode = ModAccounts.GetMyAccountCode("TblStore", "StoreID", ProductionStoreId)
        'line_no = line_no + 1
        '
        'RsDev.AddNew
        'RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        'RsDev("branch_id").value = val(Me.DcBranch.BoundText)
        'RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        'RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        'RsDev("Account_Code").value = AccountCode
        'RsDev("Value").value = val(Me.TxtSaleValue.Text)
        'RsDev("Credit_Or_Debit").value = 1
        '
        'RsDev("RecordDate").value = Me.dbRecordDate.value
        'RsDev("Notes_ID").value = val(Me.TxtNoteID.Text)   '(XPTxtID.text)
        'RsDev("Double_Entry_Vouchers_Description").value = " ﬁÌœ  ﬂ·›… «·„»Ì⁄«  " & bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        'RsDev("Double_Entry_Vouchers_Descriptione").value = " Sales Cost Vchr" & bankDes
                        
        'RsDev("UserID").value = user_id
        'RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        '
        'RsDev.update
                    
    End If
 
    '  «·«Ã«“… „’«—Ì›       «·„Œ’’« 
    If val(TxtAllocations.Text) > 0 Then
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic3, val(TxtAllocations.Text), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtSalaryVouchersTotals), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
    End If

    '  ‰Â«Ì… «·Œœ„… „’«—Ì›       «·„Œ’’« 
    If val(TxtAllocations1.Text) > 0 Then
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic4, val(TxtAllocations1.Text), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtSalaryVouchersTotals), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
    End If

    '«·„œ›Ê⁄«  «·„ﬁœ„Â
    If val(TxtAdvancedPayments.Text) > 0 Then
        line_no = line_no + 1

        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic64, val(TxtAdvancedPayments.Text), 1, bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(TxtAdvancedPayments), , , , bankDes, , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
    End If

    updateNotesValueAndNobytext (val(NoteID))
    sql = "Update TblTypicalProduction  set  NoteSerial='" & TxtNoteSerial.Text & "',NoteSerial1='" & txtNoteSerial1.Text & "' where id=" & val(TxtTypicalProductionId)
    Cn.Execute sql
    
ErrTrap:
End Function

Function CostRecieveVoucherFromProduction(cost As Double, Fromdate As Date, todate As Date, Transaction_Type As Integer)
    'On Error GoTo ErrTrap
    Dim StrSQL As String
Cn.CommandTimeout = 0
' ﬂ·Ì› ”‰œ«  «·«‰ «Ã «· «„
    StrSQL = "update dbo.Transaction_Details  set Price=" & cost & ", CostPrice=" & cost

    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
    StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
'    StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
    Cn.Execute StrSQL
 
    StrSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
    StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
'    StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
    Cn.Execute StrSQL
 
    'If
    Dim FirstPeriodDateInthisYear  As Date
    Dim fromdateS As Variant
    Dim todateS As Variant
    
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear



 If SystemOptions.CostStarting = True Then
          '  Dim FirstPeriodDateInthisYear  As Date
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
        fromdateS = Replace(Format$(FirstPeriodDateInthisYear, "MM/DD/yyyy"), "-", "/")
Else
        fromdateS = "01/01/2000"

      
End If




 
    'fromdateS = Replace(Format$(FirstPeriodDateInthisYear, "MM/DD/yyyy"), "-", "/")
    todateS = Replace(Format$(DTEnddate.value, "MM/DD/yyyy"), "-", "/")

    Transaction_Type = 19
' ﬂ·Ì› ”‰œ«  ’—› «·„„»Ì⁄« 
    'StrSQL = "update dbo.Transaction_Details  set Price= dbo.GetItemCostPrice('01/01/1900', ' 01/01/2079 ', Item_ID) , CostPrice=dbo.GetItemCostPrice('01/01/1900', ' 01/01/2079 ', Item_ID)"
    If cost = 0 Then
        StrSQL = "update dbo.Transaction_Details  set Price= 0"

        StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
        StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
'        StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 StrSQL = StrSQL & "  and  ( Doctype is null  or Doctype in(SELECT     id FROM         dbo.TblDoCumentsTypes  WHERE     (WorkWithProducction = 1))   )  "
 

        Cn.Execute StrSQL
 
        StrSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

        StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
        StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
      '  StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 
        Cn.Execute StrSQL
 
    Else
        StrSQL = "update dbo.Transaction_Details  set Price= dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID) , CostPrice=dbo.GetItemCostPrice('" & fromdateS & "', ' " & todateS & " ', Item_ID)"

        StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
        StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
      '  StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ") and Doctype is null"
 StrSQL = StrSQL & "  and  ( Doctype is null  or Doctype in(SELECT     id FROM         dbo.TblDoCumentsTypes  WHERE     (WorkWithProducction = 1))   )  "
   
        Cn.Execute StrSQL
 
        StrSQL = "update dbo.Transaction_Details  set showPrice=Price*QtyBySmalltUnit"

        StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
        StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
        StrSQL = StrSQL & " AND (Transactions.Transaction_Type = " & Transaction_Type & ")"
    '    StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
 StrSQL = StrSQL & "  and  ( Doctype is null  or Doctype in(SELECT     id FROM         dbo.TblDoCumentsTypes  WHERE     (WorkWithProducction = 1))   )  "
 
        Cn.Execute StrSQL

    End If

    'End If
ErrTrap:
End Function

Private Sub Cmd_Click(Index As Integer)
 
    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.TxtTypicalProductionId.Text = CStr(new_id("TblTypicalProduction", "id", "", True))
        
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
              
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
 
            CuurentLogdata

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

            If val(DCIntervals.BoundText) = 0 Then
            MsgBox "Õœœ «·› —…", vbCritical
            DCIntervals.SetFocus
            SendKeys ("{F4}")
            Exit Sub
            
            End If

            GetIntervalsFullData val(DCIntervals.BoundText), StartDate, EndDate
            DTStartDate.value = StartDate
            DTEnddate.value = EndDate
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
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport

            '   ViewDataList
        Case 9
    
        Case 20
     
        Case 21
            RemoveGridRow

        Case 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            ShowGL_cc Me.TxtNoteSerial.Text, , 200
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap
    Dim ItemReport As ClsItemsReport

    If TxtTypicalProductionId.Text <> "" Then
        Set ItemReport = New ClsItemsReport
        ItemReport.TypicalProduction val(TxtTypicalProductionId.Text), Me.DCIntervals.Text, Me.DcBranch.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If Me.TxtTypicalProductionId.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtNoteSerial1.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
      
            StrSQL = "Delete From notes Where NoteID=" & val(TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            If Not rs.RecordCount < 1 Then
                CostRecieveVoucherFromProduction 0, DTStartDate.value, DTEnddate.value, 28

                DoEvents
                rs.delete
                CuurentLogdata ("D")
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub RemoveGridRow()
 
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
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
    txtNoteSerial1.Text = ""
    TxtNoteSerial.Text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    DCIntervals_Click 0
    TxtNoteSerial.Text = ""
    txtNoteSerial1.Text = ""
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub
 
Private Sub dcBranch_KeyUp(KeyCode As Integer, _
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
    DTStartDate.value = StartDate
    DTEnddate.value = EndDate
    dbRecordDate.value = EndDate
    GetAllTotals StartDate, EndDate, True
End Sub

Function GetNetsalaryVouchers(notetype As Integer, Fromdate As Date, todate As Date)
    Dim StrSQL  As String
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
        
    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & notetype & ") AND (dbo.Notes.NoteDate >=" & SQLDate(Fromdate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
    StrSQL = StrSQL & " AND (branch_no = " & val(DcBranch.BoundText) & ")"
    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    StrSQL = StrSQL & "  AND (branch_no = " & val(DcBranch.BoundText) & ")"

    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
            
    Dim RsUnitData As New ADODB.Recordset
            
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        DepitValue = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        DepitValue = 0
               
    End If

    RsUnitData.Close
       
    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & notetype & ") AND (dbo.Notes.NoteDate >=" & SQLDate(Fromdate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
    'StrSQL = StrSQL & " AND (branch_no = " & Val(DcBranch.BoundText) & ")"
    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    StrSQL = StrSQL & "  AND (branch_no = " & val(DcBranch.BoundText) & ")"

    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"
              
    Dim RsUnitData1 As New ADODB.Recordset
            
    RsUnitData1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData1.RecordCount) > 0 Then
                 
        CreditValue = IIf(IsNull(RsUnitData1("Total").value), 0, (RsUnitData1("Total").value))
    Else
        CreditValue = 0
               
    End If

    RsUnitData1.Close

    GetNetsalaryVouchers = Abs(DepitValue - CreditValue)

End Function

Function GetAllTotals(Fromdate As Date, todate As Date, Optional jump As Boolean = False) As Double
If Me.TxtModFlg.Text = "R" Then Exit Function
    'TxtMaterialIssueVoucherTotals.text = Round(gettotal(240, fromdate, todate),2) '”‰œ«  ’—› «·„Ê«œ «·Œ«„
    If Me.TxtModFlg.Text = "E" Then
     
        StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
    End If

    TxtMaterialIssueVoucherTotals.Text = Round(GetISSueVoucherForProductionValue(Fromdate, todate, 0), 2) '„Ê«œ Œ«„

    TxtAccDep.Text = Round(GetTotal(90, Fromdate, todate, , val(Me.DcBranch.BoundText)), 2)  '”‰œ«  «·«Â·«ﬂ
    '66 ﬁÌœ «·«” Õﬁ«ﬁ
    '555 ﬁÌœ «·”œ«œ

    'TxtSalaryVouchersTotals.text = Round(gettotal(66, fromdate, todate),2)
    TxtSalaryVouchersTotals.Text = Round(GetNetsalaryVouchers(66, Fromdate, todate), 2)

    TxtAllocations.Text = Round(GetTotal(8023, Fromdate, todate, 0, val(Me.DcBranch.BoundText)), 2)

    TxtAllocations1.Text = Round(GetTotal(8023, Fromdate, todate, 1, val(Me.DcBranch.BoundText)), 2)
    TxtAdvancedPayments.Text = Round(GetTotal(8027, Fromdate, todate, -1, val(Me.DcBranch.BoundText)), 2)

    txtExpenses.Text = Round(GetExpensestotal(Fromdate, todate), 2)
    TxtTotal.Text = val(TxtMaterialIssueVoucherTotals.Text) + val(TxtAccDep.Text) + val(TxtSalaryVouchersTotals.Text) + val(txtExpenses.Text) + val(TxtAllocations.Text) + val(TxtAllocations1.Text) + val(TxtAdvancedPayments.Text)
    TxtTotal.Text = Round(TxtTotal.Text, 2)

    TxtTotalProductionQty.Text = Round(GetÛQTY(28, Fromdate, todate), SystemOptions.SysDefQuantityDecimal) 'ﬂ„Ì«  «·«‰ «Ã «· «„
    TxtTotalsalesQty.Text = Round(GetÛQTY(21, Fromdate, todate), SystemOptions.SysDefQuantityDecimal) '    ﬂ„Ì«  «·„»Ì⁄« 

    If val(TxtTotalProductionQty.Text) <> 0 Then
        TxtUnitValue.Text = val(TxtTotal.Text) / val(TxtTotalProductionQty.Text)
    Else
        TxtUnitValue.Text = 0
    End If

    'TxtUnitValue.text = Round((TxtUnitValue.text))
    TxtSaLePayValue.Text = Round(GetSalesValue(Fromdate, todate, 0), 2)   'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtServicesValue.Text = Round(GetSalesValue(Fromdate, todate, 1), 2) 'ﬁÌ„… Œœ„«  «·› —…

    If jump = True Then
        TxtSaleValue.Text = 0
        TxtProfit.Text = 0
        Exit Function
    End If

    'TxtSaleValue.text = Val(TxtUnitValue.text) * Val(TxtTotalsalesQty.text)
    'TxtSaleValue.text = Round(Val(TxtSaleValue.text),2)
    CostRecieveVoucherFromProduction val(TxtUnitValue.Text), DTStartDate.value, DTEnddate.value, 28

    TxtSaleValue.Text = Round(GetSalesCost(Fromdate, todate), 2) '   ﬂ·›… «·„»Ì⁄« 

    'TxtSaLePayValue.text = Round(gettotal(170, fromdate, todate),2) 'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtSaLePayValue.Text = Round(GetSalesValue(Fromdate, todate, 0), 2) 'ﬁÌ„… „»Ì⁄«  «·› —…
    TxtServicesValue.Text = Round(GetSalesValue(Fromdate, todate, 1), 2) 'ﬁÌ„… Œœ„«  «·› —…

    If chkProfitService.value = vbChecked Then
        TxtProfit.Text = val(TxtServicesValue) + val(TxtSaLePayValue.Text) - val(TxtSaleValue.Text)
    Else
        TxtProfit.Text = val(TxtSaLePayValue.Text) - val(TxtSaleValue.Text)
    End If

End Function

Function GetExpensestotalold(Fromdate As Date, todate As Date) As Double
    Dim StrSQL  As String
    'until 18 12 2012
  
    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.ExpensesType.TypicalProduction = 1) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    StrSQL = StrSQL & "       RecordDate >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND RecordDate <= " & SQLDate(todate, True)
    StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
    Debug.Print StrSQL
    
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetExpensestotalold = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        GetExpensestotalold = 0
               
    End If

    RsUnitData.Close
End Function

Function GetÛQTY(Transaction_Type As Integer, Fromdate As Date, todate As Date) As Double
    Dim StrSQL  As String

    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity) AS TotalQty"
    StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN "
    StrSQL = StrSQL & "dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & " WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True)
    StrSQL = StrSQL & " AND (Transaction_Type = " & Transaction_Type & ")"
'    StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(DcBranch.BoundText) & ")"
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetÛQTY = IIf(IsNull(RsUnitData("TotalQty").value), 0, (RsUnitData("TotalQty").value))
    Else
        GetÛQTY = 0
               
    End If

    RsUnitData.Close
End Function

Private Sub dcproject_Click(Area As Integer)

End Sub

Private Sub Dcterm_Click(Area As Integer)
 
End Sub

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

    ScreenNameArabic = "  Õ”«»  ﬂ·›… «·«‰ «Ã «·‰„ÿÌ ›Ì › —… „ÕœœÂ  "
    ScreenNameEnglish = " Cal  Cost For Typical Production Per Interval   "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 102

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
    StrSQL = "select * From TblTypicalProduction  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

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
    Lbl(4).Caption = "Employee Adv. Payments"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Frame1.Caption = "Services Revenue"
    Lbl(2).Caption = "Value"

    Me.Caption = "Typical Production Cost Calc."
    Ele(5).Caption = Me.Caption
    Lbl(7).Caption = "ID"
    Lbl(5).Caption = "Date"
    Lbl(16).Caption = "Branch"
    Lbl(3).Caption = "Notes"
  
    Lbl(14).Caption = "Period"
    Frame2.Caption = "Expenses"
   
    Lbl(22).Caption = "Expenses And Fin. Inv."
    Lbl(24).Caption = "Salaries"
    
    Lbl(29).Caption = "Vacation Alloc"
    Lbl(30).Caption = "End Of Service Alloc"
     
    Lbl(25).Caption = "Depreciation Cost"
    Lbl(26).Caption = "Materials Cost"
    Lbl(15).Caption = "Total Expenses"
    Frame3.Caption = "Unit Cost"
    Lbl(20).Caption = "From"
    Lbl(21).Caption = "To"
    Lbl(13).Caption = "total quantities produced"
    Lbl(17).Caption = "Unit Cost"
    Frame4.Caption = "Sales Data"
          
    Lbl(27).Caption = "Total quantities sold"
    Lbl(19).Caption = "Cost of Sales"
    Lbl(37).Caption = "Notes"
             
    Lbl(28).Caption = "Profit"
    Lbl(0).Caption = "This screen calculates the cost of items produced for the production of typical"
    Lbl(12).Caption = "Value of sales"
    Frame5.Caption = "Accounting data"
                
    Lbl(18).Caption = "GE no."
    Cmd(11).Caption = "Print GE"
    Frame6.Caption = "Reports"
    Cmd(7).Caption = "Summary RPT"
        
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
 
    Lbl(5).Caption = "Project"

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

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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
    Dim StrSQL As String
    Dim I As Integer

    'On Error GoTo ErrTrap
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTypicalProductionId.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    dbRecordDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)

    Me.DCIntervals.BoundText = IIf(IsNull(rs("intervalid").value), "", rs("intervalid").value)
    Me.DcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)

    txtExpenses.Text = IIf(IsNull(rs("Expenses").value), 0, rs("Expenses").value)
    TxtSalaryVouchersTotals.Text = IIf(IsNull(rs("SalaryVouchersTotals").value), 0, rs("SalaryVouchersTotals").value)
    TxtAllocations.Text = IIf(IsNull(rs("Allocations").value), 0, rs("Allocations").value)
    TxtAllocations1.Text = IIf(IsNull(rs("Allocations1").value), 0, rs("Allocations1").value)

    TxtMaterialIssueVoucherTotals.Text = IIf(IsNull(rs("MaterialIssueVoucherTotals").value), 0, rs("MaterialIssueVoucherTotals").value)
    TxtAccDep.Text = IIf(IsNull(rs("AccDep").value), 0, rs("AccDep").value)
    TxtTotal.Text = IIf(IsNull(rs("total").value), 0, rs("total").value)
    TxtUnitValue.Text = IIf(IsNull(rs("UnitValue").value), 0, rs("UnitValue").value)

    TxtSaleValue.Text = IIf(IsNull(rs("SaleValue").value), 0, rs("SaleValue").value)
    TxtTotalProductionQty.Text = IIf(IsNull(rs("TotalProductionQty").value), 0, rs("TotalProductionQty").value)
    TxtTotalsalesQty.Text = IIf(IsNull(rs("TotalsalesQty").value), 0, rs("TotalsalesQty").value)

    TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), 0, rs("NoteID").value)

    TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    txtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
 
    TxtSaLePayValue.Text = IIf(IsNull(rs("SaLePayValue").value), 0, rs("SaLePayValue").value)
    TxtServicesValue.Text = IIf(IsNull(rs("SalesValue1").value), 0, rs("SalesValue1").value)
    TxtProfit.Text = IIf(IsNull(rs("Profit").value), 0, rs("Profit").value)
  
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
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic _
       & CHR(13) & "  „”·”· " & txtNoteSerial1.Text _
       & CHR(13) & " «· «—ÌŒ   " & dbRecordDate.value _
       & CHR(13) & " «·›—⁄   " & DcBranch.Text _
       & CHR(13) & " „·«ÕŸ«   " & TxtRemarks _
       & CHR(13) & " «·› —…  " & DCIntervals.Text & " „‰   " & DTStartDate.value & " «·Ï    " & DTEnddate.value _
       & CHR(13) & "  «Ã„«·Ì «·„’—Ê›«  Ê «·›Ê« Ì— «·„«·Ì…   " & txtExpenses _
       & CHR(13) & " «Ã„«·Ì ﬁÌ„… ”‰œ«  «·—« » ··› —…   " & TxtSalaryVouchersTotals _
       & CHR(13) & "    «Ã„«·Ì ﬁÌ„…    „Œ’’«  «·«Ã«“…  " & TxtAllocations _
       & CHR(13) & "    «Ã„«·Ì ﬁÌ„…     „ﬂ«›√… ‰Â«Ì… «·Œœ„…  " & TxtAllocations1 _
       & CHR(13) & "  «·„œ›Ê⁄«  «·„ﬁœ„Â ··„ÊŸ›Ì‰   " & TxtAdvancedPayments _
       & CHR(13) & "    «Ã„«·Ì ﬁÌ„… ”‰œ«  «·«Â·«ﬂ ··› —… " & TxtAccDep _
       & CHR(13) & "  ﬁÌ„… ’—› «·„Ê«œ «·Œ«„ ··› —… ··› —…   " & TxtMaterialIssueVoucherTotals _
       & CHR(13) & "   «Ã„«·Ì    ﬂ·›… «·«‰ «Ã ⁄‰ «·› —…  " & TxtTotal _
       & CHR(13) & "  «Ã„«·Ì «·ﬂ„Ì«  «·„‰ Ã… Œ·«· «·› —…    " & TxtTotalProductionQty _
       & CHR(13) & "   ‰’Ì» «·ÊÕœ… „‰ «·„’—Ê›«    " & TxtUnitValue _
       & CHR(13) & "    «Ã„«·Ì    «·ﬂ„Ì… «·„»«⁄Â ⁄‰ «·› —…  " & TxtTotalsalesQty _
       & CHR(13) & "     ﬂ·›… «·„»Ì⁄«  Œ·«· «·› —…  " & TxtSaleValue

    If chkProfitService.value = vbChecked Then
        LogTextA = LogTextA & CHR(13) & "  Õ ÊÌ ⁄·Ï «Ì—«œ«  «·Œœ„«  "
        LogTextA = LogTextA & CHR(13) & " «Ì—«œ«  «·Œœ„«  " & TxtServicesValue
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… „»Ì⁄«  «·› —… " & TxtSaLePayValue
    LogTextA = LogTextA & CHR(13) & " «—»«Õ «·› —… " & TxtProfit
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "  NO: " & txtNoteSerial1.Text & CHR(13) & " Date   " & dbRecordDate.value & CHR(13) & " Branch   " & DcBranch.Text & CHR(13) & " Remarks  " & TxtRemarks & CHR(13) & " Period  " & DCIntervals.Text & " From    " & DTStartDate.value & " To    " & DTEnddate.value & CHR(13) & "  Expensens and Fin. Inv. Total  " & txtExpenses & CHR(13) & " Salaries Vchrs.   " & TxtSalaryVouchersTotals & CHR(13) & "    Vacations Alloc  " & TxtAllocations & CHR(13) & "   End-of-service bonus  " & TxtAllocations1 & CHR(13) & "  Adv . Payments For Employees   " & TxtAdvancedPayments & CHR(13) & "   Total Cost Of Assets Depreciation " & TxtAccDep & CHR(13) & " Total Cost Of material   " & TxtMaterialIssueVoucherTotals & CHR(13) & "   Total Production Cost  " & TxtTotal & CHR(13) & " Total Production Qty    " & TxtTotalProductionQty & CHR(13) & "   Cost Per Unit   " & TxtUnitValue & CHR(13) & "    Sales Qty  " & TxtTotalsalesQty & CHR(13) & "    Sales Cost  " & TxtSaleValue

    If chkProfitService.value = vbChecked Then
        LogTextE = LogTextE & CHR(13) & " Vchr Contain Sales Revenue "
        LogTextE = LogTextE & CHR(13) & " Sales Revenue " & TxtServicesValue
    End If

    LogTextE = LogTextE & CHR(13) & "  Sales Value " & TxtSaLePayValue
    LogTextE = LogTextE & CHR(13) & " Profit " & TxtProfit
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 102, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, , , Me.TxtNoteSerial, Me.txtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 102, Date, Time, LogTextA, LogTextE, Me.Name, "D", , , Me.TxtNoteSerial, Me.txtNoteSerial1
    End If
    
End Function
 
Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, , 102
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
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

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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
