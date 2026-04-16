VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmDocType 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«‰Ê«⁄ «·„” ‰œ«  "
   ClientHeight    =   8985
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10125
   HelpContextID   =   580
   Icon            =   "FrmDocType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   10125
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
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
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
      ButtonImage     =   "FrmDocType.frx":038A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8985
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10125
      _cx             =   17859
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
      _GridInfo       =   $"FrmDocType.frx":0724
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7950
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10065
         _cx             =   17754
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
            Width           =   9975
            _cx             =   17595
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
            Begin VB.Frame Frame6 
               Caption         =   " ﬁ«—Ì— «·«‰ «Ã"
               Height          =   1065
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   6360
               Width           =   3975
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   7
                  Left            =   1680
                  TabIndex        =   25
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
                  Left            =   1680
                  TabIndex        =   26
                  Top             =   600
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
            Begin VB.Frame Frame5 
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
               Height          =   1080
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   6360
               Width           =   5775
               Begin VB.TextBox txtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   240
                  Width           =   2160
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   11
                  Left            =   120
                  TabIndex        =   22
                  Top             =   600
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
                  TabIndex        =   23
                  Top             =   240
                  Width           =   720
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
               Picture         =   "FrmDocType.frx":07A9
               Caption         =   "«‰Ê«⁄ «·„” ‰œ«    "
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
                  ButtonImage     =   "FrmDocType.frx":1483
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
                  ButtonImage     =   "FrmDocType.frx":181D
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
                  ButtonImage     =   "FrmDocType.frx":1BB7
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
                  ButtonImage     =   "FrmDocType.frx":1F51
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
               Height          =   7635
               Index           =   1
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   240
               Width           =   15225
               _cx             =   26855
               _cy             =   13467
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
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌœŒ· ›Ì «·«‰ «Ã «·‰„ÿÌ"
                  Height          =   255
                  Index           =   6
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   1680
                  Width           =   3975
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
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   705
                  Visible         =   0   'False
                  Width           =   2160
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
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   12090
                  Width           =   2175
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   4650
                  Width           =   2310
               End
               Begin VB.TextBox TxtDoCumentsTypesid 
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
                  Left            =   7680
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   870
                  Width           =   1200
               End
               Begin VB.TextBox txtType 
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
                  Height          =   525
                  Left            =   10800
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Text            =   "0"
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   4530
                  Width           =   1590
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   2790
                  Width           =   1695
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ Ì«— ’‰›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   2790
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   465
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   2220
                  Width           =   2310
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
                  Height          =   510
                  Left            =   3960
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   95
                  Top             =   2130
                  Width           =   4920
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„⁄·Ê„« "
                  Height          =   2115
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   1050
                  Width           =   4575
                  Begin MSDataListLib.DataCombo DcBranch0 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   85
                     Top             =   120
                     Width           =   2565
                     _ExtentX        =   4524
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
                  Begin MSDataListLib.DataCombo DCGroup 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   86
                     Top             =   480
                     Width           =   2565
                     _ExtentX        =   4524
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
                     Height          =   315
                     Index           =   4
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   840
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «· Õ’Ì·« "
                     Height          =   195
                     Index           =   6
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   1150
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·„ √Œ—« "
                     Height          =   195
                     Index           =   9
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   1440
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ì »⁄ ›—⁄"
                     Height          =   315
                     Index           =   10
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
                     Top             =   240
                     Width           =   1200
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ì »⁄ „Ã„Ê⁄Â"
                     Height          =   315
                     Index           =   11
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   480
                     Width           =   1200
                  End
                  Begin VB.Label lblTotlSales 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   315
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   840
                     Width           =   1200
                  End
                  Begin VB.Label lblTotalRevenue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   1155
                     Width           =   1200
                  End
                  Begin VB.Label lblTotalLate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   1440
                     Width           =   1200
                  End
               End
               Begin VB.TextBox TxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Frame Frame2 
                  Caption         =   "«·„’—Ê›«  Œ·«· «·› —…"
                  Enabled         =   0   'False
                  Height          =   2700
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   4530
                  Width           =   5655
                  Begin VB.TextBox TxtExpenses 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   360
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtSalaryVouchersTotals 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   720
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtAccDep 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   1080
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtMaterialIssueVoucherTotals 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   1440
                     Width           =   2160
                  End
                  Begin VB.TextBox Txttotal 
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   1920
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
                     TabIndex        =   82
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
                     TabIndex        =   81
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
                     TabIndex        =   80
                     Top             =   1080
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
                     TabIndex        =   79
                     Top             =   1440
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
                     TabIndex        =   78
                     Top             =   1920
                     Width           =   2160
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H000000FF&
                     X1              =   120
                     X2              =   5400
                     Y1              =   1800
                     Y2              =   1800
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "”⁄— «· ﬂ·›… ··ÊÕœ… „"
                  Enabled         =   0   'False
                  Height          =   1050
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   4935
                  Width           =   5655
                  Begin VB.TextBox TxtTotalProductionQty 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   240
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtUnitValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
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
                     TabIndex        =   71
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
                     TabIndex        =   70
                     Top             =   600
                     Width           =   2040
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "«·„»Ì⁄« "
                  Enabled         =   0   'False
                  Height          =   1320
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   5970
                  Width           =   9735
                  Begin VB.TextBox TxtSaleValue 
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
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   600
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtTotalsalesQty 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   120
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtSaLePayValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   120
                     Width           =   2160
                  End
                  Begin VB.TextBox TxtProfit 
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
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   600
                     Width           =   2160
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
                     TabIndex        =   66
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
                     TabIndex        =   65
                     Top             =   240
                     Width           =   2760
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
                     TabIndex        =   64
                     Top             =   120
                     Width           =   1320
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
                     TabIndex        =   63
                     Top             =   600
                     Width           =   1320
                  End
               End
               Begin VB.TextBox txtname 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   3975
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   1335
                  Width           =   4905
               End
               Begin VB.TextBox txtnamee 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   1800
                  Width           =   4905
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’·«ÕÌ… «· ⁄«„·"
                  ForeColor       =   &H000000C0&
                  Height          =   1230
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   5910
                  Width           =   3975
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„” Œœ„"
                     Height          =   195
                     Index           =   2
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   840
                     Value           =   -1  'True
                     Width           =   1245
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ã„Ê⁄Â"
                     Height          =   195
                     Index           =   1
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   480
                     Width           =   885
                  End
                  Begin VB.OptionButton Authority 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ· «·„” Œœ„Ì‰"
                     Height          =   195
                     Index           =   0
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   240
                     Width           =   1845
                  End
                  Begin MSDataListLib.DataCombo DcUserGroup 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   54
                     Top             =   480
                     Width           =   2415
                     _ExtentX        =   4260
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
                     TabIndex        =   55
                     Top             =   840
                     Width           =   2415
                     _ExtentX        =   4260
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1605
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   2925
                  Width           =   7455
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   5
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   133
                     Top             =   1320
                     Visible         =   0   'False
                     Width           =   135
                  End
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   0
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   240
                     Width           =   135
                  End
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   1
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   600
                     Width           =   135
                  End
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   2
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   960
                     Width           =   135
                  End
                  Begin MSDataListLib.DataCombo DCAccount1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   44
                     Top             =   240
                     Width           =   5025
                     _ExtentX        =   8864
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCAccount2 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   45
                     Top             =   600
                     Width           =   5025
                     _ExtentX        =   8864
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCAccount3 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   46
                     Top             =   960
                     Width           =   5025
                     _ExtentX        =   8864
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ «·⁄„Ì· „œÌ‰"
                     Height          =   255
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   134
                     Top             =   1320
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ „œÌ‰ 1 "
                     Height          =   375
                     Left            =   5520
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ œ«∆‰ 1"
                     Height          =   375
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   600
                     Width           =   1815
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ Œ’„ "
                     Height          =   375
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   960
                     Width           =   1215
                  End
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1230
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   4515
                  Width           =   7455
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   3
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   240
                     Width           =   135
                  End
                  Begin VB.CheckBox Chk 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Check2"
                     Height          =   255
                     Index           =   4
                     Left            =   5160
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   600
                     Width           =   135
                  End
                  Begin MSDataListLib.DataCombo DCAccount4 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   36
                     Top             =   240
                     Width           =   5025
                     _ExtentX        =   8864
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCAccount5 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   37
                     Top             =   600
                     Width           =   5025
                     _ExtentX        =   8864
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ œ«∆‰2"
                     Height          =   375
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   600
                     Width           =   1575
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ „œÌ‰2"
                     Height          =   375
                     Left            =   5880
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   240
                     Width           =   1455
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3195
                  Left            =   9975
                  TabIndex        =   105
                  Top             =   2280
                  Width           =   9945
                  _cx             =   17542
                  _cy             =   5636
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDocType.frx":22EB
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
               Begin MSComCtl2.DTPicker dbRecordDate 
                  Height          =   285
                  Left            =   10080
                  TabIndex        =   106
                  Top             =   930
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   93782017
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   10440
                  TabIndex        =   107
                  Top             =   2670
                  Width           =   4365
                  _ExtentX        =   7699
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
               Begin MSDataListLib.DataCombo dcproject 
                  Height          =   315
                  Left            =   11760
                  TabIndex        =   108
                  Top             =   1980
                  Width           =   1605
                  _ExtentX        =   2831
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
               Begin MSDataListLib.DataCombo Dcterm 
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   109
                  Top             =   1050
                  Width           =   3285
                  _ExtentX        =   5794
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
               Begin MSComCtl2.DTPicker dbTodate 
                  Height          =   525
                  Left            =   10440
                  TabIndex        =   110
                  Top             =   2100
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   926
                  _Version        =   393216
                  Format          =   93782017
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   10560
                  TabIndex        =   111
                  Top             =   2790
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmDocType.frx":25DA
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   11280
                  TabIndex        =   112
                  Top             =   2790
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmDocType.frx":2974
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCEmp 
                  Height          =   315
                  Left            =   10320
                  TabIndex        =   113
                  Top             =   1740
                  Width           =   2565
                  _ExtentX        =   4524
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
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   10560
                  TabIndex        =   114
                  Top             =   2790
                  Width           =   4365
                  _ExtentX        =   7699
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
               Begin VSFlex8Ctl.VSFlexGrid Grid1 
                  Height          =   2130
                  Left            =   10920
                  TabIndex        =   115
                  Top             =   5145
                  Width           =   9945
                  _cx             =   17542
                  _cy             =   3757
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   19
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDocType.frx":2F0E
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
               Begin MSDataListLib.DataCombo DCActivity 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   116
                  Top             =   870
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTStartdate 
                  Height          =   285
                  Left            =   10440
                  TabIndex        =   117
                  Top             =   2805
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   93782017
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DTEndDate 
                  Height          =   285
                  Left            =   9960
                  TabIndex        =   118
                  Top             =   1800
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   93782017
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcBranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   119
                  Top             =   1260
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCDocType 
                  Height          =   315
                  Left            =   4080
                  TabIndex        =   120
                  Top             =   840
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
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
                  Height          =   585
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   1170
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   225
                  Index           =   7
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   930
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   270
                  Index           =   8
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   3480
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·”‰œ"
                  Height          =   285
                  Index           =   5
                  Left            =   6645
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   930
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ»"
                  Height          =   315
                  Index           =   0
                  Left            =   10485
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   1740
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   525
                  Index           =   2
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   2100
                  Width           =   360
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   300
                  Index           =   3
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   2130
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·‰‘«ÿ"
                  Height          =   165
                  Index           =   14
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   945
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·›—⁄"
                  Height          =   300
                  Index           =   16
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   1290
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   225
                  Index           =   29
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   1335
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   225
                  Index           =   30
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   1710
                  Width           =   945
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
               TabIndex        =   132
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
         Width           =   10065
         _cx             =   17754
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
            ButtonImage     =   "FrmDocType.frx":31D5
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
            ButtonImage     =   "FrmDocType.frx":356F
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
            ButtonImage     =   "FrmDocType.frx":3909
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7980
            TabIndex        =   7
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
            Left            =   7080
            TabIndex        =   8
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
            Left            =   6240
            TabIndex        =   9
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
            Left            =   5235
            TabIndex        =   10
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
            Left            =   4200
            TabIndex        =   11
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
            Left            =   240
            TabIndex        =   12
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
            Left            =   3270
            TabIndex        =   13
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
            TabIndex        =   14
            Tag             =   "Delete Row"
            Top             =   0
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
            MICON           =   "FrmDocType.frx":3CA3
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
            Height          =   495
            Index           =   9
            Left            =   2160
            TabIndex        =   15
            Top             =   510
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   873
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   10
            Left            =   1080
            TabIndex        =   136
            Top             =   510
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·ﬂ·"
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
            TabIndex        =   17
            Top             =   240
            Width           =   1515
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
            TabIndex        =   16
            Top             =   225
            Width           =   1740
         End
      End
   End
End
Attribute VB_Name = "FrmDocType"
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

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .Rows = 2
            .Clear flexClearScrollable
        End With

    End If

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

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub



Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
        If val(Me.DCDocType.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ‰Ê⁄ «·„” ‰œ  ..!!"
            Else
            Msg = "Please Select Type"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCDocType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        If val(Me.DCActivity.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·‰‘«ÿ..!!"
            Else
            Msg = "Please Select Activity"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCActivity.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        If val(Me.Dcbranch.BoundText) = 0 Then
            '    Msg = "ÌÃ» ≈Œ Ì«— «·›—⁄..!!"
            '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '    DcBranch.SetFocus
            '    SendKeys "{F4}"
            '    Exit Sub
        End If

        If Trim(TxtName.Text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ ··‰Ê⁄ ⁄—»Ì  ..!!"
            Else
            Msg = "Please Enter Name"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtName.SetFocus
             SendKeys "{F4}"
            Exit Sub
        End If
        
        
If TxtNameE.Text = "" Then TxtNameE.Text = TxtName.Text
        If Trim(TxtNameE.Text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ ··‰Ê⁄ «‰Ã·Ì“Ì  ..!!"
            Else
            Msg = "Please Enter Name English"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNameE.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
        TxtNoteID.Text = CStr(new_id("Notes", "NoteID", "", True))
    ElseIf Me.TxtModFlg.Text = "E" Then
   
    End If
    
    rs("id").value = TxtDoCumentsTypesid.Text
    
    rs("Doctype").value = IIf(Me.DCDocType.BoundText = "", Null, Me.DCDocType.BoundText)
    rs("ActivityId").value = IIf(Me.DCActivity.BoundText = "", Null, Me.DCActivity.BoundText)
     
    rs("branch_no").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
    rs("name").value = IIf(Me.TxtName.Text = "", "", Me.TxtName.Text)
    rs("namee").value = IIf(Me.TxtNameE.Text = "", "", Me.TxtNameE.Text)
    rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)

    If Authority(0).value = True Then
        rs("Authoritytype").value = 0
        rs("groupid").value = Null
        rs("userid").value = Null
            
    ElseIf Authority(1).value = True Then
        rs("Authoritytype").value = 1
        rs("groupid").value = IIf(Me.DcUserGroup.BoundText = "", Null, Me.DcUserGroup.BoundText)
        rs("userid").value = Null
    ElseIf Authority(2).value = True Then
        rs("Authoritytype").value = 3
        rs("groupid").value = Null
        rs("userid").value = IIf(Me.DCUser.BoundText = "", Null, Me.DCUser.BoundText)
            
    End If

    rs("Account_code1").value = IIf(Me.DcAccount1.BoundText = "", Null, Me.DcAccount1.BoundText)
    rs("Account_code2").value = IIf(Me.DcAccount2.BoundText = "", Null, Me.DcAccount2.BoundText)
    rs("Account_code3").value = IIf(Me.DcAccount3.BoundText = "", Null, Me.DcAccount3.BoundText)
    rs("Account_code4").value = IIf(Me.DcAccount4.BoundText = "", Null, Me.DcAccount4.BoundText)
    rs("Account_code5").value = IIf(Me.DCAccount5.BoundText = "", Null, Me.DCAccount5.BoundText)
 
   rs("UseAccount_code1").value = IIf(Me.Chk(0).value = vbChecked, 1, 0)
    rs("UseAccount_code2").value = IIf(Me.Chk(1).value = vbChecked, 1, 0)
    rs("UseAccount_code3").value = IIf(Me.Chk(2).value = vbChecked, 1, 0)
    rs("UseAccount_code4").value = IIf(Me.Chk(3).value = vbChecked, 1, 0)
    rs("UseAccount_code5").value = IIf(Me.Chk(4).value = vbChecked, 1, 0)
    rs("UseCustomerAcc").value = IIf(Me.Chk(5).value = vbChecked, 1, 0)
      rs("WorkWithProducction").value = IIf(Me.Chk(6).value = vbChecked, 1, 0)


    rs.update
   
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
            Msg = "This is record alredy Saved" & CHR(13)
            Msg = Msg + "you need to enter another record"
            End If

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
              MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            
            End If
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
         Msg = "Can not save data & Chr(13)"
        Msg = Msg + "I have been invalid input data " & CHR(13)
        Msg = Msg + "Make sure of the accuracy of the data and try again"
        
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
      Msg = "Sorry ...an error during Saving " & CHR(13)

    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Chk_Click(Index As Integer)
 
    Select Case Index

        Case 0

            If Chk(0).value = Checked Then
                DcAccount1.Enabled = True
          
            Else
                DcAccount1.Text = ""
                DcAccount1.Enabled = False
            
            End If
        
        Case 1

            If Chk(1).value = Checked Then
                DcAccount2.Enabled = True
          
            Else
                DcAccount2.Text = ""
                DcAccount2.Enabled = False
            
            End If
        
        Case 2

            If Chk(2).value = Checked Then
                DcAccount3.Enabled = True
          
            Else
                DcAccount3.Text = ""
                DcAccount3.Enabled = False
            
            End If
        
        Case 3

            If Chk(3).value = Checked Then
                DcAccount4.Enabled = True
          
            Else
                DcAccount4.Text = ""
                DcAccount4.Enabled = False
            
            End If
        
        Case 4

            If Chk(4).value = Checked Then
                DCAccount5.Enabled = True
          
            Else
                DCAccount5.Text = ""
                DCAccount5.Enabled = False
            
            End If
        
    End Select
  
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
 
            TxtModFlg.Text = "N"
            clear_all Me
            Me.TxtDoCumentsTypesid.Text = CStr(new_id("TblDoCumentsTypes", "id", "", True))
         
        Case 1
            TxtModFlg.Text = "E"
         
        Case 2
  
            SaveData
           
        Case 3
            Undo

        Case 4

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
            PrintReport

            '   ViewDataList
        Case 9
            print_report , 1
        Case 10
            print_report
        Case 20
            addrow

        Case 21
            

        Case 11
            ShowGL_cc Me.TxtNoteSerial.Text, , 200
    End Select

    Exit Sub
ErrTrap:

End Sub
Function print_report(Optional NoteSerial As String, Optional inde As Integer = 0)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblDoCumentsTypes.id, dbo.TblDoCumentsTypes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                      dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee, dbo.TblDoCumentsTypes.Remarks, dbo.TblDoCumentsTypes.UseCustomerAcc,"
MySQL = MySQL & "                       dbo.TblDoCumentsTypes.WorkWithProducction, dbo.TblDoCumentsTypes.Doctype, dbo.TransactionTypes.TransactionTypeName,"
MySQL = MySQL & "                       dbo.TransactionTypes.TransactionEnglishName, dbo.TblDoCumentsTypes.ActivityId, dbo.tblActivitesType.Name AS ActivName,"
MySQL = MySQL & "                       dbo.tblActivitesType.namee AS ActivNameE, dbo.TblDoCumentsTypes.Authoritytype, dbo.TblDoCumentsTypes.userid, dbo.TblUsers.UserName,"
MySQL = MySQL & "                       dbo.TblDoCumentsTypes.groupid, dbo.TblGroupUsers.Name AS GroupName, dbo.TblGroupUsers.Namee AS GroupNameE, dbo.TblDoCumentsTypes.Account_code1,"
MySQL = MySQL & "                       dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblDoCumentsTypes.Account_code2,"
MySQL = MySQL & "                       ACCOUNTS_1.Account_Name AS Account_Name2, ACCOUNTS_1.Account_Serial AS Account_Serial2, ACCOUNTS_1.Account_NameEng AS Account_NameE2,"
MySQL = MySQL & "                       dbo.TblDoCumentsTypes.Account_code3, ACCOUNTS_2.Account_Name AS Account_Name3, ACCOUNTS_2.Account_Serial AS Account_Serial3,"
MySQL = MySQL & "                       ACCOUNTS_2.Account_NameEng AS Account_NameE3, dbo.TblDoCumentsTypes.Account_code4, ACCOUNTS_3.Account_Name AS Account_Name4,"
MySQL = MySQL & "                       ACCOUNTS_3.Account_Serial AS Account_Serial4, ACCOUNTS_3.Account_NameEng AS Account_NameE4, dbo.TblDoCumentsTypes.Account_code5,"
MySQL = MySQL & "                       ACCOUNTS_4.Account_Name AS Account_Name5, ACCOUNTS_4.Account_Serial AS Account_Serial5, ACCOUNTS_4.Account_NameEng AS Account_NameE5,"
MySQL = MySQL & "                       dbo.TblDoCumentsTypes.UseAccount_code1, dbo.TblDoCumentsTypes.UseAccount_code2, dbo.TblDoCumentsTypes.UseAccount_code3,"
MySQL = MySQL & "                       dbo.TblDoCumentsTypes.UseAccount_code4 , dbo.TblDoCumentsTypes.UseAccount_code5"
MySQL = MySQL & "  FROM         dbo.TblDoCumentsTypes LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ACCOUNTS_4 ON dbo.TblDoCumentsTypes.Account_code5 = ACCOUNTS_4.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ACCOUNTS_3 ON dbo.TblDoCumentsTypes.Account_code4 = ACCOUNTS_3.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblDoCumentsTypes.Account_code3 = ACCOUNTS_2.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblDoCumentsTypes.Account_code2 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.TblDoCumentsTypes.Account_code1 = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblGroupUsers ON dbo.TblDoCumentsTypes.groupid = dbo.TblGroupUsers.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblUsers ON dbo.TblDoCumentsTypes.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.tblActivitesType ON dbo.TblDoCumentsTypes.ActivityId = dbo.tblActivitesType.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TransactionTypes ON dbo.TblDoCumentsTypes.Doctype = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblDoCumentsTypes.branch_no = dbo.TblBranchesData.branch_id"
If inde = 1 Then
MySQL = MySQL & "  Where (dbo.TblDoCumentsTypes.ID = " & val(TxtDoCumentsTypesid.Text) & ")"
End If

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTypeDucoment.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTypeDucomentE.rpt"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    'xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(4).AddCurrentValue SumTotalExpen(val(DcFixedAssets.BoundText))    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub PrintReport()

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If Me.TxtDoCumentsTypesid.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
Msg = "Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
  
            If Not rs.RecordCount < 1 Then
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available there is no record"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry ... an error during delete"
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



Function addrow()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

    If Option2.value = True Then
        If dcitems.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If

        wherestr = "  where ItemID= " & val(dcitems.BoundText)
    End If

    sql = "Select * from TblItems "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid
 
        lastrow = .Rows
    
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                       
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
                       
                'lllllllllllllll
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                 
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsUnit.RecordCount > 0 Then
                    RsUnit.MoveFirst
                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
               
                End If

                RsUnit.Close
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Function

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
    TxtNoteSerial1.Text = ""
    TxtNoteSerial.Text = ""

End Sub

Private Sub DcAccount1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Account_search.show
            Account_search.case_id = 86
   End If
End Sub

Private Sub DCAccount2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Account_search.show
            Account_search.case_id = 87
   End If
End Sub

Private Sub DCAccount3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Account_search.show
            Account_search.case_id = 89
   End If
End Sub

Private Sub DCAccount4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Account_search.show
            Account_search.case_id = 92
   End If
End Sub

Private Sub DCAccount5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Account_search.show
            Account_search.case_id = 91
   End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    DCIntervals_Click 0
End Sub

Private Sub DCDocType_Change()
    Label4.Visible = False
    Chk(5).Visible = False

    If val(DCDocType.BoundText) = 21 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                Frame8.Caption = "›« Ê—… „»Ì⁄«  "
                Frame9.Caption = "”‰œ «·’—›   "
            Else
               Frame8.Caption = "Sales Invoice "
                Frame9.Caption = "Issue Voucher"
            End If
        Frame9.Visible = True

    ElseIf val(DCDocType.BoundText) = 22 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Frame8.Caption = "›« Ê—… „‘ —»«  "
        Frame9.Caption = "”‰œ «·«” ·«„   "
    Else
    Frame8.Caption = "Buy Invoice"
        Frame9.Caption = "Recive Voucher"
    End If
        Frame9.Visible = True


    ElseIf val(DCDocType.BoundText) = 5 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Frame8.Caption = "„—œÊœ«  „‘ —»«  "
        Frame9.Caption = "”‰œ «·’—›   "
    Else
    Frame8.Caption = "Buy Return"
        Frame9.Caption = "Issue Voucher"
    End If
        Frame9.Visible = True


    ElseIf val(DCDocType.BoundText) = 19 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                      Frame8.Caption = "”‰œ ’—› „Œ“‰Ì"
             Else
                       Frame8.Caption = "Issue Voucher"
             End If
        Frame9.Visible = False
        DcAccount4.Text = ""
        DCAccount5.Text = ""
        DcAccount3.Text = ""

        DcAccount3.Visible = False
        Label3.Visible = False
        Chk(2).Visible = False

        Chk(2).value = vbUnchecked
        Chk(3).value = vbUnchecked
        Chk(4).value = vbUnchecked
        Label4.Visible = True
        Chk(5).Visible = True
        If SystemOptions.UserInterface = ArabicInterface Then
        Label4.Caption = "Õ «·⁄„Ì· „œÌ‰"
        Else
        Label4.Caption = "Cust Debit"
        End If
    ElseIf val(DCDocType.BoundText) = 20 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Frame8.Caption = "”‰œ «” ·«„ „Œ“‰Ì"
 Else
 Frame8.Caption = "Recieve Voucher"
 End If
        Frame9.Visible = False
        DcAccount4.Text = ""
        DCAccount5.Text = ""
        DcAccount3.Text = ""

        DcAccount3.Visible = False
        Label3.Visible = False
        Chk(2).Visible = False

        Chk(2).value = vbUnchecked

        Chk(3).value = vbUnchecked
        Chk(4).value = vbUnchecked
        Label4.Visible = True
        Chk(5).Visible = True
        If SystemOptions.UserInterface = ArabicInterface Then
        Label4.Caption = "Õ «·„Ê—œ œ«∆‰"
        Else
        Label4.Caption = "Supplier Credit"
        End If
    End If

End Sub

Private Sub DCDocType_Click(Area As Integer)
    DCDocType_Change
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub

Private Sub DCIntervals_Change()
    DCIntervals_Click 0
End Sub

Private Sub DCIntervals_Click(Area As Integer)
 
End Sub

Function GetAllTotals(Fromdate As Date, ToDate As Date) As Double
    TxtMaterialIssueVoucherTotals.Text = Round(gettotal(240, Fromdate, ToDate), 2) '”‰œ«  ’—› «·„Ê«œ «·Œ«„

    TxtAccDep.Text = Round(gettotal(90, Fromdate, ToDate), 2) '”‰œ«  «·«Â·«ﬂ
    '66 ﬁÌœ «·«” Õﬁ«ﬁ
    '555 ﬁÌœ «·”œ«œ

    TxtSalaryVouchersTotals.Text = Round(gettotal(66, Fromdate, ToDate), 2)
    txtExpenses.Text = Round(GetExpensestotal(Fromdate, ToDate), 2)
    txtTotal.Text = val(TxtMaterialIssueVoucherTotals.Text) + val(TxtAccDep.Text) + val(TxtSalaryVouchersTotals.Text) + val(txtExpenses.Text)
    txtTotal.Text = Round(txtTotal.Text, 2)

    TxtTotalProductionQty.Text = Round(GetÛQTY(28, Fromdate, ToDate), 2) 'ﬂ„Ì«  «·«‰ «Ã «· «„
    TxtTotalsalesQty.Text = Round(GetÛQTY(21, Fromdate, ToDate), 2) '    ﬂ„Ì«  «·„»Ì⁄« 

    If val(TxtTotalProductionQty.Text) <> 0 Then
        TxtUnitValue.Text = val(txtTotal.Text) / val(TxtTotalProductionQty.Text)
    Else
        TxtUnitValue.Text = 0
    End If

    TxtUnitValue.Text = Round(val(TxtUnitValue.Text), 2)

    TxtSaleValue.Text = val(TxtUnitValue.Text) * val(TxtTotalsalesQty.Text)
    TxtSaleValue.Text = Round(val(TxtSaleValue.Text), 2)
    TxtSaLePayValue.Text = Round(gettotal(170, Fromdate, ToDate), 2) '„»Ì⁄«  «·› —…
    TxtProfit.Text = val(TxtSaLePayValue.Text) - val(TxtSaleValue.Text)
End Function

Function GetExpensestotal(Fromdate As Date, ToDate As Date) As Double
    Dim StrSQL  As String
  
    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.ExpensesType.TypicalProduction = 1) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND"
    StrSQL = StrSQL & "       RecordDate >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND RecordDate <= " & SQLDate(ToDate, True)
    StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(Dcbranch.BoundText) & ")"
    Debug.Print StrSQL
    
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetExpensestotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        GetExpensestotal = 0
               
    End If

    RsUnitData.Close
End Function

Function GetÛQTY(Transaction_Type As Integer, Fromdate As Date, ToDate As Date) As Double
    Dim StrSQL  As String

    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity) AS TotalQty"
    StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN "
    StrSQL = StrSQL & "dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & " WHERE      dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(ToDate, True)
    StrSQL = StrSQL & " AND (Transaction_Type = " & Transaction_Type & ")"
    StrSQL = StrSQL & " AND (Transaction_Details.BranchId = " & val(Dcbranch.BoundText) & ")"
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetÛQTY = IIf(IsNull(RsUnitData("TotalQty").value), 0, (RsUnitData("TotalQty").value))
    Else
        GetÛQTY = 0
               
    End If

    RsUnitData.Close
End Function

Function gettotal(NoteType As Integer, Fromdate As Date, ToDate As Date) As Double
    Dim StrSQL  As String
        
    StrSQL = "  SELECT     SUM(Note_Value) AS Total from dbo.Notes"

    StrSQL = StrSQL & " WHERE      NoteDate >= " & SQLDate(Fromdate, True)
    StrSQL = StrSQL & "  AND NoteDate <= " & SQLDate(ToDate, True)
    StrSQL = StrSQL & " AND (NoteType = " & NoteType & ")"
    StrSQL = StrSQL & " AND (branch_no = " & val(Dcbranch.BoundText) & ")"
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        gettotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        gettotal = 0
               
    End If

    RsUnitData.Close
End Function

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
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

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    With Me.GRID1
        Set .WallPaper = GrdBack.Picture
     
    End With

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select id,name from tblActivitesType   "
    Else
        StrSQL = "  select id,namee from tblActivitesType   "
    End If

    fill_combo DCActivity, StrSQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetSalesRepData Me.DCEmP
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetSalesRepGroups Me.DCGroup

    'Dcombos.GetIntervalsData Me.DCIntervals

    Dcombos.GetAccountingCodes Me.DcAccount1, True
    Dcombos.GetAccountingCodes Me.DcAccount2, True
    Dcombos.GetAccountingCodes Me.DcAccount3, True
    Dcombos.GetAccountingCodes Me.DcAccount4, True
    Dcombos.GetAccountingCodes Me.DCAccount5, True
    Dcombos.GetDocTypes Me.DCDocType, 1
    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblDoCumentsTypes  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
    Cmd(9).Caption = "Print"
    Cmd(10).Caption = "Print All"
    
Chk(6).Caption = "Work with Regular Production"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Doc Types"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Doc Types"
    lbl(14).Caption = "Activity"
    lbl(16).Caption = "Branch"
    lbl(29).Caption = "Arabic Name"
    lbl(30).Caption = "English Name"
    lbl(3).Caption = "Remarks"
    Label1.Caption = "Depit Acc 1"
    Label2.Caption = "Credit Acc 1"
    Label3.Caption = "Discount Acc  "

    Label7.Caption = "Depit Acc 2"
    Label6.Caption = "Credit Acc 2"

    Frame7.Caption = "Priviliges"
    Authority(0).Caption = "All"
    Authority(1).Caption = "Group"
    Authority(2).Caption = "User"

    CmdRemove.Caption = "Remove Line"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        '.TextMatrix(0, .ColIndex("Emp_code")) = "Emp_code"
        '.TextMatrix(0, .ColIndex("Emp_Name")) = "Emp_Name"
        '.TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        '.TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
        '.TextMatrix(0, .ColIndex("work_status")) = "work_status"
        '.TextMatrix(0, .ColIndex("project_name")) = "project name"
        '.TextMatrix(0, .ColIndex("cost_center")) = "cost center"
        '.TextMatrix(0, .ColIndex("work_days")) = "work days"
        '.TextMatrix(0, .ColIndex("ATTENDANCE")) = "absence"
        '.TextMatrix(0, .ColIndex("late")) = "delay"
        '.TextMatrix(0, .ColIndex("discount")) = "discount"
        '.TextMatrix(0, .ColIndex("net_work_days")) = "net work days"
        .TextMatrix(0, .ColIndex("addition")) = "over time"
        '.TextMatrix(0, .ColIndex("remarks")) = "remarks"

    End With

End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, 2))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, 2))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, 2))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, 2))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

ErrTrap:
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
    Dim i As Integer

    'On Error GoTo ErrTrap
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtDoCumentsTypesid.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
    Me.DCDocType.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.DCActivity.BoundText = IIf(IsNull(rs("ActivityId").value), "", rs("ActivityId").value)
  
    Me.Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    TxtName.Text = IIf(IsNull(rs("name").value), "", rs("name").value)
    TxtNameE.Text = IIf(IsNull(rs("namee").value), "", rs("namee").value)

    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)

    If rs("Authoritytype").value = 0 Then
        Authority(0).value = True
           
    ElseIf rs("Authoritytype").value = 1 Then
        Me.DcUserGroup.BoundText = IIf(IsNull(rs("UserGroupid").value), "", rs("UserGroupid").value)
        Me.DCUser.BoundText = ""
        Authority(1).value = True
    ElseIf rs("Authoritytype").value = 2 Then
        Authority(2).value = True
        Me.DCUser.BoundText = IIf(IsNull(rs("Userid").value), "", rs("Userid").value)
        Me.DcUserGroup.BoundText = ""
    Else
        Authority(0).value = True
        Authority(1).value = False
        Authority(2).value = False
        Me.DcUserGroup.BoundText = ""
        Me.DCUser.BoundText = ""
    End If

    Me.DcAccount1.BoundText = IIf(IsNull(rs("Account_code1").value), "", rs("Account_code1").value)
    Me.DcAccount2.BoundText = IIf(IsNull(rs("Account_code2").value), "", rs("Account_code2").value)
    Me.DcAccount3.BoundText = IIf(IsNull(rs("Account_code3").value), "", rs("Account_code3").value)
    Me.DcAccount4.BoundText = IIf(IsNull(rs("Account_code4").value), "", rs("Account_code4").value)
    Me.DCAccount5.BoundText = IIf(IsNull(rs("Account_code5").value), "", rs("Account_code5").value)
    Me.Chk(0).value = IIf(IsNull(rs("UseAccount_code1").value) Or (rs("UseAccount_code1").value) = 0, vbUnchecked, vbChecked)
    Me.Chk(1).value = IIf(IsNull(rs("UseAccount_code2").value) Or (rs("UseAccount_code2").value) = 0, vbUnchecked, vbChecked)
    Me.Chk(2).value = IIf(IsNull(rs("UseAccount_code3").value) Or (rs("UseAccount_code3").value) = 0, vbUnchecked, vbChecked)
    Me.Chk(3).value = IIf(IsNull(rs("UseAccount_code4").value) Or (rs("UseAccount_code4").value) = 0, vbUnchecked, vbChecked)
    Me.Chk(4).value = IIf(IsNull(rs("UseAccount_code5").value) Or (rs("UseAccount_code5").value) = 0, vbUnchecked, vbChecked)
    Me.Chk(5).value = IIf(IsNull(rs("UseCustomerAcc").value) Or (rs("UseCustomerAcc").value) = 0, vbUnchecked, vbChecked)
     Me.Chk(6).value = IIf(IsNull(rs("WorkWithProducction").value) Or (rs("WorkWithProducction").value) = 0, vbUnchecked, vbChecked)

    Exit Sub
ErrTrap:
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
