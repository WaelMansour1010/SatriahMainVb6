VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmBankDeposite 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·«Ìœ«⁄«  «·»‰ﬂÌÂ  "
   ClientHeight    =   9375
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11100
   HelpContextID   =   580
   Icon            =   "FrmBankDeposite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   11100
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11085
      _cx             =   19553
      _cy             =   16484
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
      _GridInfo       =   $"FrmBankDeposite.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8310
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11025
         _cx             =   19447
         _cy             =   14658
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
            Height          =   7890
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   10935
            _cx             =   19288
            _cy             =   13917
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   0
               Width           =   10995
               _cx             =   19394
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
               Picture         =   "FrmBankDeposite.frx":040F
               Caption         =   "«·«Ìœ«⁄«  «·»‰ﬂÌÂ    "
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
               Begin VB.TextBox oldtxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   36
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
                  ButtonImage     =   "FrmBankDeposite.frx":10E9
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
                  TabIndex        =   37
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
                  ButtonImage     =   "FrmBankDeposite.frx":1483
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
                  TabIndex        =   38
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
                  ButtonImage     =   "FrmBankDeposite.frx":181D
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
                  TabIndex        =   39
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
                  ButtonImage     =   "FrmBankDeposite.frx":1BB7
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
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
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
               Begin VB.TextBox txtEmpCode 
                  Alignment       =   1  'Right Justify
                  Height          =   270
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   1575
                  Width           =   705
               End
               Begin VB.CheckBox Check17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  Height          =   270
                  Left            =   8640
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   4680
                  Width           =   1080
               End
               Begin VB.CheckBox chkDue 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ «·‘Ìﬂ«  «·„” Õﬁ…  ›ﬁÿ"
                  Height          =   195
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   4080
                  Width           =   2895
               End
               Begin VB.TextBox TxtTotalChequesView 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   7200
                  Width           =   1695
               End
               Begin VB.TextBox TxtTotalCashView 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   4080
                  Width           =   1695
               End
               Begin VB.TextBox TxtBankName 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10920
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.TextBox TxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   840
                  Width           =   1425
               End
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   7200
                  Width           =   1785
               End
               Begin VB.TextBox TxtTotalCheques 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   7200
                  Width           =   1575
               End
               Begin VB.TextBox TxtTotalCash 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   4080
                  Width           =   1575
               End
               Begin VB.TextBox txtchequeno 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.TextBox TxtValue1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.TextBox TxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„⁄·Ê„« "
                  Height          =   2115
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   1050
                  Width           =   4575
                  Begin MSDataListLib.DataCombo xxx 
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   58
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
                     TabIndex        =   59
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
                  Begin VB.Label lblTotalLate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "0"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   1440
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
                     TabIndex        =   61
                     Top             =   1155
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
                     TabIndex        =   60
                     Top             =   840
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
                     TabIndex        =   57
                     Top             =   480
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
                     TabIndex        =   56
                     Top             =   240
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
                     TabIndex        =   55
                     Top             =   1440
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
                     TabIndex        =   54
                     Top             =   1150
                     Width           =   1200
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
                     TabIndex        =   53
                     Top             =   840
                     Width           =   1200
                  End
               End
               Begin VB.TextBox txtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   645
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   47
                  Top             =   1260
                  Width           =   3600
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   465
                  Left            =   12720
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2220
                  Width           =   2310
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ Ì«— ’‰›"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   13560
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   2790
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   12720
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   2790
                  Width           =   1695
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   4530
                  Width           =   1590
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
                  Left            =   11280
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Text            =   "0"
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtlBanksDepositeId 
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
                  Left            =   8280
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   1200
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   4650
                  Width           =   2310
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
                  TabIndex        =   10
                  Top             =   12090
                  Width           =   2175
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
                  Left            =   5835
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   1515
                  Left            =   15
                  TabIndex        =   7
                  Top             =   2520
                  Width           =   10905
                  _cx             =   19235
                  _cy             =   2672
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
                  FormatString    =   $"FrmBankDeposite.frx":1F51
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
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin MSComCtl2.DTPicker dbRecordDate 
                  Height          =   285
                  Left            =   5160
                  TabIndex        =   12
                  Top             =   810
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   249888769
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   13680
                  TabIndex        =   13
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
                  TabIndex        =   15
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
                  Left            =   13920
                  TabIndex        =   30
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
                  Left            =   12480
                  TabIndex        =   42
                  Top             =   2100
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   926
                  _Version        =   393216
                  Format          =   249888769
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   20
                  Left            =   12480
                  TabIndex        =   48
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
                  ButtonImage     =   "FrmBankDeposite.frx":222C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   11280
                  TabIndex        =   49
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
                  ButtonImage     =   "FrmBankDeposite.frx":25C6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCEmp 
                  Height          =   315
                  Left            =   13560
                  TabIndex        =   50
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
                  Left            =   12360
                  TabIndex        =   51
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
                  Height          =   2115
                  Left            =   0
                  TabIndex        =   64
                  Top             =   5040
                  Width           =   10905
                  _cx             =   19235
                  _cy             =   3731
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
                  Cols            =   30
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBankDeposite.frx":2B60
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
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin MSDataListLib.DataCombo Dcbank 
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   66
                  Top             =   1170
                  Width           =   4125
                  _ExtentX        =   7276
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   67
                  Top             =   2160
                  Width           =   4125
                  _ExtentX        =   7276
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo Dcbank1 
                  Height          =   315
                  Left            =   2880
                  TabIndex        =   72
                  Top             =   5040
                  Visible         =   0   'False
                  Width           =   2325
                  _ExtentX        =   4101
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   7
                  Left            =   1425
                  TabIndex        =   82
                  Top             =   2160
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
                  ButtonImage     =   "FrmBankDeposite.frx":2FBD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   8
                  Left            =   240
                  TabIndex        =   83
                  Top             =   2160
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmBankDeposite.frx":3357
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   9
                  Left            =   1905
                  TabIndex        =   84
                  Top             =   4395
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
                  ButtonImage     =   "FrmBankDeposite.frx":38F1
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   10
                  Left            =   0
                  TabIndex        =   85
                  Top             =   4395
                  Visible         =   0   'False
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmBankDeposite.frx":3C8B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   88
                  Top             =   840
                  Width           =   3645
                  _ExtentX        =   6429
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCChequeBox 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   91
                  Top             =   4440
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboEmpName 
                  Height          =   315
                  Left            =   4680
                  TabIndex        =   111
                  Top             =   1560
                  Width           =   3330
                  _ExtentX        =   5874
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ÊŸ›"
                  Height          =   240
                  Index           =   64
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1560
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·‘Ìﬂ«  «·„Õœœ…"
                  Height          =   285
                  Index           =   21
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   7200
                  Width           =   1455
               End
               Begin VB.Label TxtPaymentCounts 
                  Alignment       =   1  'Right Justify
                  Caption         =   "0"
                  Height          =   375
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   7200
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   285
                  Index           =   19
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   7200
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ Õ«›Ÿ… «·‘Ìﬂ« "
                  Height          =   285
                  Index           =   18
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   4470
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   285
                  Index           =   17
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   870
                  Width           =   735
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‘Ìﬂ« "
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   7200
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·‰ﬁœ"
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   4080
                  Width           =   1095
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   255
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   5160
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„Â"
                  Height          =   255
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   5040
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»‰ﬂ"
                  Height          =   285
                  Index           =   16
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   5070
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—’Ìœ"
                  Height          =   255
                  Left            =   3480
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»‰ﬂ «·„Êœ⁄ »Â"
                  Height          =   285
                  Index           =   15
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1200
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰ «·Œ“Ì‰Â"
                  Height          =   285
                  Index           =   14
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   2190
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‘Ìﬂ« "
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   13
                  Left            =   8025
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   4080
                  Width           =   1680
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìœ«⁄«  ‰ﬁœÌ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   12
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1800
                  Width           =   2040
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   3
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1260
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
                  TabIndex        =   41
                  Top             =   2100
                  Width           =   360
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‰œÊ»"
                  Height          =   315
                  Index           =   0
                  Left            =   12525
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   1740
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   285
                  Index           =   5
                  Left            =   6645
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   930
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   270
                  Index           =   8
                  Left            =   13200
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   3480
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   240
                  Index           =   7
                  Left            =   8985
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   930
                  Width           =   825
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
                  TabIndex        =   8
                  Top             =   1170
                  Width           =   855
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
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8355
         Width           =   11025
         _cx             =   19447
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
            TabIndex        =   17
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
            ButtonImage     =   "FrmBankDeposite.frx":4225
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   18
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
            ButtonImage     =   "FrmBankDeposite.frx":45BF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   19
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
            ButtonImage     =   "FrmBankDeposite.frx":4959
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7140
            TabIndex        =   23
            Top             =   30
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
            TabIndex        =   24
            Top             =   30
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
            TabIndex        =   25
            Top             =   30
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
            TabIndex        =   26
            Top             =   30
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
            TabIndex        =   27
            Top             =   30
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
            TabIndex        =   28
            Top             =   30
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
            TabIndex        =   29
            Top             =   30
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
            Left            =   11040
            TabIndex        =   34
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
            MICON           =   "FrmBankDeposite.frx":4CF3
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
            Index           =   11
            Left            =   5400
            TabIndex        =   94
            Top             =   480
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
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   12
            Left            =   1440
            TabIndex        =   106
            Top             =   30
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·”‰œ"
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
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   375
            Left            =   4080
            TabIndex        =   109
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«·„—›ﬁ« "
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   6840
            TabIndex        =   113
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   22
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   20
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   37
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   600
            Width           =   615
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   600
            Width           =   825
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
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
      ButtonImage     =   "FrmBankDeposite.frx":4D0F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   27
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   9360
      Width           =   7155
   End
End
Attribute VB_Name = "FrmBankDeposite"
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
Dim ReturnAcc As String
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
 Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer

    With Grid

        For i = .FixedRows To .rows - 1
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 19) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 19) = vbWhite
            End If

        Next i

    End With

 

    With GRID1

        For i = .FixedRows To .rows - 1
        
            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 29) = &HFFFFC0
            Else
                .cell(flexcpBackColor, i, 1, i, 29) = vbWhite
            End If

        Next i

    End With

     
    
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & TxtNoteSerial1.text & CHR(13) & "   «· «—ÌŒ " & dbRecordDate & CHR(13) & "   «·›—⁄ " & dcBranch & CHR(13) & "   «·»‰ﬂ «·„Êœ⁄ »Â  " & Dcbank & CHR(13) & "   „·«ÕŸ«  " & txtRemarks & CHR(13) & "   —ﬁ„ «·ﬁÌœ " & TxtNoteSerial & CHR(13) & "   «Ã„«·Ì «·‰ﬁœ " & TxtTotalCashView & CHR(13) & "   «Ã„«·Ì «·‘Ìﬂ«  " & TxtTotalChequesView
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Serial " & TxtNoteSerial1.text & CHR(13) & "   Date " & dbRecordDate & CHR(13) & "   Branch " & dcBranch & CHR(13) & "Deposite Bank" & Dcbank & CHR(13) & "   Remarks " & txtRemarks & CHR(13) & " Ge NO" & TxtNoteSerial & CHR(13) & "  Total Cash " & TxtTotalCashView & CHR(13) & "  Total Cheques " & TxtTotalChequesView
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
 
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
 
 
End Function

Function Create_dev()
 
End Function

Function Create_dev1()
 
End Function

Private Sub ALLButton2_Click()
    'dcbank.text = ""

    dcproject.text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
 
End Sub

Private Sub CboPayMentType_Change()
 
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .rows = 2
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

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

'     On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.Dcbank.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ..!!"
            Else
                Msg = "Specify Bank.!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dcbank.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
  my_branch = val(dcBranch.BoundText)
    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), dbRecordDate.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), dbRecordDate.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), dbRecordDate.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1.text = "" Then
        If Voucher_coding(val(my_branch), dbRecordDate.value, 17, 20) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «Ìœ«⁄  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Voucher_coding(val(my_branch), dbRecordDate.value, 17, 20) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«Ìœ«⁄  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), dbRecordDate.value, 17, 20)
            End If
        End If
    End If
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
        TXTNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
    ElseIf Me.TxtModFlg.text = "E" Then
                 
        Cn.Execute "delete TblBanksDepositedetails where TblBanksDepositeId=" & val(Me.TxtlBanksDepositeId.text)
        StrSQL = "Delete notes where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
    End If
    
    rs("id").value = TxtlBanksDepositeId.text
     
    rs("branch_no").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
    
    rs("bankid").value = IIf(Me.Dcbank.BoundText = "", Null, Me.Dcbank.BoundText)
    rs("RecordDate").value = dbRecordDate.value
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
  rs("Emp_ID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)

    rs("NoteID").value = CStr(TXTNoteID.text)
    rs("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
    rs("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    If chkDue.value = vbChecked Then
   rs("chkDue").value = 1
    Else
    rs("chkDue").value = 0
    End If
    ''// 17 05 2015
      rs("UserID").value = val(Me.DCboUserName.BoundText)
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
   ' RsDev.Open "TblBanksDepositeDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
       StrSQL = "SELECT     *  from dbo.TblBanksDepositeDetails Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
    Dim i As Integer

    With Me.Grid

        For i = 1 To .rows - 1
 
            If .TextMatrix(i, .ColIndex("BoxID")) <> "" Then
         
                RsDev.AddNew
                RsDev("TblBanksDepositeId").value = Me.TxtlBanksDepositeId.text
                RsDev("box_or_bank").value = 0
                RsDev("BoxID").value = val(.TextMatrix(i, .ColIndex("BoxID")))
                RsDev("value").value = val(.TextMatrix(i, .ColIndex("value")))
                RsDev("Remarks").value = .TextMatrix(i, .ColIndex("Remarks"))
                RsDev("box_or_bank").value = 0
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    RsDev.Close
   ' RsDev.Open "TblBanksDepositeDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     * from dbo.TblBanksDepositeDetails Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    With Me.GRID1

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("BoxID")) <> "" And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         
                RsDev.AddNew
                RsDev("TblBanksDepositeId").value = Me.TxtlBanksDepositeId.text
                RsDev("box_or_bank").value = 0
                 RsDev("Cheqid").value = val(.TextMatrix(i, .ColIndex("id")))
                RsDev("Bankname").value = (.TextMatrix(i, .ColIndex("Bankname")))
                RsDev("BankID").value = val(.TextMatrix(i, .ColIndex("BankID")))
                RsDev("BoxID").value = val(.TextMatrix(i, .ColIndex("BoxID")))
                RsDev("value").value = val(.TextMatrix(i, .ColIndex("value")))
                RsDev("ChequeNo").value = .TextMatrix(i, .ColIndex("ChequeNo"))
                RsDev("Remarks").value = .TextMatrix(i, .ColIndex("Remarks"))
                RsDev("DueDate").value = .TextMatrix(i, .ColIndex("DueDate"))
                RsDev("NoteID").value = .TextMatrix(i, .ColIndex("NoteID"))
                 
                 RsDev("CreditAccount").value = .TextMatrix(i, .ColIndex("CreditAccount"))
                  RsDev("Returntransaction").value = val(.TextMatrix(i, .ColIndex("Returntransaction")))
                  
                
                RsDev("box_or_bank").value = 1
                RsDev.update
                Cn.Execute "update  TblChecqueBoxContent set Deposited=1,BankID=" & val(Me.Dcbank.BoundText) & " where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
             
            Else
                Cn.Execute "update  TblChecqueBoxContent set Deposited=0,BankID=Null  where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
            End If

        Next i

    End With

    createVoucher
    Cn.CommitTrans
    BeginTrans = False
 
    CuurentLogdata

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·»‰ﬂ" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = "Saved" & CHR(13)
                Msg = Msg + "Do you want enter another One"
            End If
   
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

            '   Retrive
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If

            lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
            '  Fg_Journal.Enabled = False
    End Select

    Retrive
    TxtModFlg.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function createVoucher()
    Dim bankDes As String
    Dim AccountCode As String
    Dim AccountCode1 As String
     Dim DebitBranch As Integer
        Dim CreditBranch As Integer
        
        
        Dim DebitAccount As String
        Dim CreditAccount As String
        
    Dim NoteID As String
    Dim sql As String
  Dim bankDes1 As String
    If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = " «Ìœ«⁄«  ‰ﬁœÌ… ›Ì     " & Me.Dcbank.text
    Else
        bankDes = "  Cash  Deposite  " & Me.Dcbank.text
  
    End If
bankDes1 = bankDes
    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
  '  RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (1 = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    If Me.TxtModFlg.text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TXTNoteID.text)
    End If

    RsNotes.AddNew
    NoteID = CStr(TXTNoteID.text)
    RsNotes("NoteID").value = CStr(TXTNoteID.text)
    RsNotes("NoteType").value = 20
    RsNotes("NoteDate").value = dbRecordDate.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(17) '‰Ê⁄  —ﬁÌ„ ”‰œ «·«Ìœ«⁄
    RsNotes("sanad_year").value = year(dbRecordDate.value)
    RsNotes("sanad_month").value = Month(dbRecordDate.value)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtTotalCash.text) + val(TxtTotalCheques.text), "0.00"), 0, True, ".")
    RsNotes("remark").value = txtRemarks.text & bankDes
    RsNotes("Branch_no").value = val(Me.dcBranch.BoundText)
                
    RsNotes.update
                
    line_no = 1
Dim i As Integer
  With Grid
                    For i = .FixedRows To .rows - 1
            
                        If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
                           
                            If SystemOptions.UserInterface = ArabicInterface Then
                                bankDes = bankDes & " „‰  " & .TextMatrix(i, .ColIndex("BoxName")) & CHR(13)
                            Else
                                bankDes = bankDes & " From  " & .TextMatrix(i, .ColIndex("BoxName")) & CHR(13)
                            End If
                     
                    End If
            Next i
End With

    If Grid.rows > 1 Then
        Dim RsDev  As ADODB.Recordset
        Set RsDev = New ADODB.Recordset
     '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
 

        '«·ÿ—› «·„œÌ‰      «·»‰ﬂ
        AccountCode = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
                                
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = val(Me.TxtTotalCash.text)
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.dbRecordDate.value
        RsDev("Notes_ID").value = val(Me.TXTNoteID.text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = txtRemarks.text & CHR(13) & bankDes  'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    '‰ﬁœÌ
    '  ‰ﬁœÌ «·ÿ—› «·œ«∆‰
    If Grid.rows > 1 Then
 
         
        Dim LngDevID  As Long

        With Grid
 
            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
               
                    AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(.TextMatrix(i, .ColIndex("BoxId"))))
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("Value"))), 1, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes1, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i

        End With
    
    End If
  
    '  ‘Ìﬂ«   «·ÿ—› «·„œÌ‰
    If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = " «Ìœ«⁄«  ‘Ìﬂ«  ›Ì   " & Me.Dcbank.text
    Else
        bankDes = "  Chquee  Deposite  " & Me.Dcbank.text
  
    End If

  With GRID1
                    For i = .FixedRows To .rows - 1
            
                        If .TextMatrix(i, .ColIndex("BoxId")) <> "" And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                           
                            If SystemOptions.UserInterface = ArabicInterface Then
                                bankDes = bankDes & "    ‘Ìﬂ —ﬁ„  " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " ⁄·Ï »‰ﬂ " & (.TextMatrix(i, .ColIndex("BankName"))) & " „‰" & .TextMatrix(i, .ColIndex("Remarks")) & CHR(13)
                            Else
                                bankDes = bankDes & "  Cheque NO: " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " Bank Name:   " & (.TextMatrix(i, .ColIndex("BankName"))) & " From" & .TextMatrix(i, .ColIndex("Remarks")) & CHR(13)
                            End If
                     
                    End If
            Next i
End With
    If GRID1.rows > 1 And checkSelectCheque = True Then
        
        AccountCode = ModAccounts.GetMyAccountCodeRefined("BanksData", "BankId", val(Me.Dcbank.BoundText), "Account_code1")
        line_no = line_no + 1
  
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

        If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(Me.TxtTotalCheques.text), 0, txtRemarks.text & CHR(13) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , val(Me.TxtTotalCheques.text), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
                    
        End If
        
        
         
    End If
    
    With GRID1
  
        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("BoxId")) <> "" And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    bankDes = "   «Ìœ«⁄ »‰ﬂÌ ›Ì »‰ﬂ " & Me.Dcbank.text & "  „‰ ‘Ìﬂ —ﬁ„  " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " ⁄·Ï »‰ﬂ " & (.TextMatrix(i, .ColIndex("BankName")))
                Else
                    bankDes = " bank  Deposite Bank Name" & Me.Dcbank.text & "  Cheque NO: " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " Bank Name:   " & (.TextMatrix(i, .ColIndex("BankName")))
                End If
         
                AccountCode = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxId", val(.TextMatrix(i, .ColIndex("BoxId"))), "Account_code1")
                
             If val(.TextMatrix(i, .ColIndex("Returntransaction"))) <> 0 Then        'Õ«·Â «‰ «·‘Ìﬂ ﬂ«‰ „— œ ”«»ﬁ«
             
             AccountCode = ReturnAcc
             
             
             
             End If
                line_no = line_no + 1
  
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("Value"))), 1, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , val(.TextMatrix(i, .ColIndex("branch_id")))) = False Then
                    GoTo ErrTrap
                    
                End If
                
                '''''''''''''''''''''''''''''''''''''''ÃÊ«—Ì
   
        
        
              DebitBranch = val(dcBranch.BoundText)
              
                    DebitAccount = getBranchCurrentAccount(DebitBranch)
                    
                 
                  
                  CreditBranch = val(.TextMatrix(i, .ColIndex("branch_id")))
                          CreditAccount = getBranchCurrentAccount(CreditBranch)
                 If DebitBranch = CreditBranch Then GoTo NOSusAcc
                          line_no = line_no + 1
                          
                              If ModAccounts.AddNewDev(LngDevID, line_no, DebitAccount, val(.TextMatrix(i, .ColIndex("Value"))), 0, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , CreditBranch) = False Then
                                                                GoTo ErrTrap
                                                                
                           End If
                                                            

                    
                                line_no = line_no + 1
                          
                              If ModAccounts.AddNewDev(LngDevID, line_no, CreditAccount, val(.TextMatrix(i, .ColIndex("Value"))), 1, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , DebitBranch) = False Then
                                                                GoTo ErrTrap
                                                                
                           End If
            '''''''''''''''''''''''''''''''''''''''ÃÊ«—Ì
NOSusAcc:
         
            End If

        Next i

    End With

'„Ê÷Ê⁄ «· √ﬂœ „‰ «·œ›⁄«  «·„ﬁœ„…
Dim CusID As Long
If SystemOptions.CustomerhavethreeAccounts = True And 1 = 0 Then
    With GRID1
  
        For i = .FixedRows To .rows - 1

                     If .TextMatrix(i, .ColIndex("BoxId")) <> "" And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                                    If CheckNoteAdvancedPayments(val(.TextMatrix(i, .ColIndex("NoteID"))), CusID) = True Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                bankDes = "   «Ìœ«⁄ »‰ﬂÌ ›Ì »‰ﬂ " & Me.Dcbank.text & "  „‰ ‘Ìﬂ —ﬁ„  " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " ⁄·Ï »‰ﬂ " & (.TextMatrix(i, .ColIndex("BankName")))
                                                            Else
                                                                bankDes = " bank  Deposite Bank Name" & Me.Dcbank.text & "  Cheque NO: " & (.TextMatrix(i, .ColIndex("ChequeNo"))) & " Bank Name:   " & (.TextMatrix(i, .ColIndex("BankName")))
                                                            End If
                                                     
                     '                                       AccountCode = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxId", val(.TextMatrix(i, .ColIndex("BoxId"))), "Account_code1")
                      'œ›⁄«  „ﬁœ„…
                            AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", CusID, "Account_code2")
                                       
                              AccountCode1 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", CusID, "Account_code1")
                                            
                                             
                                                            line_no = line_no + 1
                                              
                                                            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                                            
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, val(.TextMatrix(i, .ColIndex("Value"))), 0, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                                                GoTo ErrTrap
                                                                
                                                            End If
                                              
                                                        line_no = line_no + 1
                                              
                                                            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                                            
                                                            If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode1, val(.TextMatrix(i, .ColIndex("Value"))), 1, txtRemarks.text & CHR(13) & .TextMatrix(i, .ColIndex("Remarks")) & bankDes, val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.dbRecordDate.value, user_id, , , , .TextMatrix(i, .ColIndex("Value")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                                                GoTo ErrTrap
                                                                
                                                            End If
                                                            
                                                 End If
                    'ÃÊ«—Ì
           
                           
                           
                    End If
        Next i

    End With
End If

    updateNotesValueAndNobytext (val(Me.TXTNoteID.text))

ErrTrap:
End Function

Function checkSelectCheque() As Boolean
    checkSelectCheque = False
    Dim i As Integer

    With Me.GRID1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("BoxId")) <> "" And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
              
                checkSelectCheque = True
                Exit Function
            End If

        Next i

    End With

End Function

Private Sub Check17_Click()
 

    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.GRID1
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
            Next i

        End With

    Else

        With Me.GRID1

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If

     



        ReLineGrid
End Sub

Private Sub Cmd_Click(Index As Integer)
'     On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Me.TxtlBanksDepositeId.text = CStr(new_id("TblBanksDeposite", "id", "", True))
       
            'dbRecordDate.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
            Grid.Enabled = True
         
            GRID1.Clear flexClearScrollable, flexClearEverything
            GRID1.rows = 1
            GRID1.Enabled = True
            Me.dcBranch.BoundText = Current_branch
            Me.DCboUserName.BoundText = user_id
         
        Case 1
        If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Because This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
           
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 
            TxtModFlg.text = "E"
            'Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
         
            ' Grid1.Rows = Grid1.Rows + 1
            GRID1.Enabled = True
            CuurentLogdata
Me.DCboUserName.BoundText = user_id
        Case 2
        If ChekClodePeriod(dbRecordDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = val(Me.dcBranch.BoundText)
         
            If 1 = 1 Then
ReturnAcc = get_account_code_branch(126, my_branch)
 
        If ReturnAcc = "NO branch" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            Else
                                MsgBox "Branch Not Created ", vbCritical
                            End If
                
                            GoTo ErrTrap
        ElseIf ReturnAcc = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‘Ìﬂ«  „— œ… ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
  End If
           


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

            Load FrmBankDepositeSearch
    
            FrmBankDepositeSearch.show vbModal

        Case 6
            Unload Me

        Case 7
    
            If val(TxtValue.text) < 0 Then
    
                MsgBox "—’Ìœ «·Œ“Ì‰… œ«∆‰ ·« Ì„ﬂ‰ «·«Ìœ«⁄ ÊÌ„ﬂ‰ﬂ ﬂ «»… «·„»·€ «·„—«œ «Ìœ«⁄Â ÌœÊÌ«", vbInformation
                TxtValue.text = 0
                Exit Sub
            End If

            addrow

        Case 8
            RemoveGridRow
    
            '   ViewDataList
        Case 9
            addrow1

        Case 10
            RemoveGridRow1

        Case 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200, Me.TXTNoteID
        
        Case 12

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport (TxtlBanksDepositeId.text)
    End Select

    Exit Sub
ErrTrap:

End Sub

Function PrintReport(ID As Integer)

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
    MySQL = "SELECT     TOP 100 PERCENT dbo.TblBanksDepositeDetails.TblBanksDepositeId, dbo.TblBanksDepositeDetails.box_or_bank, dbo.TblBanksDepositeDetails.[value], "
    MySQL = MySQL & "                   dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.Remarks, dbo.TblBanksDepositeDetails.BoxID, dbo.TblBoxesData.BoxName,"
    MySQL = MySQL & "                   dbo.TblBoxesData.BoxNameE, dbo.TblBanksDepositeDetails.BankName, dbo.TblBanksDepositeDetails.DueDate, dbo.TblBanksDeposite.NoteSerial1,"
    MySQL = MySQL & "                   dbo.TblBanksDeposite.NoteSerial, dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.bankid, dbo.BanksData.BankName AS DepositeBankName,"
    MySQL = MySQL & "                   dbo.TblBanksDeposite.id, dbo.TblBanksDeposite.Remarks AS Remarkss, dbo.TblBanksDeposite.Emp_id, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode,"
    MySQL = MySQL & "                  dbo.TblEmployee.Emp_Namee4 , dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1"
    MySQL = MySQL & "  FROM         dbo.TblBanksDepositeDetails INNER JOIN"
    MySQL = MySQL & "                   dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID INNER JOIN"
    MySQL = MySQL & "                   dbo.TblBanksDeposite ON dbo.TblBanksDepositeDetails.TblBanksDepositeId = dbo.TblBanksDeposite.id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee ON dbo.TblBanksDeposite.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.BanksData ON dbo.TblBanksDeposite.bankid = dbo.BanksData.BankID"
    MySQL = MySQL & "  Where   dbo.TblBanksDeposite.ID=" & ID
    MySQL = MySQL & "  ORDER BY dbo.TblBanksDeposite.NoteSerial1"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "BankDeposite.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "BankDeposite.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim i As Integer
    On Error GoTo ErrTrap

    'check Cheque Not Payed

    With Me.GRID1

        For i = 1 To .rows - 1
                 
             If .TextMatrix(i, .ColIndex("NoteID")) <> "" Then
                If ChequeBoxCollect(val(.TextMatrix(i, .ColIndex("NoteID"))), val(.TextMatrix(i, .ColIndex("Returntransaction")))) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   Õ’Ì·  "
                    Msg = Msg & CHR(13) & " ··‘Ìﬂ —ﬁ„ " & .TextMatrix(i, .ColIndex("ChequeNo"))
                    Msg = Msg & CHR(13) & "»ﬁÌ„… " & .TextMatrix(i, .ColIndex("Value"))
                    Msg = Msg & CHR(13) & " ⁄·Ï »‰ﬂ " & .TextMatrix(i, .ColIndex("BankName"))
                                          
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                          
                    Exit Sub
                End If
                                        
            End If
                                  
        Next i

    End With
 
    If TxtlBanksDepositeId.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
      
            StrSQL = "Delete From notes Where NoteID=" & val(TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

            StrSQL = "Delete From TblBanksDepositeDetails  Where TblBanksDepositeId=" & val(TxtlBanksDepositeId.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            With Me.GRID1

                For i = 1 To .rows - 1

                    If .TextMatrix(i, .ColIndex("BoxID")) <> "" Then
          
                        Cn.Execute "update  TblChecqueBoxContent set Deposited=0 where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
              
                    End If

                Next i

            End With
 
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                
                    Grid.Clear flexClearScrollable, flexClearEverything
                    Grid.rows = 1
          
                    GRID1.Clear flexClearScrollable, flexClearEverything
                    GRID1.rows = 1
               
                    TxtModFlg_Change
           
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
    
        With Me.Grid
  .RemoveItem .Row
  ReLineGrid
            If Me.TxtModFlg = "E" Then Exit Sub
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
            LogTextA = "  Õ–› «·Œ“Ì‰…   " & .cell(flexcpTextDisplay, .Row, .ColIndex("BoxName")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
            LogTexte = "  Delete  Box   " & .cell(flexcpTextDisplay, .Row, .ColIndex("BoxName")) & " With Value " & .cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtNoteSerial), TxtNoteSerial1
        End With
  
      
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()

    With Me.GRID1

        If .Row <= 0 Then Exit Sub
    
        Cn.Execute "update  TblChecqueBoxContent set Deposited=0 where NoteID=" & val(.TextMatrix(.Row, .ColIndex("NoteID")))
                                                        
        .RemoveItem .Row

    End With

    ReLineGrid
End Sub

Function addrow()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    On Error Resume Next

    If Trim(Me.DcboBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“Ì‰…..!!"
        Else
            Msg = "Specify Box.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DcboBox.SetFocus
        Sendkeys "{F4}"
        Exit Function
    End If
 
    If val(TxtValue.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„Â  ’ÕÌÕÂ"
        Else
            MsgBox "Enter Correct Value"
        End If

        TxtValue.SetFocus
        Exit Function
    End If
 
    Me.Grid.rows = Me.Grid.rows + 1
    LngRow = Me.Grid.rows - 1
 
    With Me.Grid
  
        .TextMatrix(LngRow, .ColIndex("BoxId")) = val(DcboBox.BoundText)
    
        .TextMatrix(LngRow, .ColIndex("BoxName")) = DcboBox.text
    
        .TextMatrix(LngRow, .ColIndex("Value")) = val(TxtValue.text)
    
        .TextMatrix(LngRow, .ColIndex("Remarks")) = ""
     
        If Me.TxtModFlg = "E" Then
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                           
            LogTextA = "  «Ìœ«⁄ „‰ «·Œ“Ì‰…  " & DcboBox & " »ﬁÌ„… " & val(TxtValue.text)
            LogTexte = "Deposite From Box  " & DcboBox & " With Value " & val(TxtValue.text)
                    
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtNoteSerial, TxtNoteSerial1
        End If
                                                     
        .AutoSize 0, .Cols - 1, False
    End With
 
    Me.TxtValue.text = ""
    DcboBox.BoundText = ""
    ReLineGrid

End Function

Function addrow1()
Dim branchname As String
Dim branchnamee As String
Dim branch_id As Double
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
    On Error Resume Next

    Dim rs As New ADODB.Recordset
    Dim i As Integer
    StrSQL = "select * from TblChecqueBoxContent where (Deposited=0  or  CustomerReturn=1 )  and    ChequeBoxID= " & val(DCChequeBox.BoundText)
   If chkDue.value = vbChecked Then
           StrSQL = StrSQL + " and (DueDate <=" & SQLDate(dbRecordDate.value, True) & ")"
  End If
   
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          
    GRID1.Clear flexClearScrollable, flexClearEverything
    GRID1.rows = 1

    For i = 1 To rs.RecordCount
        Me.GRID1.rows = Me.GRID1.rows + 1
        LngRow = Me.GRID1.rows - 1
   
        With Me.GRID1


    If Check17.value = vbChecked Then

  .TextMatrix(LngRow, .ColIndex("Select")) = True
End If

.TextMatrix(LngRow, .ColIndex("Returntransaction")) = IIf(IsNull(rs("Returntransaction").value), 0, rs("Returntransaction").value)

            .TextMatrix(LngRow, .ColIndex("BoxID")) = val(DCChequeBox.BoundText)
     .TextMatrix(LngRow, .ColIndex("id")) = IIf(IsNull(rs("id").value), 0, rs("id").value)
            .TextMatrix(LngRow, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), 0, rs("NoteID").value)
            .TextMatrix(LngRow, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
    
            .TextMatrix(LngRow, .ColIndex("Value")) = IIf(IsNull(rs("ChequeValue").value), "", rs("ChequeValue").value)
    
            .TextMatrix(LngRow, .ColIndex("ChequeNo")) = IIf(IsNull(rs("ChequeNo").value), "", rs("ChequeNo").value)
            .TextMatrix(LngRow, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
          If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", Format(rs("DueDate").value, "yyyy/mm/dd"))
        Else
        .TextMatrix(LngRow, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", Format(rs("DueDate").value, "dd/mm/yyyy"))
        End If
        
        
        GetBranchnmeFromnotes val(.TextMatrix(LngRow, .ColIndex("NoteID"))), branch_id, branchname, branchnamee
        .TextMatrix(LngRow, .ColIndex("branch_id")) = branch_id
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(LngRow, .ColIndex("branchname")) = branchname
        Else
        .TextMatrix(LngRow, .ColIndex("branchname")) = branchnamee
        End If
        
    .cell(flexcpChecked, LngRow, .ColIndex("CustomerReturn")) = IIf(IsNull(rs("CustomerReturn").value), 0, 1)
    
    
             .TextMatrix(LngRow, .ColIndex("Returntransaction")) = IIf(IsNull(rs("Returntransaction").value), "", rs("Returntransaction").value)
              .TextMatrix(LngRow, .ColIndex("CreditAccount")) = IIf(IsNull(rs("CreditAccount").value), "", rs("CreditAccount").value)
                
    
            .AutoSize 0, .Cols - 1, False
        End With

        rs.MoveNext
    Next i
 
    'Me.TxtValue.text = ""
    'txtchequeno.text = ""
    'Dcbank1.BoundText = ""
    'TxtValue1.text = ""

    ReLineGrid

End Function

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

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments TxtNoteSerial1, "0712201404"

End Sub

Private Sub dbRecordDate_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub dcbank_Click(Area As Integer)
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
    
    If Grid.rows > 1 Then
        If Grid.rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub

Private Sub DcboBox_Change()
    Dim AccountCode As String
    Dim Balance As Double
    Dim balancetype As Integer
    Dim FirstPeriodDateInthisYear  As Date

    If val(DcboBox.BoundText) = 0 Then TxtValue.text = 0: Exit Sub

    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(DcboBox.BoundText))
    'get_balanceFromGlNew Accountcode, , , , FirstPeriodDateInthisYear, Date, , , Balance, Val(Me.DcBranch.BoundText)

    Balance = GetActualAccountBalance(AccountCode, , FirstPeriodDateInthisYear, dbRecordDate.value)
    'getBalanceWithOpeningBalance Accountcode, Val(dcBranch.BoundText), Date, balance, balanceType

    TxtValue.text = Balance
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcChequeBox_Click(Area As Integer)
'addrow1
End Sub

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

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then TxtEmpCode.text = "": Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtEmpCode.text = EmpCode
    
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 22
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
End Sub


Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmpCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub


Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = "«·«Ìœ«⁄«  «·»‰ﬂÌÂ  "
    ScreenNameEnglish = "Bank Deposite"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 20
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(12).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
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

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic
Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBanks Me.Dcbank
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.Dcbank1
    Dcombos.GetChequeBox Me.DCChequeBox
   Dcombos.GetEmployees Me.DcboEmpName

    Dcombos.GetBranches Me.dcBranch

    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblBanksDeposite  WHERE 1=1  "
  StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
        If SystemOptions.usertype <> UserAdminAll Then
  '      StrSQL = StrSQL & " where   branch_no=" & Current_branch
    End If
    StrSQL = StrSQL & "  order by noteserial1 "
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
Cmd(12).Caption = "Print"
Check17.Caption = "Select All"
CmdAttach.Caption = "Attachments"
lbl(64).Caption = "Emp"
lbl(22).Caption = "By"
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
chkDue.Caption = "Show Due only"
    Cmd(11).Caption = "JE Print"
    Label4.Caption = "Total Cash"
    Label6.Caption = "Total Cheque"
    lbl(19).Caption = "JE NO"
    lbl(21).Caption = "Cheques Sel."
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "  Banks Deposit"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = " Date"
    'Ele(3).Caption = "Select Interval"
    'lbl(2).Caption = "Year"
    lbl(17).Caption = "Branch"

    lbl(15).Caption = "Depit Bank"
    lbl(3).Caption = "Remarks"
    lbl(12).Caption = "Cash Deposite"
    lbl(14).Caption = "From Box "
    Label1.Caption = "Value"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Remove"

    lbl(13).Caption = "Cheques "
    lbl(18).Caption = "Cheques  Box"
    lbl(16).Caption = " From Bank"
    Label3.Caption = "Chq. NO"
    Label2.Caption = "Value"
    Cmd(9).Caption = "Add"
    Cmd(10).Caption = "Remove"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("BoxName")) = "BoxId"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With

    With Me.GRID1
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Select")) = "Select"

        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("ChequeNO")) = "Cheque NO"

        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"

    End With

    lbl(20).Caption = "Curr Rec."
    lbl(37).Caption = "Total Rec."
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

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
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
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
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

        .rows = .rows + 1
        .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows - 1, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows - 1, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

ErrTrap:
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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 20
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
 
        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    On Error Resume Next
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
 
    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
If .rows > 1 Then


        Me.TxtTotalCash.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
 Else
  Me.TxtTotalCash.text = 0
 End If
 
    End With
                 
    IntCounter = 0

    With Me.GRID1

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("BoxId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i

        Me.TxtTotalCheques.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
    End With

    CalCulateParts
    Coloring
End Sub

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntCount As Integer

    Dim SngTotal As Double

    With Me.GRID1
        SngTotal = 0
        IntCount = 0

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                IntCount = IntCount + 1
                SngTotal = SngTotal + val(.TextMatrix(i, .ColIndex("Value")))
            End If

        Next i

    End With

    Me.TxtPaymentCounts.Caption = IntCount
    Me.TxtTotalCheques.text = SngTotal
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)
    'Exit Sub
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
          
    GRID1.Clear flexClearScrollable, flexClearEverything
    GRID1.rows = 1
          
    TxtTotalCash.text = 0
    TxtTotalCheques.text = 0
DCChequeBox.text = ""
DcboBox.text = ""

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
 
    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If

    If Lngid <> 0 Then
        rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    If NoteID <> 0 Then
        rs.Find "NoteID=" & NoteID, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    
    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    Me.TxtlBanksDepositeId.text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    dbRecordDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)

    Dcbank.BoundText = IIf(IsNull(rs("bankid").value), "", rs("bankid").value)
DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
   
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
If Not IsNull(rs("chkDue").value) Then
chkDue.value = IIf((rs("chkDue").value) = True, vbChecked, vbUnchecked)

Else
chkDue.value = vbUnchecked
End If
''// 17 05 2015
 Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
  
    
    'StrSQL = " SELECT   * FROM         dbo.TbllBanksDepositeDetails  "
    'StrSQL = StrSQL & "  where box_or_bank=0 and  TbllBanksDepositeId=" & Val(Me.TxtlBanksDepositeId.text)
  
    StrSQL = "SELECT     dbo.TblBanksDepositeDetails.TblBanksDepositeId, dbo.TblBanksDepositeDetails.box_or_bank, dbo.TblBanksDepositeDetails.[value], "
    StrSQL = StrSQL & "  dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.Remarks, dbo.TblBanksDepositeDetails.BoxID, dbo.TblBoxesData.BoxName,"
    StrSQL = StrSQL & "  dbo.TblBoxesData.BoxNameE"
    StrSQL = StrSQL & "  FROM         dbo.TblBanksDepositeDetails INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID"
    StrSQL = StrSQL & "   WHERE     (dbo.TblBanksDepositeDetails.TblBanksDepositeId = " & val(Me.TxtlBanksDepositeId.text) & ") AND (dbo.TblBanksDepositeDetails.box_or_bank = 0)"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
  
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(RsDev("BoxID").value), 0, val(RsDev("BoxID").value))
            
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("value").value), 0, val(RsDev("value").value))
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(RsDev("BoxName").value), "", RsDev("BoxName").value)
                Else
                    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(RsDev("BoxNameE").value), "", RsDev("BoxNameE").value)
                End If
              
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    StrSQL = "SELECT    Cheqid, BoxID, TblBanksDepositeId, box_or_bank, [value], ChequeNo, Remarks, BankName,DueDate,noteid,CreditAccount,Returntransaction"
    StrSQL = StrSQL & " From dbo.TblBanksDepositeDetails"
    StrSQL = StrSQL & "   WHERE     (dbo.TblBanksDepositeDetails.TblBanksDepositeId = " & val(Me.TxtlBanksDepositeId.text) & ") AND (dbo.TblBanksDepositeDetails.box_or_bank = 1)"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.GRID1
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
  
                .TextMatrix(i, .ColIndex("Boxid")) = IIf(IsNull(RsDev("Boxid").value), 0, val(RsDev("Boxid").value))
            
                '                .TextMatrix(i, .ColIndex("bankid")) = IIf(IsNull(RsDev("bankid").value), _
                                 0, Val(RsDev("bankid").value))
            
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("value").value), 0, val(RsDev("value").value))
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsDev("BankName").value), "", RsDev("BankName").value)
                Else
                    .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsDev("BankName").value), "", RsDev("BankName").value)
                End If
              
                              .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsDev("Cheqid").value), 0, RsDev("Cheqid").value)
                .TextMatrix(i, .ColIndex("ChequeNo")) = IIf(IsNull(RsDev("ChequeNo").value), "", RsDev("ChequeNo").value)
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
                .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(RsDev("DueDate").value), "", RsDev("DueDate").value)
            
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDev("NoteID").value), 0, RsDev("NoteID").value)
            Dim branchname As String
               Dim branchnamee As String
               Dim branch_id As Double
               
        GetBranchnmeFromnotes val(.TextMatrix(i, .ColIndex("NoteID"))), branch_id, branchname, branchnamee
        .TextMatrix(i, .ColIndex("branch_id")) = branch_id
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(i, .ColIndex("branchname")) = branchname
        Else
        .TextMatrix(i, .ColIndex("branchname")) = branchnamee
        End If
            
            
                .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
            .TextMatrix(i, .ColIndex("CreditAccount")) = IIf(IsNull(RsDev("CreditAccount").value), "", RsDev("CreditAccount").value)
            .TextMatrix(i, .ColIndex("Returntransaction")) = IIf(IsNull(RsDev("Returntransaction").value), "", RsDev("Returntransaction").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    Coloring
    
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    If Col <> Grid.ColIndex("Remarks") Then
        Cancel = True
    End If

End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With GRID1

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select

        ReLineGrid
    
        If Me.TxtModFlg = "E" Then

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
            If .cell(flexcpChecked, Row, .ColIndex("Select")) = flexChecked Then
                LogTextA = "   ÕœÌœ «·‘Ìﬂ —ﬁ„   " & .cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "⁄·Ï »‰ﬂ " & .cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                LogTexte = "Select Cheque No  " & .cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " With Value " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "On Bank " & .cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                                                         
            Else
                                                          
                LogTextA = "«·€«¡    ÕœÌœ «·‘Ìﬂ —ﬁ„   " & .cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " »ﬁÌ„… " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "⁄·Ï »‰ﬂ " & .cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                LogTexte = "DeSelect Cheque No  " & .cell(flexcpTextDisplay, Row, .ColIndex("ChequeNo")) & " With Value " & .cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "On Bank " & .cell(flexcpTextDisplay, Row, .ColIndex("BankName"))
                                                         
            End If
                                                         
            AddToLogFile CInt(user_id), 20, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtNoteSerial, TxtNoteSerial1
        End If
                                                     
    End With

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)
    Dim Msg As String

    With GRID1
 
        Select Case .ColKey(Col)

            Case "Remarks"
                Cancel = False
                Exit Sub

            Case "Select"
     
                If .TextMatrix(.Row, .ColIndex("NoteID")) <> "" Then
                    If ChequeBoxCollect(val(.TextMatrix(.Row, .ColIndex("NoteID"))), val(.TextMatrix(.Row, .ColIndex("Returntransaction")))) = False Then
                        Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                        Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   Õ’Ì· "
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Cancel = True
                        Undo
                        '     .Cell(flexcpChecked, .Row, .ColIndex("Select")) = flexChecked
           
                        Exit Sub
                    End If
                End If
    
                Cancel = False
                Exit Sub
        End Select

        Cancel = True
    End With

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

        Cmd(9).Enabled = True
    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        Cmd(9).Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
        Cmd(9).Enabled = False
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtTotalCash_Change()
    TxtTotalCashView.text = Format(val(TxtTotalCash.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtTotalCheques_Change()
    TxtTotalChequesView.text = Format(val(TxtTotalCheques.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue.text, 0)
End Sub

Private Sub TxtValue1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue1.text, 0)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

 '   On Error GoTo ErrTrap

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
