VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmVendorContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "« ›«ﬁÌ«  «·„Ê—œÌ‰"
   ClientHeight    =   7545
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmVendorContract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12240
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7545
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16005
      _cx             =   28231
      _cy             =   13309
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
      _GridInfo       =   $"FrmVendorContract.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6510
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15945
         _cx             =   28125
         _cy             =   11483
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
            Height          =   6090
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   15855
            _cx             =   27966
            _cy             =   10742
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
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   0
               Width           =   12075
               _cx             =   21299
               _cy             =   1349
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   24
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
               BackColor       =   16777215
               ForeColor       =   4210688
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Picture         =   "FrmVendorContract.frx":0410
               Caption         =   "« ›«ﬁÌ«  «·„Ê—œÌ‰ "
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
                  TabIndex        =   30
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmVendorContract.frx":10EA
                  ColorButton     =   16777215
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
                  TabIndex        =   31
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmVendorContract.frx":1484
                  ColorButton     =   16777215
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
                  TabIndex        =   32
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmVendorContract.frx":181E
                  ColorButton     =   16777215
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
                  TabIndex        =   33
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmVendorContract.frx":1BB8
                  ColorButton     =   16777215
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
               Height          =   5955
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   -120
               Width           =   15225
               _cx             =   26855
               _cy             =   10504
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
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1575
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   840
                  Width           =   12015
                  Begin VB.TextBox TxtSearchCode 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   9120
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   600
                     Width           =   1335
                  End
                  Begin VB.TextBox TxtTblVendorContractD 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   9120
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   240
                     Width           =   1320
                  End
                  Begin VB.CheckBox ChkLocked 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ìﬁ«› «· ⁄«„·"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   8160
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   960
                     Width           =   2310
                  End
                  Begin VB.TextBox txtRemarks 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   675
                     Left            =   240
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   42
                     Top             =   720
                     Width           =   4320
                  End
                  Begin MSComCtl2.DTPicker dbFromDate 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   45
                     Top             =   240
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   80674817
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker dbTodate 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   46
                     Top             =   240
                     Width           =   1935
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   80674817
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DBCboClientName 
                     Height          =   315
                     Left            =   5640
                     TabIndex        =   47
                     Top             =   600
                     Width           =   3405
                     _ExtentX        =   6006
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
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   7
                     Left            =   10800
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„œ Â« „‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   5
                     Left            =   4800
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   240
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Ê—œ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   0
                     Left            =   10800
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   600
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ï"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   2
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   240
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   3
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   720
                     Width           =   840
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Height          =   735
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   2280
                  Width           =   12015
                  Begin VB.TextBox TxtItemCode 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   240
                     Width           =   1815
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
                     Left            =   8760
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1335
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
                     Left            =   10200
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   240
                     Width           =   1695
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   20
                     Left            =   1320
                     TabIndex        =   37
                     Top             =   240
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
                     ButtonImage     =   "FrmVendorContract.frx":1F52
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   390
                     Index           =   21
                     Left            =   120
                     TabIndex        =   38
                     Top             =   240
                     Width           =   930
                     _ExtentX        =   1640
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
                     ButtonImage     =   "FrmVendorContract.frx":22EC
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin MSDataListLib.DataCombo dcitems 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   39
                     Top             =   240
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
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   14280
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   3600
                  Width           =   1590
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   14400
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   3240
                  Width           =   2310
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   9360
                  Width           =   2175
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5835
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2835
                  Left            =   120
                  TabIndex        =   7
                  Top             =   3000
                  Width           =   11970
                  _cx             =   21114
                  _cy             =   5001
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
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmVendorContract.frx":2886
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
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   13680
                  TabIndex        =   11
                  Top             =   2040
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
                  Left            =   14280
                  TabIndex        =   12
                  Top             =   1560
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
                  Left            =   13440
                  TabIndex        =   26
                  Top             =   840
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   8
                  Left            =   14280
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   2640
                  Width           =   1785
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   375
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   960
                  Width           =   855
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6555
         Width           =   15945
         _cx             =   28125
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
            Left            =   13080
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   0
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
            ButtonImage     =   "FrmVendorContract.frx":2B4F
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   13320
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
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
            ButtonImage     =   "FrmVendorContract.frx":2EE9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   10680
            TabIndex        =   19
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9360
            TabIndex        =   20
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   7920
            TabIndex        =   21
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
            CausesValidation=   0   'False
            Height          =   375
            Index           =   3
            Left            =   6360
            TabIndex        =   22
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            Height          =   435
            Index           =   4
            Left            =   4920
            TabIndex        =   23
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   767
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
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
            Index           =   5
            Left            =   3480
            TabIndex        =   25
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   11160
            TabIndex        =   28
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
            BCOL            =   12632319
            BCOLO           =   12632319
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVendorContract.frx":3283
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   1800
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
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
            ButtonImage     =   "FrmVendorContract.frx":329F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
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
            TabIndex        =   17
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
            TabIndex        =   16
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
      ButtonImage     =   "FrmVendorContract.frx":9B01
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmVendorContract"
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

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function Create_dev1()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    'StrAccountCode = Account_Code_dynamic
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

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

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        If Trim(Me.DBCboClientName.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete TblVendorContractDetails where TblVendorContractD=" & val(Me.TxtTblVendorContractD.text)
   
    End If
    
    rs("TblVendorContractD").value = TxtTblVendorContractD.text
    rs("VendorId").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
  '  rs("TxtVendorCode").value = IIf(TxtSearchCode.text <> "", Trim(TxtSearchCode.text), Null)
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
      
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", "", Me.txtRemarks.text)
 
    If ChkLocked.value = vbChecked Then
        rs("Locked").value = 1
    Else
        rs("Locked").value = 0
    End If

    rs.update
    CuurentLogdata
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblVendorContractDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
         
                RsDev.AddNew
                RsDev("TblVendorContractD").value = Me.TxtTblVendorContractD.text
            
                RsDev("ItemId").value = val(.TextMatrix(i, .ColIndex("ItemId")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                RsDev.update
                    
            End If
            
            '
        Next i

        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
            
    End With
 
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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
     
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Me.TxtTblVendorContractD.text = CStr(new_id("TblVendorContract", "TblVendorContractD", "", True))
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True
            Option2.value = True

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
         
            CuurentLogdata

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

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmVendorContractSearch
            FrmVendorContractSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            '   ViewDataList
        Case 20
            addrow

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If TxtTblVendorContractD.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (TxtTblVendorContractD.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.Execute "delete TblVendorContractDetails where TblVendorContractD=" & val(Me.TxtTblVendorContractD.text)
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
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

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
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

Private Sub DBCboClientName_Change()
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 2
    TxtSearchCode.text = fullcode
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
      FrmCompanySearch.lblSearchtype.Caption = 20915
       FrmCompanySearch.show vbModal
        
        
    End If
End Sub

Private Sub dcitems_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 20915
        FrmItemSearch.show vbModal
    End If

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

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = " « ›«ﬁÌ«  «·„Ê—œÌ‰   "
    ScreenNameEnglish = "  Supplier Contracts "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

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

    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
    Dcombos.GetItemsNames dcitems

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
    StrSQL = "select * From TblVendorContract  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

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
    ISButton2.Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Supplier Contract"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Start Date"
    lbl(2).Caption = "End Date"
    lbl(0).Caption = "Supplier"
    lbl(3).Caption = "Remarks"
    ChkLocked.Caption = "Locked"
    Option1.Caption = "All Items"
    Option2.Caption = "Select Item"

    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"

    CmdRemove.Caption = "Remove Line"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"

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

Function CuurentLogdata(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·« ›«ﬁÌ…    " & TxtTblVendorContractD.text & Chr(13) & " «·„Ê—œ " & DBCboClientName.text & Chr(13) & "  „œ Â« „‰  " & dbFromDate & Chr(13) & "  «·Ï " & dbTodate & Chr(13) & "  „·«ÕŸ«  " & txtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & Chr(13) & "   „ «Ìﬁ«› «· ⁄«„· "
    End If
                    
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Contract No    " & TxtTblVendorContractD.text & Chr(13) & " Supplier " & DBCboClientName.text & Chr(13) & " From   " & dbFromDate & Chr(13) & "  To  " & dbTodate & Chr(13) & "  Remarks " & txtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & Chr(13) & " Locked "
    End If
                    
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
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
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    If Col = Grid.ColIndex("ItemCode") Or Col = Grid.ColIndex("ItemName") Then
        RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("UnitName") Then
        RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), Grid.TextMatrix(Row, Grid.ColIndex("UnitName")), , , , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("Price") Then
        RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , (Grid.TextMatrix(Row, Grid.ColIndex("Price"))), , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("Discount") Then
        RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , Grid.TextMatrix(Row, Grid.ColIndex("Discount")), , Me.TxtTblVendorContractD

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTblVendorContractD.text = IIf(IsNull(rs("TblVendorContractD").value), "", rs("TblVendorContractD").value)
 
    dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)

    DBCboClientName.BoundText = IIf(IsNull(rs("VendorId").value), "", rs("VendorId").value)
   ' TxtSearchCode.text = IIf(IsNull(rs("TxtVendorCode").value), "", rs("TxtVendorCode").value)
   
    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
    
    If (rs("Locked").value) = True Then
     ChkLocked.value = vbChecked
     Else
     ChkLocked.value = vbUnchecked
     End If

    StrSQL = " SELECT     dbo.TblVendorContractDetails.TblVendorContractD, dbo.TblVendorContractDetails.UnitID, dbo.TblVendorContractDetails.ItemID, dbo.TblVendorContractDetails.Discount, "
    StrSQL = StrSQL & "     dbo.TblVendorContractDetails.Price , dbo.TblUnites.unitname, dbo.TblItems.ItemName, dbo.TblItems.ItemCode"
    StrSQL = StrSQL & " FROM         dbo.TblVendorContractDetails INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblVendorContractDetails.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblVendorContractDetails.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where TblVendorContractD=" & val(Me.TxtTblVendorContractD.text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub ISButton2_Click()
On Error GoTo ErrTrap
   If val(Me.TxtTblVendorContractD.text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
    sql = "SELECT    dbo.TblVendorContract.TblVendorContractD, dbo.TblVendorContract.VendorId, dbo.TblVendorContract.FromDate, dbo.TblVendorContract.Todate, dbo.TblVendorContract.Remarks,"
    sql = sql & "     dbo.TblVendorContract.Locked, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Type, dbo.TblCustemers.Fullcode,"
    sql = sql & "      dbo.TblCustemers.CustomerandVendor, dbo.TblVendorContractDetails.Id, dbo.TblVendorContractDetails.TblVendorContractD AS TblVendorContractDETALS, dbo.TblVendorContractDetails.UnitID,"
    sql = sql & "      dbo.TblVendorContractDetails.ItemID, dbo.TblVendorContractDetails.Price, dbo.TblVendorContractDetails.Discount, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.code,"
    sql = sql & "      dbo.TblItems.Fullcode AS FullcodeITEM, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee"
    sql = sql & " FROM         dbo.TblUnites RIGHT OUTER JOIN"
    sql = sql & "      dbo.TblVendorContractDetails ON dbo.TblUnites.UnitID = dbo.TblVendorContractDetails.UnitID LEFT OUTER JOIN"
    sql = sql & "     dbo.TblItems ON dbo.TblVendorContractDetails.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
    sql = sql & "      dbo.TblVendorContract ON dbo.TblVendorContractDetails.TblVendorContractD = dbo.TblVendorContract.TblVendorContractD LEFT OUTER JOIN"
    sql = sql & "      dbo.TblCustemers ON dbo.TblVendorContract.VendorId = dbo.TblCustemers.CusID"
    sql = sql & " Where (dbo.TblVendorContract.TblVendorContractD = " & val(TxtTblVendorContractD.text) & ")"
        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "VendorContractRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "VendorContractRPTEE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
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
ErrTrap:
  End Function

Private Sub TxtItemCode_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 20915
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
        ISButton2.Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        ISButton2.Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
        ISButton2.Enabled = True
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
 Public Function FindRec(ByVal TblVendorContractD As Long)
    On Error GoTo ErrTrap
    rs.find "TblVendorContractD=" & TblVendorContractD, , adSearchForward, 1
    If Not (rs.EOF) Then
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
       ' BtnUndo_Click
    End If
  End Function
' aladein add
Private Sub DBCboClientName_Click(Area As Integer)
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 2
    TxtSearchCode.text = fullcode
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
        DBCboClientName.BoundText = CUSTID
    End If
 End Sub

'''''' items
Private Sub dcitems_Click(Area As Integer)
  On Error Resume Next
    If val(dcitems.BoundText) = 0 Then Exit Sub
     Me.TxtItemCode.text = GetItemCode(val(Me.dcitems.BoundText))
   End Sub
Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.dcitems.BoundText = ""
        Else
            Me.dcitems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If
End Sub
''''''''''''''''''''''''''''''''
