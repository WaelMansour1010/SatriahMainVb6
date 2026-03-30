VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProductionOrder22 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«„— «·‘€· / «·«‰ «Ã   / «· Ã„Ì⁄"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14715
   Icon            =   "FrmProductionOrder2.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   14715
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   8415
      Index           =   15
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14715
      _cx             =   25956
      _cy             =   14843
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   735
         Index           =   6
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   14700
         _cx             =   25929
         _cy             =   1296
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
         Caption         =   "«„— «·‘€· / «·«‰ «Ã   / «· Ã„Ì⁄"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   0
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
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
            Height          =   375
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtnots2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   510
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1860
            TabIndex        =   4
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmProductionOrder2.frx":000C
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
            Left            =   1005
            TabIndex        =   5
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
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
            ButtonImage     =   "FrmProductionOrder2.frx":03A6
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
            Left            =   2670
            TabIndex        =   6
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
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
            ButtonImage     =   "FrmProductionOrder2.frx":0740
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
            Left            =   165
            TabIndex        =   7
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmProductionOrder2.frx":0ADA
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   495
            Left            =   4800
            Top             =   120
            Width           =   4005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »⁄„· «Ê«„— «·«‰ «Ã Ê«‰‘«¡ «–Ê‰«  «” ·«„ «·«‰ «Ã «· «„ «·Ì« ÊÕ”«» «· ﬂ«·Ì›"
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
            Height          =   420
            Index           =   44
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   3975
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   6840
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   14730
         _cx             =   25982
         _cy             =   12065
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
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   $"FrmProductionOrder2.frx":0E74
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
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmProductionOrder2.frx":0F2C
         Picture(1)      =   "FrmProductionOrder2.frx":12C6
         Picture(2)      =   "FrmProductionOrder2.frx":1660
         Begin C1SizerLibCtl.C1Elastic EleMain 
            Height          =   6375
            Left            =   45
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
               Height          =   435
               Index           =   3
               Left            =   270
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   5880
               Width           =   13545
               _cx             =   23892
               _cy             =   767
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
               BorderWidth     =   0
               ChildSpacing    =   0
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
               Begin VB.TextBox XPTxtSum 
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
                  Height          =   360
                  Left            =   9990
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   1635
               End
               Begin MSDataListLib.DataCombo DCboUserName 
                  Height          =   315
                  Left            =   7110
                  TabIndex        =   13
                  Top             =   45
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ì «·„Ê«œ «·Œ«„"
                  Height          =   255
                  Index           =   3
                  Left            =   11625
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”Ã· «·Õ«·Ì:"
                  Height          =   240
                  Index           =   0
                  Left            =   3150
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   120
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·”Ã·« :"
                  Height          =   240
                  Index           =   2
                  Left            =   1095
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   120
                  Width           =   1305
               End
               Begin VB.Label XPTxtCurrent 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   270
                  Left            =   2055
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   105
                  Width           =   825
               End
               Begin VB.Label XPTxtCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   240
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   135
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ—— »Ê«”ÿ… : "
                  Height          =   315
                  Index           =   1
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   75
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ: "
                  Height          =   315
                  Index           =   32
                  Left            =   5610
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   120
                  Width           =   1365
               End
               Begin VB.Label LblTotalQty 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   4380
                  TabIndex        =   14
                  Top             =   0
                  Width           =   1500
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3975
               Index           =   5
               Left            =   0
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   2340
               Width           =   14640
               _cx             =   25823
               _cy             =   7011
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
                  Height          =   660
                  Index           =   2
                  Left            =   0
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   14505
                  _cx             =   25585
                  _cy             =   1164
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
                  Begin VB.TextBox TxtPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   825
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   285
                     Width           =   2055
                  End
                  Begin VB.TextBox TxtQuantity 
                     Alignment       =   1  'Right Justify
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   2880
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   285
                     Width           =   1905
                  End
                  Begin VB.ComboBox CboItemCase 
                     Height          =   315
                     Left            =   5070
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   24
                     Top             =   285
                     Width           =   1905
                  End
                  Begin MSDataListLib.DataCombo DCboItemsName 
                     Height          =   315
                     Left            =   6975
                     TabIndex        =   27
                     Top             =   285
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCboItemsCode 
                     Height          =   315
                     Left            =   10950
                     TabIndex        =   28
                     Top             =   285
                     Width           =   3420
                     _ExtentX        =   6033
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin ImpulseButton.ISButton CmdAdd 
                     Height          =   360
                     Left            =   135
                     TabIndex        =   29
                     Top             =   255
                     Width           =   555
                     _ExtentX        =   979
                     _ExtentY        =   635
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
                     BackStyle       =   0
                     ButtonImage     =   "FrmProductionOrder2.frx":19FA
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
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”⁄—"
                     Height          =   240
                     Index           =   26
                     Left            =   825
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   0
                     Width           =   2055
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬂ„Ì…"
                     Height          =   240
                     Index           =   27
                     Left            =   3015
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   0
                     Width           =   1905
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ«·… «·’‰›"
                     Height          =   240
                     Index           =   29
                     Left            =   5340
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   0
                     Width           =   1635
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈”„ «·’‰›"
                     Height          =   240
                     Index           =   30
                     Left            =   7260
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   0
                     Width           =   3000
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂÊœ «·’‰›"
                     Height          =   240
                     Index           =   31
                     Left            =   11220
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     Top             =   0
                     Width           =   3015
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FG 
                  Height          =   2370
                  Left            =   0
                  TabIndex        =   245
                  Top             =   840
                  Width           =   14505
                  _cx             =   25585
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
                  Rows            =   2
                  Cols            =   28
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmProductionOrder2.frx":1D94
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
               Begin MSComctlLib.Toolbar TBar 
                  Height          =   630
                  Left            =   540
                  TabIndex        =   35
                  Top             =   3210
                  Width           =   12180
                  _ExtentX        =   21484
                  _ExtentY        =   1111
                  ButtonWidth     =   609
                  ButtonHeight    =   1005
                  Appearance      =   1
                  _Version        =   393216
               End
               Begin VB.Label LblItemsCount 
                  Alignment       =   2  'Center
                  BackColor       =   &H00404040&
                  ForeColor       =   &H0000FFFF&
                  Height          =   390
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   3210
                  Width           =   540
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2355
               Index           =   0
               Left            =   0
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   -75
               Width           =   14640
               _cx             =   25823
               _cy             =   4154
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
               Begin VB.TextBox txtShipmentPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   10950
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   -240
                  Width           =   2190
               End
               Begin VB.TextBox XPTxtBillID 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   -510
                  Visible         =   0   'False
                  Width           =   1920
               End
               Begin VB.TextBox TxtFillData 
                  Alignment       =   1  'Right Justify
                  Height          =   270
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   -510
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   -450
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.TextBox TxtTransSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   11355
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   120
                  Width           =   1650
               End
               Begin VB.ComboBox CboPriceType 
                  Height          =   315
                  Left            =   15330
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   87
                  Top             =   360
                  Width           =   2325
               End
               Begin VB.Frame Frame2 
                  Height          =   1815
                  Left            =   15180
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1680
                  Width           =   5760
                  Begin VB.TextBox Text2 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.TextBox Text3 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   1320
                     Width           =   1455
                  End
                  Begin VB.TextBox Text7 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   5400
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   600
                     Width           =   3855
                  End
                  Begin MSComCtl2.DTPicker DTPicker1 
                     Height          =   315
                     Left            =   240
                     TabIndex        =   78
                     Top             =   1320
                     Width           =   1320
                     _ExtentX        =   2328
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DataCombo9 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   79
                     Top             =   240
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo11 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   80
                     Top             =   960
                     Width           =   1425
                     _ExtentX        =   2514
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰Ê⁄ «·«„—"
                     Height          =   285
                     Index           =   19
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "»‰«¡ ⁄·Ï"
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Index           =   20
                     Left            =   9600
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   480
                     Width           =   1095
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·⁄„·…"
                     Height          =   285
                     Index           =   21
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   960
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "—ﬁ„ «·Õ”«»"
                     Height          =   285
                     Index           =   22
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   1320
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   285
                     Index           =   23
                     Left            =   1560
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   960
                     Width           =   975
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   " «·«‰ Â«¡"
                     Height          =   285
                     Index           =   24
                     Left            =   1680
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   1320
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame3 
                  Height          =   1815
                  Left            =   15600
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   1680
                  Width           =   7800
                  Begin VB.TextBox Text5 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   4080
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   600
                     Width           =   2295
                  End
                  Begin MSComCtl2.DTPicker DTPicker2 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   63
                     Top             =   600
                     Width           =   2100
                     _ExtentX        =   3704
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DTPicker3 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   64
                     Top             =   960
                     Width           =   1620
                     _ExtentX        =   2858
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DTPicker4 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   65
                     Top             =   960
                     Width           =   2100
                     _ExtentX        =   3704
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DTPicker5 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   66
                     Top             =   1320
                     Width           =   1620
                     _ExtentX        =   2858
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DTPicker6 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   67
                     Top             =   1320
                     Width           =   2100
                     _ExtentX        =   3704
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·—ﬁ„"
                     Height          =   375
                     Left            =   6720
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   720
                     Width           =   975
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«· «—ÌŒ"
                     Height          =   375
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   720
                     Width           =   1335
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
                     Height          =   375
                     Left            =   6480
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   1080
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«· «—ÌŒ «·›⁄·Ì"
                     Height          =   375
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   1200
                     Width           =   1215
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     Caption         =   " «—ÌŒ «· √ŒÌ—"
                     Height          =   255
                     Left            =   6480
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   1440
                     Width           =   1215
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
                     Height          =   255
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   1440
                     Width           =   1575
                  End
               End
               Begin VB.TextBox TXTNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5205
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Text            =   "Text4"
                  Top             =   -960
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.TextBox TxtNoteSerial1 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   -480
                  Visible         =   0   'False
                  Width           =   2190
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1815
                  Left            =   15600
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   14505
                  Begin VB.CheckBox chkshipped 
                     Alignment       =   1  'Right Justify
                     Caption         =   " „ «·‘Õ‰"
                     Height          =   195
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   -2760
                     Width           =   975
                  End
                  Begin VB.TextBox TxtShipmentArae 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   15600
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   600
                     Width           =   3735
                  End
                  Begin VB.ComboBox CboPayMentType 
                     Height          =   315
                     Left            =   16680
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   43
                     Top             =   600
                     Width           =   2145
                  End
                  Begin VB.TextBox TxtStation 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   -240
                     Visible         =   0   'False
                     Width           =   1545
                  End
                  Begin VB.TextBox txtMixID 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.TextBox ProkerId 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   -120
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin MSDataListLib.DataCombo Dccurrency 
                     Height          =   315
                     Left            =   15000
                     TabIndex        =   46
                     Top             =   600
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo4 
                     Height          =   315
                     Left            =   16800
                     TabIndex        =   47
                     Top             =   960
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcshipmentMethod 
                     Height          =   315
                     Left            =   16800
                     TabIndex        =   48
                     Top             =   240
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo6 
                     Height          =   315
                     Left            =   14760
                     TabIndex        =   49
                     Top             =   600
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo7 
                     Height          =   315
                     Left            =   15240
                     TabIndex        =   50
                     Top             =   240
                     Width           =   2145
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo8 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   51
                     Top             =   2040
                     Width           =   1905
                     _ExtentX        =   3360
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·⁄„·Â"
                     Height          =   285
                     Index           =   12
                     Left            =   15720
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   600
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·»·œ"
                     Height          =   285
                     Index           =   13
                     Left            =   14880
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   14
                     Left            =   15720
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   15
                     Left            =   15120
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   600
                     Width           =   975
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«· ’‰Ì›"
                     Height          =   285
                     Index           =   16
                     Left            =   16080
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«· ”⁄Ì—"
                     Height          =   285
                     Index           =   18
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   2040
                     Width           =   975
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "ÃÂ… «· ”·Ì„"
                     ForeColor       =   &H00000000&
                     Height          =   375
                     Index           =   0
                     Left            =   15600
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   600
                     Width           =   855
                  End
               End
               Begin VB.TextBox TxtManualNo1 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   8340
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   120
                  Width           =   1920
               End
               Begin MSDataListLib.DataCombo DCboStoreName1 
                  Height          =   315
                  Left            =   9855
                  TabIndex        =   93
                  Top             =   4710
                  Visible         =   0   'False
                  Width           =   2460
                  _ExtentX        =   4339
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker XPDtbBill 
                  Height          =   315
                  Left            =   6015
                  TabIndex        =   94
                  Top             =   120
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   96141313
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton XPBtnNewClients 
                  Height          =   450
                  Left            =   6300
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   4470
                  Width           =   270
                  _ExtentX        =   476
                  _ExtentY        =   794
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmProductionOrder2.frx":223A
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdTemplate 
                  Height          =   480
                  Left            =   3420
                  TabIndex        =   96
                  Top             =   -1485
                  Visible         =   0   'False
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   847
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈œ—«Ã ⁄—÷ Ã«Â“"
                  BackColor       =   12632256
                  ForeColor       =   16711680
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   12632256
                  ColorHighlight  =   16777215
                  ColorHoverText  =   255
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledText=   16711680
                  ColorToggledHoverText=   255
                  ColorTextShadow =   -2147483637
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   630
                  Index           =   4
                  Left            =   5340
                  TabIndex        =   97
                  TabStop         =   0   'False
                  Top             =   -1800
                  Width           =   3690
                  _cx             =   6509
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
                  Begin VB.TextBox XPTxtTaxValue 
                     Alignment       =   1  'Right Justify
                     Height          =   390
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   150
                     Width           =   915
                  End
                  Begin VB.CheckBox XPChkTAX 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                     Height          =   330
                     Left            =   1860
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   210
                     Width           =   1815
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Enabled         =   0   'False
                     Height          =   240
                     Index           =   4
                     Left            =   990
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   285
                     Width           =   720
                  End
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   15330
                  TabIndex        =   101
                  Top             =   720
                  Width           =   2325
                  _ExtentX        =   4101
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   15045
                  TabIndex        =   102
                  Top             =   1050
                  Width           =   2325
                  _ExtentX        =   4101
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "6"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdConvert 
                  Height          =   285
                  Left            =   11355
                  TabIndex        =   103
                  Top             =   5640
                  Visible         =   0   'False
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   503
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " ÕÊÌ· ≈·Ì ›« Ê—…"
                  BackColor       =   12632256
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColorButton     =   12632256
                  ColorHighlight  =   16777215
                  ColorHoverText  =   255
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledText=   16711680
                  ColorToggledHoverText=   255
                  ColorTextShadow =   -2147483637
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   135
                  TabIndex        =   104
                  Top             =   120
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1935
                  Index           =   16
                  Left            =   0
                  TabIndex        =   199
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   14640
                  _cx             =   25823
                  _cy             =   3413
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
                  Begin VB.ComboBox CBoBasedON 
                     Height          =   315
                     ItemData        =   "FrmProductionOrder2.frx":25D4
                     Left            =   5160
                     List            =   "FrmProductionOrder2.frx":25D6
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   254
                     Top             =   120
                     Width           =   1515
                  End
                  Begin VB.TextBox TxtBatchNo 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   246
                     Top             =   1200
                     Width           =   1650
                  End
                  Begin VB.TextBox TxtWorkHour 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   231
                     Top             =   2040
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.TextBox TxtEmployeeID1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   12180
                     RightToLeft     =   -1  'True
                     TabIndex        =   208
                     Top             =   1560
                     Width           =   675
                  End
                  Begin VB.TextBox Text4 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   12180
                     RightToLeft     =   -1  'True
                     TabIndex        =   207
                     Top             =   1200
                     Width           =   675
                  End
                  Begin VB.TextBox txtMIxCode 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   206
                     Top             =   840
                     Width           =   1650
                  End
                  Begin VB.TextBox TxtStoreID 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   12180
                     RightToLeft     =   -1  'True
                     TabIndex        =   205
                     Top             =   480
                     Width           =   675
                  End
                  Begin VB.TextBox TxtStoreID1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   12180
                     RightToLeft     =   -1  'True
                     TabIndex        =   204
                     Top             =   840
                     Width           =   675
                  End
                  Begin VB.TextBox TxtResProductionNo 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   120
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   203
                     Top             =   480
                     Width           =   1650
                  End
                  Begin VB.TextBox TxtProductionPlanno 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   202
                     Top             =   120
                     Width           =   1650
                  End
                  Begin VB.TextBox Txt_order_no 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2730
                     RightToLeft     =   -1  'True
                     TabIndex        =   201
                     Top             =   120
                     Width           =   2370
                  End
                  Begin VB.TextBox txtRemark 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   135
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   200
                     Top             =   1560
                     Width           =   6570
                  End
                  Begin MSComCtl2.DTPicker startDate 
                     Height          =   330
                     Left            =   5190
                     TabIndex        =   209
                     Top             =   480
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   582
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DCboStoreName2 
                     Height          =   315
                     Left            =   8340
                     TabIndex        =   210
                     Top             =   480
                     Width           =   3840
                     _ExtentX        =   6773
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker EndDate 
                     Height          =   330
                     Left            =   5190
                     TabIndex        =   211
                     Top             =   840
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   582
                     _Version        =   393216
                     Format          =   96141313
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker startTime 
                     Height          =   285
                     Left            =   2730
                     TabIndex        =   212
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   503
                     _Version        =   393216
                     CustomFormat    =   "'Time: 'hh:mm tt"
                     Format          =   96141315
                     UpDown          =   -1  'True
                     CurrentDate     =   39240
                  End
                  Begin MSComCtl2.DTPicker EndTime 
                     Height          =   285
                     Left            =   2730
                     TabIndex        =   213
                     Top             =   840
                     Visible         =   0   'False
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   503
                     _Version        =   393216
                     CustomFormat    =   "'Time: 'hh:mm tt"
                     Format          =   96141315
                     UpDown          =   -1  'True
                     CurrentDate     =   39240
                  End
                  Begin MSDataListLib.DataCombo DCboStoreName 
                     Height          =   315
                     Left            =   8340
                     TabIndex        =   214
                     Top             =   840
                     Width           =   3840
                     _ExtentX        =   6773
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DBCboClientName 
                     Height          =   315
                     Left            =   8340
                     TabIndex        =   215
                     Top             =   120
                     Width           =   4515
                     _ExtentX        =   7964
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCDriver 
                     Height          =   315
                     Left            =   8340
                     TabIndex        =   216
                     Top             =   1200
                     Width           =   3840
                     _ExtentX        =   6773
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcHey 
                     Height          =   315
                     Left            =   2790
                     TabIndex        =   217
                     Top             =   1200
                     Width           =   3915
                     _ExtentX        =   6906
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "7"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCEmp1 
                     Height          =   315
                     Left            =   8340
                     TabIndex        =   218
                     Top             =   1560
                     Width           =   3840
                     _ExtentX        =   6773
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»‰«¡ ⁄·Ï"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   55
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   253
                     Top             =   120
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·»« ‘"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   53
                     Left            =   1785
                     RightToLeft     =   -1  'True
                     TabIndex        =   247
                     Top             =   1200
                     Width           =   810
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„Ãÿ…"
                     ForeColor       =   &H00000000&
                     Height          =   165
                     Index           =   48
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   233
                     Top             =   1560
                     Visible         =   0   'False
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã„«·Ì ”«⁄«  «·«” Â·«ﬂ ··Œÿ"
                     ForeColor       =   &H00000000&
                     Height          =   435
                     Index           =   37
                     Left            =   2460
                     RightToLeft     =   -1  'True
                     TabIndex        =   232
                     Top             =   2040
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ  »œ«Ì… «·«‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   28
                     Left            =   6705
                     RightToLeft     =   -1  'True
                     TabIndex        =   230
                     Top             =   480
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õœœ «·”«∆ﬁ"
                     Height          =   255
                     Index           =   82
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   229
                     Top             =   1200
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„‘—›"
                     Height          =   255
                     Index           =   52
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   228
                     Top             =   1560
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„Êﬁ⁄"
                     Height          =   195
                     Index           =   50
                     Left            =   6705
                     TabIndex        =   227
                     Top             =   1200
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬂÊœ «·„Ìﬂ”"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   49
                     Left            =   1785
                     RightToLeft     =   -1  'True
                     TabIndex        =   226
                     Top             =   840
                     Width           =   810
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "”‰œ ÕÃ“"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   47
                     Left            =   1635
                     RightToLeft     =   -1  'True
                     TabIndex        =   225
                     Top             =   480
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Œÿ… ≈‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   45
                     Left            =   1635
                     RightToLeft     =   -1  'True
                     TabIndex        =   224
                     Top             =   120
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄„Ì·"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   42
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   223
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.Label Label9 
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ« "
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   7455
                     RightToLeft     =   -1  'True
                     TabIndex        =   222
                     Top             =   1560
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„Œ“‰ «·„Ê«œ «·Œ«„"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   33
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   221
                     Top             =   480
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„Œ“‰  «·«‰ «Ã «· «„"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   34
                     Left            =   13080
                     RightToLeft     =   -1  'True
                     TabIndex        =   220
                     Top             =   840
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ ‰Â«Ì… «·«‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   35
                     Left            =   6705
                     RightToLeft     =   -1  'True
                     TabIndex        =   219
                     Top             =   840
                     Width           =   1500
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Œ“‰"
                  Height          =   270
                  Index           =   8
                  Left            =   12450
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   5520
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ì· / «·„Ê—œ"
                  ForeColor       =   &H00000000&
                  Height          =   240
                  Index           =   7
                  Left            =   15330
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   6
                  Left            =   7395
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   120
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·«„—"
                  ForeColor       =   &H00000000&
                  Height          =   270
                  Index           =   5
                  Left            =   13545
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   120
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·«„—"
                  Height          =   240
                  Index           =   9
                  Left            =   16560
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—ﬂ“ «· ﬂ·›…"
                  Height          =   285
                  Index           =   10
                  Left            =   16425
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‘—Ê⁄"
                  Height          =   270
                  Index           =   11
                  Left            =   14775
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   810
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  ForeColor       =   &H00000000&
                  Height          =   240
                  Index           =   36
                  Left            =   4245
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   120
                  Width           =   540
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
                  Index           =   43
                  Left            =   5055
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ ÌœÊÌ"
                  ForeColor       =   &H00000000&
                  Height          =   270
                  Index           =   46
                  Left            =   10395
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   120
                  Width           =   690
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   7
            Left            =   15375
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TxtTotalQty 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   4440
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   4200
               Width           =   1590
            End
            Begin VB.TextBox txtCount 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   6480
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   4200
               Width           =   1590
            End
            Begin VB.TextBox TxtTotalMaterials 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   1365
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   4200
               Width           =   2190
            End
            Begin VSFlex8UCtl.VSFlexGrid FG1 
               Height          =   2670
               Left            =   1365
               TabIndex        =   117
               Top             =   1320
               Width           =   12315
               _cx             =   21722
               _cy             =   4710
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
               Rows            =   2
               Cols            =   17
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProductionOrder2.frx":25D8
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰ »«·„Ê«œ «·Œ«„ «·„ÿ·Ê»… ·Â–« «·«„— Ê«· Ì ”Ì „ ”Õ»Â« „‰  „Œ“‰ «·„Ê«œ «·Œ«„"
               Height          =   255
               Left            =   8205
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   600
               Width           =   6030
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„Ê«œ «·Œ«„"
               Height          =   375
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   4320
               Width           =   2475
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   8
            Left            =   15675
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TxtUsedPowerPriceHTotal 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   4320
               Width           =   2055
            End
            Begin VB.TextBox TxtUsedElectricPriceHTotal 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   4800
               Width           =   2055
            End
            Begin VB.TextBox TxtHourdippTotal 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   5280
               Width           =   2055
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ì „ «· ⁄«„· „⁄ ŒÿÊÿ «·«‰ «Ã"
               Enabled         =   0   'False
               Height          =   375
               Left            =   10260
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   240
               Width           =   2880
            End
            Begin VB.TextBox TXTLineExpenses 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   5760
               Width           =   2055
            End
            Begin VB.TextBox Shifttime 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7815
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   4800
               Width           =   2055
            End
            Begin VSFlex8Ctl.VSFlexGrid FGLine 
               Height          =   3525
               Left            =   135
               TabIndex        =   124
               Top             =   720
               Width           =   14475
               _cx             =   25532
               _cy             =   6218
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
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   15
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmProductionOrder2.frx":2887
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
            Begin MSDataListLib.DataCombo DcLine 
               Height          =   315
               Left            =   7815
               TabIndex        =   125
               Top             =   4440
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   20
               Left            =   6990
               TabIndex        =   126
               Top             =   4320
               Width           =   690
               _ExtentX        =   1217
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
               ButtonImage     =   "FrmProductionOrder2.frx":2AED
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   21
               Left            =   5775
               TabIndex        =   127
               Top             =   4320
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
               ButtonImage     =   "FrmProductionOrder2.frx":2E87
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DTFrom 
               Height          =   315
               Left            =   10890
               TabIndex        =   128
               Top             =   5280
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   96141315
               UpDown          =   -1  'True
               CurrentDate     =   39240
            End
            Begin MSComCtl2.DTPicker DTTo 
               Height          =   315
               Left            =   7815
               TabIndex        =   129
               Top             =   5280
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   96141315
               UpDown          =   -1  'True
               CurrentDate     =   39240
            End
            Begin MSDataListLib.DataCombo DcShift 
               Height          =   315
               Left            =   10890
               TabIndex        =   130
               Tag             =   "«Œ — «·‘Ì›"
               Top             =   4800
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì ﬁÌ„… «” Â·«ﬂ «·ÊﬁÊœ ›Ì «·”«⁄Â"
               Height          =   375
               Left            =   2175
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   4320
               Width           =   3015
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «” Â·«ﬂ «·ﬂÂ—»«¡ ›Ì «·”«⁄…"
               Height          =   375
               Left            =   2175
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   4800
               Width           =   3015
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì ﬁÌ„… «·«Â·«ﬂ ›Ì «·”«⁄…"
               Height          =   375
               Left            =   2175
               RightToLeft     =   -1  'True
               TabIndex        =   256
               Top             =   5280
               Width           =   3015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — ŒÿÊÿ «·«‰ «Ã "
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   25
               Left            =   12825
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   4440
               Width           =   1515
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì „’«—Ì› ŒÿÊÿ «·«‰ «Ã "
               Height          =   375
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   5880
               Width           =   3015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   40
               Left            =   13665
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   5280
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   41
               Left            =   10425
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   5280
               Width           =   405
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘Ì› "
               Height          =   285
               Index           =   12
               Left            =   13395
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   4800
               Width           =   945
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”«⁄« "
               Height          =   375
               Left            =   9795
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   4800
               Width           =   960
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   9
            Left            =   15975
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TxtworkerTotalPerHour 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   540
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   4680
               Width           =   1095
            End
            Begin VB.TextBox TxtworkerTotal 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   540
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   5160
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid GridWorker 
               Height          =   3180
               Left            =   540
               TabIndex        =   139
               Top             =   1200
               Width           =   12315
               _cx             =   21722
               _cy             =   5609
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
               Rows            =   1
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmProductionOrder2.frx":3421
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   8
               Left            =   12045
               TabIndex        =   140
               Top             =   4440
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ⁄«„·"
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
               ButtonImage     =   "FrmProductionOrder2.frx":364E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì „’«—Ì› ⁄„«·Â  «·«‰ «Ã ›Ì «·”«⁄Â"
               Height          =   375
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   4680
               Width           =   3015
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  ⁄„«· «·«‰ «Ã"
               Height          =   375
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   720
               Width           =   2460
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì „’«—Ì› ⁄„«·Â  «·«‰ «Ã "
               Height          =   375
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   5160
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   10
            Left            =   16275
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.Frame Frame4 
               Enabled         =   0   'False
               Height          =   3615
               Left            =   15180
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   4800
               Visible         =   0   'False
               Width           =   14505
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3615
               Index           =   17
               Left            =   135
               TabIndex        =   234
               TabStop         =   0   'False
               Top             =   360
               Width           =   14370
               _cx             =   25347
               _cy             =   6376
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
               Begin VB.CommandButton Command2 
                  Caption         =   "⁄—÷ «·„’—Ê›« "
                  Height          =   480
                  Left            =   9720
                  RightToLeft     =   -1  'True
                  TabIndex        =   242
                  Top             =   3135
                  Visible         =   0   'False
                  Width           =   4200
               End
               Begin VB.TextBox Txt_EXport 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   9195
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   241
                  Text            =   "0"
                  Top             =   2895
                  Width           =   2415
               End
               Begin VB.TextBox TXTFinacilaTotal 
                  Alignment       =   2  'Center
                  Height          =   405
                  Left            =   510
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   240
                  Text            =   "0"
                  Top             =   2895
                  Width           =   2175
               End
               Begin VB.CommandButton Command6 
                  Caption         =   "Command6"
                  Height          =   375
                  Left            =   6705
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   3135
                  Visible         =   0   'False
                  Width           =   4155
               End
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   2340
                  Left            =   7335
                  TabIndex        =   235
                  Tag             =   "1"
                  Top             =   360
                  Width           =   6540
                  _cx             =   11536
                  _cy             =   4128
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
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmProductionOrder2.frx":3BE8
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
               Begin VSFlex8UCtl.VSFlexGrid grid4 
                  Height          =   2340
                  Left            =   165
                  TabIndex        =   236
                  Tag             =   "1"
                  Top             =   360
                  Width           =   7005
                  _cx             =   12356
                  _cy             =   4128
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
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmProductionOrder2.frx":3D93
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  ”‰œ«  «·„’—Ê›« "
                  Height          =   285
                  Index           =   51
                  Left            =   10890
                  RightToLeft     =   -1  'True
                  TabIndex        =   244
                  Top             =   2895
                  Width           =   2850
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·›Ê« Ì— «·„«·Ì…"
                  Height          =   285
                  Index           =   60
                  Left            =   -75
                  RightToLeft     =   -1  'True
                  TabIndex        =   243
                  Top             =   2895
                  Width           =   5115
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·›Ê« Ì— «·„«·Ì…"
                  Height          =   285
                  Index           =   38
                  Left            =   795
                  RightToLeft     =   -1  'True
                  TabIndex        =   238
                  Top             =   120
                  Width           =   3360
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰œ«  «·’—›"
                  Height          =   285
                  Index           =   54
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„’—Ê›«  Ê «·›Ê« Ì—  «·„«·ÌÂ"
               Height          =   375
               Left            =   11490
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   120
               Width           =   2460
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   11
            Left            =   16575
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TXTFactoryExpenses 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   6150
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   3600
               Width           =   1650
            End
            Begin VB.TextBox TxtIndirectCostForProduction 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   6150
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   4080
               Width           =   1650
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   10950
               TabIndex        =   149
               Top             =   3480
               Width           =   1635
               _ExtentX        =   2884
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
               ButtonImage     =   "FrmProductionOrder2.frx":3F59
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   2340
               Left            =   270
               TabIndex        =   150
               Top             =   1080
               Width           =   12585
               _cx             =   22199
               _cy             =   4128
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
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
               Rows            =   1
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmProductionOrder2.frx":44F3
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   240
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   151
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox TxtDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   152
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label LblDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   153
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   154
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2955
                  _cx             =   1973752924
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmProductionOrder2.frx":4653
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   145
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„’«—Ì› «·’‰«⁄Ì…"
               Height          =   375
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   3600
               Width           =   2595
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ Ì«— «·„’—Ê›«  «·’‰«⁄Ì…"
               Height          =   375
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   720
               Width           =   2460
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì   «· ﬂ«·Ì› €Ì— «·„»«‘—…  ÿ»ﬁ« ··‰”»… «·„Õœœ…"
               Height          =   495
               Left            =   7935
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   4200
               Width           =   2460
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   12
            Left            =   16875
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.CommandButton CmdIssueVoucher 
               Caption         =   " «‰‘«¡ «–‰ ’—› «·Ì"
               Height          =   315
               Left            =   9990
               TabIndex        =   166
               Top             =   960
               Width           =   2730
            End
            Begin VB.TextBox TxtIssueSerial 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6975
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   960
               Width           =   1920
            End
            Begin VB.TextBox TxtresiveVoucher 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6960
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   1320
               Width           =   1920
            End
            Begin VB.CommandButton CmdResiveVoucher 
               Caption         =   "«‰‘«¡ «–‰ «÷«›… «·Ì"
               Height          =   315
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   1320
               Width           =   2865
            End
            Begin VB.CommandButton Command3 
               Caption         =   "⁄—÷ «·«–‰"
               Height          =   315
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   960
               Width           =   2190
            End
            Begin VB.CommandButton Command4 
               Caption         =   "⁄—÷ «·«–‰"
               Height          =   315
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   1320
               Width           =   2190
            End
            Begin VB.CommandButton Command5 
               Caption         =   "⁄—÷ «·ﬁÌœ"
               Height          =   315
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   960
               Width           =   2880
            End
            Begin VB.CommandButton Command7 
               Caption         =   "⁄—÷ «·ﬁÌœ"
               Height          =   315
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   1320
               Width           =   2880
            End
            Begin MSComCtl2.DTPicker ReciveDate 
               Height          =   315
               Left            =   6975
               TabIndex        =   167
               Top             =   1680
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   556
               _Version        =   393216
               Format          =   96141313
               CurrentDate     =   38784
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   4260
               Left            =   0
               TabIndex        =   250
               TabStop         =   0   'False
               Top             =   2040
               Width           =   17385
               _cx             =   30665
               _cy             =   7514
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
               Begin VSFlex8UCtl.VSFlexGrid FgMix 
                  Height          =   3690
                  Left            =   120
                  TabIndex        =   251
                  Top             =   420
                  Width           =   17235
                  _cx             =   30401
                  _cy             =   6509
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
                  Rows            =   2
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmProductionOrder2.frx":4BED
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
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«’‰«› «·„ﬂ”"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   435
                  Left            =   11280
                  TabIndex        =   252
                  Top             =   120
                  Width           =   3135
               End
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„  «·«–‰"
               Height          =   255
               Left            =   8895
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   1020
               Width           =   825
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   4515
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   840
               Width           =   7665
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·«–‰"
               Height          =   255
               Left            =   8895
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   1395
               Width           =   825
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "  „ﬂ‰ „‰ Œ·«· Â–… «·‘«‘… „‰ «‰‘«¡ ”‰œ ’—› ··„Ê«œ «·Œ«„ Ê”‰œ «” ·«„ ··„Ê«œ «· Ì  „ «‰ «ÃÂ« «·Ì«"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   420
               Index           =   39
               Left            =   2730
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   360
               Width           =   10545
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               Height          =   495
               Left            =   2730
               Top             =   360
               Width           =   10545
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” ·«„"
               Height          =   255
               Left            =   8895
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1680
               Width           =   960
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   13
            Left            =   17175
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TXTTotalIssueVouchers 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   540
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   4200
               Width           =   2055
            End
            Begin VSFlex8UCtl.VSFlexGrid GridIssueVoucer 
               Height          =   2670
               Left            =   540
               TabIndex        =   175
               Top             =   1440
               Width           =   12180
               _cx             =   21484
               _cy             =   4710
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
               Rows            =   2
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmProductionOrder2.frx":4EC2
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  ”‰œ«  «·’—› «·„Œ“‰Ì"
               Height          =   375
               Left            =   2595
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   4200
               Width           =   2460
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰ »«·„Ê«œ «·Œ«„ «·„ÿ·Ê»… ·Â–« «·«„— Ê«· Ì ”Ì „ ”Õ»Â« „‰  „Œ“‰ «·„Ê«œ «·Œ«„"
               Height          =   255
               Left            =   6435
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   960
               Visible         =   0   'False
               Width           =   6150
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6375
            Index           =   14
            Left            =   17475
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   45
            Width           =   14640
            _cx             =   25823
            _cy             =   11245
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
            Begin VB.TextBox TxtTotalEstimatedCost 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   405
               Left            =   6000
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   3600
               Width           =   1695
            End
            Begin VSFlex8Ctl.VSFlexGrid GridEstimatedCost 
               Height          =   2340
               Left            =   120
               TabIndex        =   180
               Top             =   1080
               Width           =   14400
               _cx             =   25400
               _cy             =   4128
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
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
               Rows            =   1
               Cols            =   17
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmProductionOrder2.frx":527E
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
               Begin VB.PictureBox Picture1 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   240
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   181
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox Text6 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   182
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label Label23 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   183
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo SmartCombo1 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   184
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2955
                  _cx             =   1973752924
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmProductionOrder2.frx":5519
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   145
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   11640
               TabIndex        =   185
               Top             =   3480
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
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
               ButtonImage     =   "FrmProductionOrder2.frx":5AB3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ Ì«— «·„’—Ê›«  «· ﬁœÌ—Ì…"
               Height          =   375
               Left            =   11880
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   480
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„’«—Ì› «·’‰«⁄Ì…"
               Height          =   375
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   3720
               Visible         =   0   'False
               Width           =   2055
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   7860
         Width           =   14715
         _cx             =   25956
         _cy             =   979
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   0
            Left            =   11340
            TabIndex        =   189
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   688
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
            Height          =   390
            Index           =   1
            Left            =   10260
            TabIndex        =   190
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   688
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
            Height          =   390
            Index           =   2
            Left            =   9105
            TabIndex        =   191
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   688
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
            Height          =   390
            Index           =   3
            Left            =   8100
            TabIndex        =   192
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
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
            Height          =   390
            Index           =   4
            Left            =   7065
            TabIndex        =   193
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
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
            Height          =   390
            Index           =   5
            Left            =   6030
            TabIndex        =   194
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   688
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
            Height          =   390
            Index           =   6
            Left            =   3720
            TabIndex        =   195
            Top             =   90
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   688
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
            Height          =   390
            Index           =   7
            Left            =   5025
            TabIndex        =   196
            Top             =   90
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   688
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
            Height          =   390
            Left            =   4530
            TabIndex        =   197
            Top             =   90
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   688
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   11
            Left            =   2385
            TabIndex        =   198
            Top             =   90
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «„— «· Õ„Ì·"
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
Attribute VB_Name = "FrmProductionOrder2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch
Dim expenses_total As Variant
Dim dblIndirectCostForProduction As Variant
Dim TxtNoteSerialV As String
Dim TxtNoteSerial1V As String
Dim autoedit As Integer
  Dim CurrentTransactionType  As Integer
Function GetCostItem(Optional Row As Integer) As Double
Dim i As Integer
Dim SumValu As Double
SumValu = 0
With FgMix
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("MianItemID"))) <> 0 And (.TextMatrix(i, .ColIndex("MixCode"))) = (FG.TextMatrix(Row, FG.ColIndex("MixNo"))) And val(.TextMatrix(i, .ColIndex("MianItemID"))) = val(FG.TextMatrix(Row, FG.ColIndex("Code"))) Then
SumValu = SumValu + val(.TextMatrix(i, .ColIndex("Valu")))
End If
Next i
End With
GetCostItem = SumValu
End Function
Function cal_expenses()
    On Error Resume Next
    Dim RowNum As Integer

    Dim item_Expenses_percentage As Double
    Dim RsUnitData As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
    Dim QtyBySmalltUnit As Double
    Dim StrSQL As String
If SystemOptions.ProductionRawMaterMix = True Then
   With FG
        For RowNum = 1 To FG.Rows - 1
            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If val(val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))) = 0 Then
              FG.TextMatrix(RowNum, FG.ColIndex("Count")) = 1
             End If
                FG.TextMatrix(RowNum, FG.ColIndex("Price")) = GetCostItem(RowNum) / val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))
                FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))) * val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            
            End If

        Next RowNum

    End With
Else
    If Not IsNumeric(TxtWorkHour) Then TxtWorkHour = 1

    If SystemOptions.AllowIndirectCost = True Then

        dblIndirectCostForProduction = (SystemOptions.IndirectCostPercentage / 100) * (val(Txt_EXport) + val(TXTFactoryExpenses.Text) + val(TXTFinacilaTotal) + val(Me.TXTTotalIssueVouchers) + (val(TXTLineExpenses) + val(TxtworkerTotal)))
    Else

        dblIndirectCostForProduction = 0
    End If

    Me.TxtIndirectCostForProduction = dblIndirectCostForProduction
 
    expenses_total = val(Txt_EXport) + val(TXTFactoryExpenses.Text) + val(TXTFinacilaTotal) + val(Me.TXTTotalIssueVouchers) + (val(TXTLineExpenses) + val(TxtworkerTotal)) + Round(dblIndirectCostForProduction, 2)
            
    With FG

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                'item_Expenses_percentage = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / XPTxtSum
       
                item_Expenses_percentage = (expenses_total / val(LblTotalQty))
               
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    QtyBySmalltUnit = RsUnitData("UnitFactor").value
           
                End If
             
                'FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = Round(item_Expenses_percentage * Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))), 2)
                If val(FG.TextMatrix(RowNum, FG.ColIndex("DistibutePercentage"))) > 0 Then
                    FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = (((expenses_total * val(FG.TextMatrix(RowNum, FG.ColIndex("DistibutePercentage")))) / 100) / val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                Else
                    FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = (item_Expenses_percentage) * QtyBySmalltUnit
        
                End If
        
                FG.TextMatrix(RowNum, FG.ColIndex("Price")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Expenses"))) + val(FG.TextMatrix(RowNum, FG.ColIndex("EstimatedCost")))
        
                FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))) * val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            End If

        Next RowNum

    End With
End If
End Function

Function cal_expensesnew()
    On Error Resume Next
    Dim RowNum As Integer

    Dim item_Expenses_percentage As Double
    Dim QtyTotal As Double
    Dim itemvalue As Double

    If QtyTotal > 0 Then
        itemvalue = expenses_total / QtyTotal
    End If

    If Not IsNumeric(TxtWorkHour) Then TxtWorkHour = 1
    expenses_total = (val(TXTLineExpenses) + val(TxtworkerTotal)) + (val(Txt_EXport) + val(TXTFinacilaTotal) + val(TXTFactoryExpenses.Text))

    With FG

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                '  item_Expenses_percentage = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / XPTxtSum
                'FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = Round((item_Expenses_percentage * expenses_total) / Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))), 2)
                '     FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = Round((item_Expenses_percentage * expenses_total) / Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))), 2)
                FG.TextMatrix(RowNum, FG.ColIndex("Price")) = Round(itemvalue, 2)
            
            End If

        Next RowNum

    End With

End Function

Private Sub CBoBasedON_Change()
Txt_order_no_Change
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim intDef As Integer

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"

            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            'FG.SetFocus
            'FG.Col = FG.ColIndex("Code")
            'FG.Row = FG.Rows - 1
            Me.CboPriceType.ListIndex = 0
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
          
            Me.FGLine.Clear flexClearScrollable, flexClearEverything
            Me.FGLine.Rows = 1

            Me.GridWorker.Clear flexClearScrollable, flexClearEverything
            Me.GridWorker.Rows = 1
            ' ⁄»… «–Ê‰«  «·’—›
            fillExpensesGrid
            ' ⁄»…   «·›Ê« Ì— «·„«·Ì…
            fillFinancialInvoiceGrid

            Dcbranch.BoundText = Current_branch

        Case 1
             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
If autoedit = 1 Then
autoedit = 0
Else
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 

End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            CuurentLogdata
        
        Case 2
         If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            Dim Msg  As String
            my_branch = Me.Dcbranch.BoundText
             
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '    DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            

    
    'DCboStoreName2
       Retrive_orders_data (val(TxtTransSerial.Text))
            SaveData

        Case 3
            Undo

        Case 4
                If ChekClodePeriod(XPDtbBill.value) = True Then
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

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            ' FrmBuySearch.DealingForm = GridTransType.ProductionOrder
            ' FrmBuySearch.Caption = "«·»ÕÀ ⁄‰  «„— «‰ «Ã "
            ' FrmBuySearch.Show
           Order_no_search2.RetrunType = 4
            Order_no_search2.show vbModal
            

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me

        Case 8
            RemoveWorker

        Case 9
            RemoveFactoryExpenses

        Case 20
            add_line (val(Me.DcLine.BoundText))

        Case 21
            remove_line
    
Case 11
   

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport2
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Function RemoveFactoryExpenses()

    With Me.Fg_Journal
  
        If .Row <= 0 Then Exit Function
        .RemoveItem .Row

        If .Rows > 1 Then
            TXTFactoryExpenses = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows, .ColIndex("value"))
        Else
            TXTFactoryExpenses = 0
        End If

    End With

    ReLineGrid

End Function

Function RemoveWorker()

    With Me.GridWorker
  
        If .Row <= 0 Then Exit Function
        .RemoveItem .Row
    End With

    With GridWorker
        TxtworkerTotal.Text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        TxtworkerTotalPerHour.Text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("hourprice"), .Rows - 1, .ColIndex("hourprice"))
 
    End With

    ReLineGrid

End Function

Function CalculateNets()

    With Me.FGLine

        If .Rows = 1 Then TXTLineExpenses = 0: Exit Function
    End With

    With Me.FGLine
        .Rows = .Rows + 1
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "Total"
        End If
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Variant
        Dim SngTotal1 As Variant
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedPowerPriceH"), .Rows - 1, .ColIndex("UsedPowerPriceH"))
        .TextMatrix(.Rows - 1, .ColIndex("UsedPowerPriceH")) = SngTotal
         TxtUsedPowerPriceHTotal.Text = SngTotal
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedElectricPriceH"), .Rows - 1, .ColIndex("UsedElectricPriceH"))
        .TextMatrix(.Rows - 1, .ColIndex("UsedElectricPriceH")) = SngTotal
         TxtUsedElectricPriceHTotal.Text = SngTotal
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("WorkerPriceH"), .Rows - 1, .ColIndex("WorkerPriceH"))
        .TextMatrix(.Rows - 1, .ColIndex("WorkerPriceH")) = SngTotal
         SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Hourdipp"), .Rows - 1, .ColIndex("Hourdipp"))
        .TextMatrix(.Rows - 1, .ColIndex("Hourdipp")) = SngTotal
         TxtHourdippTotal.Text = SngTotal
    
        '  SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedPowerPriceH"), .Rows - 1, .ColIndex("UsedPowerPriceH"))
        '  SngTotal1 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedElectricPriceH"), .Rows - 1, .ColIndex("UsedElectricPriceH"))
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        .TextMatrix(.Rows - 1, .ColIndex("total")) = SngTotal
        TXTLineExpenses = SngTotal
           
        '.TextMatrix(.Rows - 1, .ColIndex("LinePriceH")) = SngTotal
        ' TXTLineExpenses = Val(.TextMatrix(.Rows - 1, .ColIndex("UsedPowerPriceH"))) + Val(.TextMatrix(.Rows - 1, .ColIndex("UsedElectricPriceH"))) '= SngTotal  SngTotal + SngTotal1
    
        '    .AutoSize 0, .Cols - 1, False

    End With

    If Me.TxtModFlg.Text <> "R" Then
        Showworker
    End If

End Function

Function addWorkerToGrid(LineID As Long, Shift As Integer, FromTime As String, ToTime As String, Hour As Double, shiftname As String) As Boolean
    Dim StrSQL As String
    Dim i As Integer
    '»Ì«‰«  «·⁄«„·Ì‰ ›Ì «·Œÿ
    Dim RsEmployee As ADODB.Recordset
    Set RsEmployee = New ADODB.Recordset
    StrSQL = "Select * From TblProductLineWorker Where LineID=" & LineID

    If Shift = 1 Then
        StrSQL = StrSQL + "and shift1=1 "
    ElseIf Shift = 2 Then
        StrSQL = StrSQL + "and shift2=1 "
    ElseIf Shift = 3 Then
        StrSQL = StrSQL + "and shift3=1 "
    ElseIf Shift = 4 Then
        StrSQL = StrSQL + "and shift4=1 "
    End If

    StrSQL = StrSQL + " Order By id"
    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then

        With Me.GridWorker
            Dim Row As Long
            Row = .Rows
            .Rows = .Rows + RsEmployee.RecordCount
            For i = Row To .Rows - 1
                
                
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsEmployee("EmpID").value), 0, (RsEmployee("EmpID").value))
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsEmployee("EmpCode").value), "", RsEmployee("EmpCode").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsEmployee("EmpIname").value), "", RsEmployee("EmpIname").value)
                .TextMatrix(i, .ColIndex("hourprice")) = IIf(IsNull(RsEmployee("WorkerPriceH").value), 0, (RsEmployee("WorkerPriceH").value))
                .TextMatrix(i, .ColIndex("shift")) = shiftname
                '.TextMatrix(Row, .ColIndex("shift2")) = IIf(IsNull(RsEmployee("Shift2").value), 0, RsEmployee("Shift2").value)
                '.TextMatrix(Row, .ColIndex("shift3")) = IIf(IsNull(RsEmployee("Shift3").value), 0, RsEmployee("Shift3").value)
                '.TextMatrix(Row, .ColIndex("shift4")) = IIf(IsNull(RsEmployee("Shift4").value), 0, RsEmployee("Shift4").value)
                .TextMatrix(i, .ColIndex("from")) = FromTime
                .TextMatrix(i, .ColIndex("to")) = ToTime
                .TextMatrix(i, .ColIndex("hour")) = Hour
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("hourprice"))) * Hour
                        
                RsEmployee.MoveNext
            Next i

            '.AutoSize 0, .Cols - 1, False
                    
        End With

    End If

End Function

Function Showworker()

    Dim RowNum As Integer
    GridWorker.Clear flexClearScrollable, flexClearEverything
    GridWorker.Rows = 1
          
    For RowNum = 1 To FGLine.Rows - 1

        If FGLine.TextMatrix(RowNum, FGLine.ColIndex("id")) <> "" Then
            If addWorkerToGrid(val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("id"))), FGLine.TextMatrix(RowNum, FGLine.ColIndex("shift")), FGLine.TextMatrix(RowNum, FGLine.ColIndex("from")), FGLine.TextMatrix(RowNum, FGLine.ColIndex("to")), FGLine.TextMatrix(RowNum, FGLine.ColIndex("hour")), FGLine.TextMatrix(RowNum, FGLine.ColIndex("shiftname"))) Then
                        
            End If
        End If

    Next RowNum

    With GridWorker
        TxtworkerTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        TxtworkerTotalPerHour = .Aggregate(flexSTSum, .FixedRows, .ColIndex("hourprice"), .Rows - 1, .ColIndex("hourprice"))
        
    End With
    
End Function

Function remove_line()

    With Me.FGLine

        If .Rows - 1 = .Row Then Exit Function
        If .Rows >= 0 Then
            .RemoveItem Me.FGLine.Rows - 1
        End If

    End With

    With Me.FGLine

        If .Row <= 0 Then Exit Function
        .RemoveItem .Row
    End With

    CalculateNets

    With Me.FGLine

        If .Rows = 2 Then
    
            .RemoveItem Me.FGLine.Rows - 1
        End If

    End With

End Function

Function add_line(ID As Integer)
    On Error Resume Next
    Dim LngRow As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sql = "select * from TblProductLine where id=" & ID

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then Exit Function
    
    If Me.DcLine.BoundText = "" Then Exit Function

    With Me.FGLine

        If .Rows >= 3 Then
            .RemoveItem Me.FGLine.Rows - 1
        End If

    End With

    LngRow = Me.FGLine.Rows
    Me.FGLine.Rows = Me.FGLine.Rows + 1
    With Me.FGLine
        .TextMatrix(LngRow, .ColIndex("id")) = ID
        .TextMatrix(LngRow, .ColIndex("name")) = rs("name").value
        .TextMatrix(LngRow, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
        .TextMatrix(LngRow, .ColIndex("UsedPowerPriceH")) = IIf(Not IsNumeric(rs("UsedPowerPriceH").value), 0, rs("UsedPowerPriceH").value)
        .TextMatrix(LngRow, .ColIndex("UsedElectricPriceH")) = IIf(Not IsNumeric(rs("UsedElectricPriceH").value), 0, rs("UsedElectricPriceH").value)
        .TextMatrix(LngRow, .ColIndex("WorkerPriceH")) = IIf(Not IsNumeric(rs("WorkerPriceH").value), 0, rs("WorkerPriceH").value)
        .TextMatrix(LngRow, .ColIndex("LinePriceH")) = IIf(Not IsNumeric(rs("LinePriceH").value), 0, rs("LinePriceH").value)
        .TextMatrix(LngRow, .ColIndex("Hourdipp")) = IIf(Not IsNumeric(rs("HourdippTotal").value), 0, rs("HourdippTotal").value)
        .TextMatrix(LngRow, .ColIndex("from")) = Me.DTFrom.value
        .TextMatrix(LngRow, .ColIndex("to")) = Me.DTTo.value
        .TextMatrix(LngRow, .ColIndex("shift")) = val(DcShift.BoundText)
        .TextMatrix(LngRow, .ColIndex("shiftname")) = DcShift.Text
        Dim Hour As Integer
        Dim Minute As Double
        Dim totalhour As Double
        Hour = val(mId(Me.Shifttime.Text, 1, 2))
        Minute = val(mId(Me.Shifttime.Text, 4, 2)) / 60
        totalhour = Round(Hour + Minute, 2)
        .TextMatrix(LngRow, .ColIndex("hour")) = totalhour
        .TextMatrix(LngRow, .ColIndex("total")) = (val(.TextMatrix(LngRow, .ColIndex("Hourdipp"))) + val(.TextMatrix(LngRow, .ColIndex("UsedPowerPriceH"))) + val(.TextMatrix(LngRow, .ColIndex("UsedElectricPriceH")))) * totalhour
    End With

    CalculateNets
End Function

Function PrintReport1(order_no As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From QRY_items_orders_data where order_no='" & order_no & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status"
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub cmdAdd_Click()
'show_parts
End Sub

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.CboPriceType.ListIndex = 0 Then
        Set Frm = New frmsalebill
    ElseIf Me.CboPriceType.ListIndex = 1 Then
        Set Frm = New FrmBillBuy
    End If

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Dccurrency.BoundText = Me.Dccurrency.BoundText

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdHelp_Click()
   ' SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
   ' SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
    
       On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtTransSerial, "0506201801"

 
 
End Sub
Sub CreateIssueVoucher(Optional Row As Long)
Dim Msg As String
            If Trim(DCboStoreName2.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "ÌÃ»  ÕœÌœ      „Œ“‰ «·„Ê«œ «·Œ«„"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '    DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
           DeleteTransactiomsVoucher val(FG.TextMatrix(Row, FG.ColIndex("IssuTransID")))
          
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
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
   ' Dim Msg As String
    Dim mytext As Double
    '>>>>>>>>>>>>>>>>>>>>>>>>>
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ ’—› „‰ Â–… «·«„—   .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    Else
        Msg = "Create ISSUE Voucher to this order ?"
    End If

    ' On Error GoTo ErrTrap

    If MsgBox(Msg, vbYesNo, App.title) = vbYes Then

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
             
        If TxtNoteSerialV = "" Then
            If Notes_coding(val(Dcbranch.BoundText), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(Dcbranch.BoundText), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(Dcbranch.BoundText), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(Dcbranch.BoundText), XPDtbBill.value, 10, 180, , 27) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(Dcbranch.BoundText), XPDtbBill.value, 10, 180, , 27) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(Dcbranch.BoundText), XPDtbBill.value, 10, 180, , 27)
                End If
            End If
        End If

       mytext = Transaction_ID ' CStr(new_id("Transactions", "Transaction_ID", "", True))
            
      '  Me.TxtIssueSerial = TxtNoteSerial1V
        FG.TextMatrix(Row, FG.ColIndex("IssueSerial")) = TxtNoteSerial1V
                   
                   'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
        If Me.TxtModFlg.Text = "N" Then
    
        Else
 
           ' general_noteid = val(FG.TextMatrix(Row, FG.ColIndex("IssuNoteID")))
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
       ' TXTNoteID.Text = general_noteid
       FG.TextMatrix(Row, FG.ColIndex("IssuNoteID")) = general_noteid
        RsNotesGeneral("branch_no").value = val(Dcbranch.BoundText)
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 240
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim sql As String
         sql = "INSERT INTO  Transactions ( BranchId,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots2,NoteSerial,NoteSerial1,NoteId,Transaction_Type_Sub,WorkOrderNO)SELECT   " & val(Dcbranch.BoundText) & "," & Transaction_ID & "," & mytext & ",Transaction_Date,Transaction_Type = 27,CusID,StoreID1,UserID,Emp_ID,nots2=" & TxtTransSerial.Text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & " ,Transaction_Type_Sub=27 , " & TxtTransSerial.Text & " From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 26"

         Cn.Execute sql
            
            rs!nots2 = Transaction_ID
        rs!Product_Issue_voucher_Serial = TxtNoteSerial1V
        rs.update
       FG.TextMatrix(Row, FG.ColIndex("IssuTransID")) = Transaction_ID
 Cn.Execute "update Transaction_Details set  IssueSerial='" & (FG.TextMatrix(Row, FG.ColIndex("IssueSerial"))) & "',IssuTransID=" & val(FG.TextMatrix(Row, FG.ColIndex("IssuTransID"))) & " ,IssuNoteID=" & val(FG.TextMatrix(Row, FG.ColIndex("IssuNoteID"))) & " where Transaction_ID=" & val(XPTxtBillID.Text) & " and Item_ID =" & val(FG.TextMatrix(Row, FG.ColIndex("code"))) & " and MixNo='" & (FG.TextMatrix(Row, FG.ColIndex("MixNo"))) & "' "
        Set RSTransDetails = New ADODB.Recordset
      '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   Dim rs2 As ADODB.Recordset
   Set rs2 = New ADODB.Recordset
 
   StrSQL = "Select * from TblProductMixItems where MixCode='" & (FG.TextMatrix(Row, FG.ColIndex("MixNo"))) & "' and  TransectionID=" & val(XPTxtBillID.Text) & " and MianItemID =" & val(FG.TextMatrix(Row, FG.ColIndex("code"))) & " "
   rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   rs2.MoveFirst
        For RowNum = 1 To rs2.RecordCount
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = IIf(IsNull(rs2("ItemID").value), Null, rs2("ItemID").value)
                RSTransDetails("Quantity").value = IIf(IsNull(rs2("Qty").value), Null, rs2("Qty").value)
                RSTransDetails("SHOWQTY").value = IIf(IsNull(rs2("Qty").value), Null, rs2("Qty").value)
                RSTransDetails("showPrice").value = IIf(IsNull(rs2("Cost").value), Null, rs2("Cost").value)
                RSTransDetails("UnitID").value = IIf(IsNull(rs2("UnitId").value), Null, rs2("UnitId").value)
                          '«·ÊÕœ« 
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            LngUnitID = IIf(IsNull(rs2("UnitId").value), 0, rs2("UnitId").value)
            DblQty = IIf(IsNull(rs2("Qty").value), 0, rs2("Qty").value)

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = IIf(IsNull(rs2("Cost").value), 0, rs2("Cost").value) / RSTransDetails("QtyBySmalltUnit").value
            
            End If
            
             
                RSTransDetails.update
            rs2.MoveNext

        Next RowNum
       UpdateTransactionsCost CStr(Transaction_ID)
        CREATE_VOUCHER_GE2 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, Row
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «‰‘«¡ «·”‰œ"
        Else
        MsgBox "Create Successfully"
        End If

    End If
 
    TxtNoteSerial1V = ""
 rs.Resync
ErrTrap:
End Sub
Private Sub CmdIssueVoucher_Click()
Dim Msg As String
                       If Trim(DCboStoreName2.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "ÌÃ»  ÕœÌœ      „Œ“‰ «·„Ê«œ «·Œ«„"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '    DcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            
    'On Error GoTo errortrap
'    If TxtIssueSerial <> "" Then MsgBox " „ «‰‘«¡ «·”‰œ „‰ ﬁ»·": Exit Sub
           DeleteTransactiomsVoucher val(Txtnots2.Text)
       '   DeleteTransactiomsVoucher val(Txtnots2.text)
          
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
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
   ' Dim Msg As String
    Dim mytext As Double
    '>>>>>>>>>>>>>>>>>>>>>>>>>

 

'    rs.Close
'    rs.Open "select * from Transactions where Transaction_Serial = " & TxtTransSerial.text & " and Transaction_type = 26"
'
'    If rs.RecordCount = 0 Then MsgBox "«Õ›Ÿ «„— «·«‰ «Ã «Ê·«": Exit Sub
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ ’—› „‰ Â–… «·«„—   .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    Else
        Msg = "Create ISSUE Voucher to this order ?"
    End If

    ' On Error GoTo ErrTrap

    If MsgBox(Msg, vbYesNo, App.title) = vbYes Then

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
             
        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 27) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 27) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 27)
                End If
            End If
        End If

        ' ÕœÌÀ ÃœÊ· «· transaction ÊÊ÷€ —ﬁ„ «–‰ «·’—› ›Ì…
        'mytext = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=27"))
       mytext = Transaction_ID ' CStr(new_id("Transactions", "Transaction_ID", "", True))
         

         
        Me.TxtIssueSerial = TxtNoteSerial1V

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
        If Me.TxtModFlg.Text = "N" Then
    
        Else
 
            general_noteid = val(TXTNoteID.Text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.Text = general_noteid
          RsNotesGeneral("branch_no").value = val(Dcbranch.BoundText)
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 240
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim sql As String

        'sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId)SELECT " & Transaction_ID & "," & mytext & ",Transaction_Date,Transaction_Type = 19,CusID," & Val(Me.DCboStoreName1.BoundText) & ",UserID,Emp_ID,nots=" & TxtTransSerial.text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & "  From Transactions Where Transaction_ID =" & XPTxtBillID.text + " And Transaction_Type = 26"
         sql = "INSERT INTO  Transactions ( BranchId,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots2,NoteSerial,NoteSerial1,NoteId,Transaction_Type_Sub,WorkOrderNO)SELECT   " & val(Dcbranch.BoundText) & "," & Transaction_ID & "," & mytext & ",Transaction_Date,Transaction_Type = 27,CusID,StoreID1,UserID,Emp_ID,nots2=" & TxtTransSerial.Text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & " ,Transaction_Type_Sub=27 , " & TxtTransSerial.Text & " From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 26"

         Cn.Execute sql
            
            rs!nots2 = Transaction_ID
        rs!Product_Issue_voucher_Serial = TxtNoteSerial1V
        rs.update
        Txtnots2.Text = Transaction_ID
        
        'fill transaction details table
 
        Set RSTransDetails = New ADODB.Recordset
      '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
        For RowNum = 1 To FG1.Rows - 1

            If FG1.TextMatrix(RowNum, FG1.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                
         
        
                RSTransDetails("Item_ID").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("id")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("id"))))
                 RSTransDetails("Quantity").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty"))))
                RSTransDetails("SHOWQTY").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty"))))
               RSTransDetails("showPrice").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("Cost")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("Cost"))))
             RSTransDetails("UnitID").value = IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("unitid")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("unitid"))))
     
          
                          '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG1.TextMatrix(RowNum, FG1.ColIndex("id")))
            LngUnitID = val(FG1.TextMatrix(RowNum, FG1.ColIndex("unitid")))  ' val(Fg1.Cell(flexcpData, RowNum, Fg1.ColIndex("unitid")))
            DblQty = val(FG1.TextMatrix(RowNum, FG1.ColIndex("TotalQty")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
             '   RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
             '   RSTransDetails("OpeningSalesValue").value = IIf((Fg1.TextMatrix(RowNum, Fg1.ColIndex("Valu")) = ""), Null, val(Fg1.TextMatrix(RowNum, Fg1.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG1.TextMatrix(RowNum, FG1.ColIndex("Cost")) = ""), Null, val(FG1.TextMatrix(RowNum, FG1.ColIndex("Cost"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If
            
                        Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName2.BoundText), Transaction_ID, OldQty, OldCost, NewQty, NewCost
'       RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
'       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
'
       
         
             
                RSTransDetails.update
            End If

        Next RowNum
       UpdateTransactionsCost CStr(Transaction_ID)
        '       Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT round(showPrice + ToTAlELSHahn/ShowQty,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
'  Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,itemsize,UnitId,ShowQty,QtyBySmalltUnit,order_no,classid)SELECT   (showPrice) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (Price ), ColorID,itemsize, UnitId, ShowQty, QtyBySmalltUnit,order_no,classid From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        
        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «‰‘«¡ «·”‰œ"
       Else
       MsgBox "Create Successfully"
       End If

    End If
 
    '
 
ErrTrap:

End Sub
Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
With FG
Select Case .ColKey(.Col)
Case "MixNo"
Unload FrmSearchDevComItem
FrmSearchDevComItem.lbltype = 4
FrmSearchDevComItem.show
End Select
End With
End If
End If
End Sub
Function CREATE_VOUCHER_GE2(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, Optional Row As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Variant

    TOTAL_COST = val(FG.TextMatrix(Row, FG.ColIndex("Valu")))
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = TOTAL_COST

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName2.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
         '   Dim groupAccount As String
         '
         '   Dim line_value As Variant
'
'            With FG1
'
'                For i = 1 To FG1.Rows - 1
'
'                    If FG1.TextMatrix(i, FG1.ColIndex("Code")) <> "" Then
'
'                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
'                        groupAccount = get_item_group_account_inventory(FG1.TextMatrix(i, FG1.ColIndex("id")), DCboStoreName2.BoundText, 0)
'
'                        If groupAccount = "Error" Then
'                            If SystemOptions.UserInterface = ArabicInterface Then
'                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
'                            Else
'                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
'                            End If
'
'                            GoTo ErrTrap
'                        End If
'
                        '         line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
'                        line_value = val(FG1.TextMatrix(i, FG1.ColIndex("total")))
'
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã " & Me.TxtTransSerial.Text
'                        Else
'                            StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
'                        End If
'
'                        LngDevNO = LngDevNO + 1
'
'                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
'                            GoTo ErrTrap
'                        End If
'
'                    End If
'
'                Next i
'
      '      End With

        End If

        '«·ÿ—› «·„œÌ‰
   '     SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
 SngTemp = TOTAL_COST
        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«œ    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
                Else
                    StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
      
        End If
    End If

ErrTrap:
End Function
Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Variant
    '   With FG

    '             For i = 1 To FG.Rows - 1
    '
    '                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    '                    TOTAL_COST = TOTAL_COST + ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)
    '                    End If
    '                Next i
    '     End With
    TOTAL_COST = val(TxtTotalMaterials.Text)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = TOTAL_COST

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName2.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

            With FG1

                For i = 1 To FG1.Rows - 1

                    If FG1.TextMatrix(i, FG1.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG1.TextMatrix(i, FG1.ColIndex("id")), DCboStoreName2.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        '         line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
                        line_value = val(FG1.TextMatrix(i, FG1.ColIndex("total")))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã " & Me.TxtTransSerial.Text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
   '     SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
 SngTemp = TOTAL_COST
        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«œ    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  ÕœÌœ „’«—Ì› «‰ «Ã «·„Ê«
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "»‰«¡ ⁄·Ï «„— «‰ «Ã" & Me.TxtTransSerial.Text
                Else
                    StrTempDes = "Issue Voucher No. " & Me.TxtTransSerial.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
      
        End If
    End If

ErrTrap:
End Function

Sub CheckAccounts()
    Dim SngTemp  As Variant
    Dim Vchr_result As String
    Dim notes_result As String
    Dim Account_Code_dynamic As String
    '«·ÿ—› «·„œÌ‰
    '  SngTemp = NewGrid.GetItemsTotal(5)
    SngTemp = Round(val(TXTTotalIssueVouchers.Text), 2) + Round(val(TxtworkerTotal), 2) + Round(val(TXTLineExpenses.Text), 2) + Round(val(Txt_EXport.Text), 2) + Round(val(TXTFinacilaTotal.Text), 2) + Round(val(TXTFactoryExpenses.Text), 2) + Round(val(TxtTotalEstimatedCost.Text), 2) + Round(val(TxtIndirectCostForProduction.Text), 2)

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
                            
                End If
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
 
        End If
    End If
     
    Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28)

    If Vchr_result = "error" Then
        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
    Else
                       
        If Vchr_result = "" Then
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        Else
  
        End If
    End If
                    
    notes_result = Notes_coding(val(my_branch), ReciveDate.value)

    If notes_result = "error" Then
        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
    Else
                       
        If notes_result = "" Then
            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
        Else
                        
        End If
    End If

    Exit Sub
ErrTrap:
       
End Sub
Sub CreateRecevVoucher(Optional Row As Long)
    'On Error GoTo errortrap
 '   autoedit = 1
 '   Cmd_Click (1)
'autoedit = 0
    DoEvents
 '   Cmd_Click (2)

    DoEvents


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
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim mytext As String


    If rs.RecordCount = 0 Then MsgBox "«Õ›Ÿ «„— «·«‰ «Ã «Ê·«": Exit Sub
    If SystemOptions.UserInterface = ArabicInterface Then
        
        Msg = "”Ê› Ì „ «‰‘«¡  ”‰œ  «÷«›…     .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        
    Else
        Msg = "Create Recieve Voucher to this bill ?"
    End If

    ' On Error GoTo ErrTrap

    If MsgBox(Msg, vbYesNo, App.title) = vbYes Then

        Dim Transaction_ID As Long
        

        'set rs!Transaction_Serial=  where Transaction_Type=20
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
    
        TxtNoteSerial1V = ""
        TxtNoteSerialV = ""
   
        my_branch = val(Me.Dcbranch.BoundText)
        Dim NoteSerial As String
        Dim Vchr_result As String
        Dim notes_result As String
         DeleteTransactiomsVoucher val(FG.TextMatrix(Row, FG.ColIndex("ReceivTransID")))
        If TxtresiveVoucher = "" Then
      
            If TxtNoteSerial1V = "" Then
                Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28, , val(DCboStoreName.BoundText))

                If Vchr_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                       
                    If Vchr_result = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        TxtNoteSerial1V = Vchr_result
                    End If
                End If
            End If
                    
            If TxtNoteSerialV = "" Then
                notes_result = Notes_coding(val(my_branch), ReciveDate.value)

                If notes_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                       
                    If notes_result = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        TxtNoteSerialV = notes_result
                    End If
                End If
            End If
        
         '   DeleteTransactiomsVoucher val(Text1.text)
            'TxtresiveVoucher = TxtNoteSerial1V
            
            
        Else 'Õ«·… «· ⁄œÌ·
    
            TxtNoteSerial1V = TxtresiveVoucher
            TxtNoteSerialV = get_transaction_NoteSerial2(val(FG.TextMatrix(Row, FG.ColIndex("ReceivTransID"))))

            If Trim(TxtNoteSerialV) = "" Then
                TxtNoteSerialV = Notes_coding(val(my_branch), ReciveDate.value)
            End If
    
         '   DeleteTransactiomsVoucher val(Text1.text)
    
        End If

        mytext = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=28"))
         FG.TextMatrix(Row, FG.ColIndex("ReceiveSerial")) = TxtNoteSerial1V
        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
       ' RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
 
        general_noteid = CStr(new_id("Notes", "NoteID", "", True))
      
        
       
       ' If FG.TextMatrix(Row, FG.ColIndex("ReceivTransID")) = "" Then FG.TextMatrix(Row, FG.ColIndex("ReceivTransID")) = 0
     Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
         Cn.Execute "INSERT INTO  Transactions (order_no,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,Transaction_Type_Sub,WorkOrderNO,BranchId)SELECT '" & txt_ORDER_NO.Text & "'," & Transaction_ID & "," & mytext & "," & SQLDate(ReciveDate.value, True) & ",Transaction_Type = 28,CusID,StoreID,UserID,Emp_ID,nots='" & TxtTransSerial.Text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",Transaction_Type_Sub=28,Transaction_Serial,BranchId From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 26"
        '
        'Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT round(showPrice + ToTAlELSHahn/ShowQty,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, Price*rate+ToTAlELSHahn, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,itemsize,UnitId,ShowQty,QtyBySmalltUnit,order_no,classid,OldQty,OldCost,NewQty,NewCost )SELECT   (showPrice+TotalPriceNoHours/ShowQty) as showPrice ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (Price +TotalPriceNoHours/ShowQty) as Price, ColorID,itemsize, UnitId, ShowQty, QtyBySmalltUnit,order_no,classid,OldQty,OldCost,NewQty,NewCost From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text & " and Item_ID =" & val(FG.TextMatrix(Row, FG.ColIndex("code"))) & " and MixNo='" & (FG.TextMatrix(Row, FG.ColIndex("MixNo"))) & "'"
        
       rs!nots = Transaction_ID
        rs!Product_Receive_voucher_Serial = TxtNoteSerial1V
        rs.update
        FG.TextMatrix(Row, FG.ColIndex("ReceivTransID")) = Transaction_ID
                  
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = general_noteid ' CStr(new_id("Notes", "NoteID", "", True))
        'general_noteid = RsNotesGeneral("NoteID").value
        FG.TextMatrix(Row, FG.ColIndex("ResiveNoteID")) = general_noteid
        'TXTNoteID.Text = general_noteid
        
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = ReciveDate.value
        RsNotesGeneral("Branch_no").value = val(Me.Dcbranch.BoundText)
         
        RsNotesGeneral("NoteType").value = 250
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
'        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))

        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(19) '«–‰ «÷«›…
        RsNotesGeneral("sanad_year").value = year(ReciveDate.value)
        RsNotesGeneral("sanad_month").value = Month(ReciveDate.value)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        'Cn.Execute "update Transactions Set Transaction_Serial = Transaction_Serial Where Transaction_Type = 20"
       FG.TextMatrix(Row, FG.ColIndex("ReceivTransID")) = Transaction_ID
       Cn.Execute "update Transaction_Details set  ReceiveSerial='" & (FG.TextMatrix(Row, FG.ColIndex("ReceiveSerial"))) & "',ReceivTransID=" & val(FG.TextMatrix(Row, FG.ColIndex("ReceivTransID"))) & " ,ResiveNoteID=" & val(FG.TextMatrix(Row, FG.ColIndex("ResiveNoteID"))) & " where Transaction_ID=" & val(XPTxtBillID.Text) & " and Item_ID =" & val(FG.TextMatrix(Row, FG.ColIndex("code"))) & " and MixNo='" & (FG.TextMatrix(Row, FG.ColIndex("MixNo"))) & "' "
  
         UpdateTransactionsCost CStr(Transaction_ID)
        CREATE_VOUCHER_GE12 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, Row
rs.Resync
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·”‰œ"
        Else
            MsgBox " Vouchers Created "
        End If
    End If


  ' Me.Retrive val(Me.XPTxtBillID.Text)
    '----------------------------------------------------------------
TxtresiveVoucher = ""
ErrTrap:
End Sub
Private Sub CmdResiveVoucher_Click()
    'On Error GoTo errortrap
    autoedit = 1
    Cmd_Click (1)
autoedit = 0
    DoEvents
    Cmd_Click (2)

    DoEvents

    'If TxtresiveVoucher <> "" Then MsgBox " „ «‰‘«¡ ”‰œ «·«” ·«„ „‰ ﬁ»· ": Exit Sub
    cal_expenses
    'DeleteTransactiomsVoucher Val(Text1.text)

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
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim mytext As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    'rs.Close
    '
    '        rs.Open "select * from Transactions where nots = " & TxtTransSerial.text & " and Transaction_type = 20"
    '          If rs.RecordCount > 0 Then
    '        If rs!nots <> "" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '             Msg = "·ﬁœ  „ ⁄„·   «–‰ «÷«›…    .."
    '             Msg = " »«·«–‰ —ﬁ„ " + Me.TxtresiveVoucher & Chr(13)
    '            Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
    '        Else
    '          Msg = "This bill already converted" & Chr(13)
    '          Msg = Msg + " Voucher No " + Me.TxtresiveVoucher & Chr(13)
    '        End If
    '          MsgBox Msg, vbOKOnly, App.Title
    '        Exit Sub
    '        End If
    '        End If

 '   rs.Close
 '   rs.Open "select * from Transactions where Transaction_Serial = " & TxtTransSerial.text & " and Transaction_type = 26"

    If rs.RecordCount = 0 Then MsgBox "«Õ›Ÿ «„— «·«‰ «Ã «Ê·«": Exit Sub
    If SystemOptions.UserInterface = ArabicInterface Then
        
        Msg = "”Ê› Ì „ «‰‘«¡  ”‰œ  «÷«›…     .."
        Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        
    Else
        Msg = "Create Recieve Voucher to this bill ?"
    End If

    ' On Error GoTo ErrTrap

    If MsgBox(Msg, vbYesNo, App.title) = vbYes Then

        Dim Transaction_ID As Long
        

        'set rs!Transaction_Serial=  where Transaction_Type=20
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
    
        TxtNoteSerial1V = ""
        TxtNoteSerialV = ""
   
        my_branch = val(Me.Dcbranch.BoundText)
        Dim NoteSerial As String
        Dim Vchr_result As String
        Dim notes_result As String
         DeleteTransactiomsVoucher val(Text1.Text)

        If TxtresiveVoucher = "" Then
      
            If TxtNoteSerial1V = "" Then
                Vchr_result = Voucher_coding(val(my_branch), ReciveDate.value, 19, 250, , 28, , val(DCboStoreName.BoundText))

                If Vchr_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «” ·«„ „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                       
                    If Vchr_result = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        TxtNoteSerial1V = Vchr_result
                    End If
                End If
            End If
                    
            If TxtNoteSerialV = "" Then
                notes_result = Notes_coding(val(my_branch), ReciveDate.value)

                If notes_result = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                       
                    If notes_result = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        TxtNoteSerialV = notes_result
                    End If
                End If
            End If
        
         '   DeleteTransactiomsVoucher val(Text1.text)
            TxtresiveVoucher = TxtNoteSerial1V
        Else 'Õ«·… «· ⁄œÌ·
    
            TxtNoteSerial1V = TxtresiveVoucher
            TxtNoteSerialV = get_transaction_NoteSerial2(val(Text1.Text))

            If Trim(TxtNoteSerialV) = "" Then
                TxtNoteSerialV = Notes_coding(val(my_branch), ReciveDate.value)
            End If
    
         '   DeleteTransactiomsVoucher val(Text1.text)
    
        End If

        mytext = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=28"))
         
        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
       ' RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
 
        general_noteid = CStr(new_id("Notes", "NoteID", "", True))
      
        
       
        If txt_ORDER_NO.Text = "" Then txt_ORDER_NO.Text = 0
     Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
         Cn.Execute "INSERT INTO  Transactions (order_no,Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,Transaction_Type_Sub,WorkOrderNO,BranchId)SELECT '" & txt_ORDER_NO.Text & "'," & Transaction_ID & "," & mytext & "," & SQLDate(ReciveDate.value, True) & ",Transaction_Type = 28,CusID,StoreID,UserID,Emp_ID,nots='" & TxtTransSerial.Text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",Transaction_Type_Sub=28,Transaction_Serial,BranchId From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 26"
        '
        'Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT round(showPrice + ToTAlELSHahn/ShowQty,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, Price*rate+ToTAlELSHahn, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,itemsize,UnitId,ShowQty,QtyBySmalltUnit,order_no,classid,OldQty,OldCost,NewQty,NewCost )SELECT   (showPrice) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (Price ), ColorID,itemsize, UnitId, ShowQty, QtyBySmalltUnit,order_no,classid,OldQty,OldCost,NewQty,NewCost From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text
        
       rs!nots = Transaction_ID
        rs!Product_Receive_voucher_Serial = TxtNoteSerial1V
        rs.update
                   UpdateTransactionsCost CStr(Transaction_ID)
                   
      
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = general_noteid ' CStr(new_id("Notes", "NoteID", "", True))
        'general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.Text = general_noteid
        
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = ReciveDate.value
        RsNotesGeneral("Branch_no").value = val(Me.Dcbranch.BoundText)
         
        RsNotesGeneral("NoteType").value = 250
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
'        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))

        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(19) '«–‰ «÷«›…
        RsNotesGeneral("sanad_year").value = year(ReciveDate.value)
        RsNotesGeneral("sanad_month").value = Month(ReciveDate.value)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        'Cn.Execute "update Transactions Set Transaction_Serial = Transaction_Serial Where Transaction_Type = 20"

        CREATE_VOUCHER_GE1 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·”‰œ"
        Else
            MsgBox " Vouchers Created "
        End If
    End If

    'Transaction_ID

    '----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
 '   StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=26"

 '   Set rs = New ADODB.Recordset
 '   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   Me.Retrive val(Me.XPTxtBillID.Text)
    '----------------------------------------------------------------
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
    'If Text1.text <> "" Then
    '    Msg = " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»· Ê·« Ì„ﬂ‰  ÕÊÌ·Â«  "
    '            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    'On Error GoTo ErrTrap
    'Screen.MousePointer = vbArrowHourglass
    '    Set Frm = New FrmInpout
    'With Frm
    '    .Convert
    ''    .XPTxtBillID.Text = XPTxtBillID.Text
    '    .XPDtbBill.Value = XPDtbBill.Value
    '    .DBCboClientName.BoundText = DBCboClientName.BoundText
    '    .DCboStoreName.BoundText = DCboStoreName.BoundText
    '    .CboPayMentType.ListIndex = CboPayMentType.ListIndex
    '    .Text1.text = TxtTransSerial.text
    '    .Text2.text = XPTxtBillID.text
    '
    '
    '    For RowNum = 1 To FG.Rows - 1
    '        If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
    '           .FG.Rows = .FG.Rows + 1
    '        End If
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
    ''        Dim StrSQL As String
    '        Dim RsUnit As New ADODB.Recordset
    'StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 22) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
    'Set RsUnit = New ADODB.Recordset
    'RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '
    '
    '
    '        .FG.Cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
    '        .FG.TextMatrix(RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))
    '         Rs!nots = TxtTransSerial.text
    '         Rs.update
    '
    '
    ''        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
    ''        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
    ''           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
    ''        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
    ''        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))
    '
    '    Next RowNum
    '    .Cala
    'End With
    'Screen.MousePointer = vbDefault
    'Cmd_Click (2)
    'Frm.Hide
    'Exit Sub
    'errortrap:
    'Screen.MousePointer = vbDefault
    'MsgBox " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»·", vbCritical
ErrTrap:

End Sub
Function CREATE_VOUCHER_GE12(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, Optional Row As Long)
    Dim LngDevID As Long
    Dim NoHours As Double
    Dim TempValue As Double
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Variant

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    '  SngTemp = NewGrid.GetItemsTotal(5)
    SngTemp = val(FG.TextMatrix(Row, FG.ColIndex("Valu"))) + val(val(FG.TextMatrix(Row, FG.ColIndex("TotalPriceNoHours"))))

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

        

        End If

        '«·ÿ—› «·œ«∆‰
        '   SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) '* Val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)
        SngTemp = val(FG.TextMatrix(Row, FG.ColIndex("Valu")))
        NoHours = val(FG.TextMatrix(Row, FG.ColIndex("NoHours")))

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „Ê«œ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                TempValue = val(TxtUsedPowerPriceHTotal.Text) * NoHours
                If NoHours > 0 Then
                    StrTempAccountCode = get_account_code_branch(39, my_branch) '  „’«—Ì›  ÊﬁÊœ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TempValue, 1, StrTempDes & "Õ”«» „’«—Ì› «·ÊﬁÊœ", general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                End If
                
                 TempValue = val(TxtworkerTotalPerHour.Text) * NoHours
                If NoHours > 0 Then
                    StrTempAccountCode = get_account_code_branch(38, my_branch) '  „’«—Ì›  «ÃÊ—
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TempValue, 1, StrTempDes & "Õ”«» «ÃÊ— ", general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                End If
                
                 TempValue = val(TxtUsedElectricPriceHTotal.Text) * NoHours
                If NoHours > 0 Then
                    StrTempAccountCode = get_account_code_branch(79, my_branch) '  „’«—Ì›  ﬂÂ—»«¡
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TempValue, 1, StrTempDes & "Õ”«» „’«—Ì› «·ﬂÂ—»«¡", general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                End If
                TempValue = val(TxtHourdippTotal.Text) * NoHours
                If NoHours > 0 Then
                    StrTempAccountCode = get_account_code_branch(151, my_branch) '  „’«—Ì›  «·«Â·«ﬂ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TempValue, 1, StrTempDes & "Õ”«» „’«—Ì› «·«Â·«ﬂ", general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                End If

            End If
        End If
        
        'ﬁÌœ «ÃÊ— «·⁄„«·
      '  SngTemp = Round(val(TxtworkerTotal), 2)
'
'        If SngTemp > 0 Then
'            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then
'
'                Account_Code_dynamic = get_account_code_branch(38, my_branch)
'
'                If Account_Code_dynamic = "NO branch" Then
'                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                    GoTo ErrTrap
'                Else
'
'                    If Account_Code_dynamic = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã «ÃÊ— ⁄„«·… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
''                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã «ÃÊ— ⁄„«·…
 '
 '               If SystemOptions.UserInterface = ArabicInterface Then
 '                   StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & Chr(13) & txtRemark.Text
 '               Else
 '                   StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & Chr(13) & txtRemark.Text
 '               End If
 ''
  '              LngDevNO = LngDevNO + 1
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'
'            End If
'        End If

        'ﬁÌœ „’—Ê›«  ’‰«⁄Ì…
'        SngTemp = Round(val(TXTFactoryExpenses.Text), 2) + Round(val(TXTLineExpenses.Text), 2) + Round(val(TxtIndirectCostForProduction.Text), 2)

 '       If SngTemp > 0 Then
 '           If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then
'
'                Account_Code_dynamic = get_account_code_branch(39, my_branch)
'
'                If Account_Code_dynamic = "NO branch" Then
'                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                    GoTo ErrTrap
'                Else
'
'                    If Account_Code_dynamic = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã , „’—Ê›«  ’‰«⁄Ì… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „’—Ê›«  ’‰«⁄Ì…
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & Chr(13) & txtRemark.Text
'                Else
'                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & Chr(13) & txtRemark.Text
'                End If
'
'                LngDevNO = LngDevNO + 1
'
 '               If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
 '                   GoTo ErrTrap
 '               End If
'
'            End If
'        End If
 
  '      'ﬁÌœ «·„’—Ê›« 
  '      Dim Account_Code As String
  '      Dim Note_Value As Variant
'
'        For i = 1 To Grid.Rows - 1
'
'            If val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = 1 Or val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = -1 Then
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & Chr(13) & txtRemark.Text
'                Else
'                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & Chr(13) & txtRemark.Text
'                End If
'
'                LngDevNO = LngDevNO + 1
'                Account_Code = Grid.TextMatrix(i, Grid.ColIndex("Account_code"))
'                Note_Value = Round(Grid.TextMatrix(i, Grid.ColIndex("Note_value")), 2)
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'        Next
'
'        'ﬁÌœ «·›Ê« Ì—
'        For i = 1 To grid4.Rows - 1
'
'            If val(grid4.TextMatrix(i, grid4.ColIndex("select"))) = 1 Or val(grid4.TextMatrix(i, grid4.ColIndex("select"))) = -1 Then
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & Chr(13) & txtRemark.Text
'                Else
'                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & Chr(13) & txtRemark.Text
'                End If
'
'                LngDevNO = LngDevNO + 1
'                Account_Code = grid4.TextMatrix(i, grid4.ColIndex("Account_code"))
'                Note_Value = Round(grid4.TextMatrix(i, grid4.ColIndex("Note_value")), 2)
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'        Next
'
'        ' «·ﬁÌœ «·„’—Ê›«  «· ﬁœÌ—Ì…
'
'        Dim LineDes As String
'
'        For i = 1 To GridEstimatedCost.Rows - 1
'
'            If (GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountCode"))) <> "" Then
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & Chr(13) & txtRemark.Text
'                Else
'                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & Chr(13) & txtRemark.Text
'                End If
'
'                LngDevNO = LngDevNO + 1
'                LineDes = GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountName"))
'                Account_Code = GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountCode"))
'                Note_Value = Round(GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("Total")), 2)
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes + LineDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'        Next
        
    End If

ErrTrap:
End Function
Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Variant

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    '  SngTemp = NewGrid.GetItemsTotal(5)
    SngTemp = Round(val(TXTTotalIssueVouchers.Text), 2) + Round(val(TxtworkerTotal), 2) + Round(val(TXTLineExpenses.Text), 2) + Round(val(Txt_EXport.Text), 2) + Round(val(TXTFinacilaTotal.Text), 2) + Round(val(TXTFactoryExpenses.Text), 2) + Round(val(TxtTotalEstimatedCost.Text), 2) + Round(val(TxtIndirectCostForProduction.Text), 2)

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Variant

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        'total_shahn = Round((line_value) / Val(LblTotal.Caption), 2)       'ﬁÌ„… «Ã„«·Ì  ”ÿ— »«·„’—Ê›« 
                        line_value = line_value + val(FG.TextMatrix(i, FG.ColIndex("Expenses"))) * FG.TextMatrix(i, FG.ColIndex("Count"))
                        line_value = Round(line_value, 0)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, Round(line_value, 0), 0, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        '   SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) '* Val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)
        SngTemp = Round(val(TXTTotalIssueVouchers.Text), 2)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „Ê«œ
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
        End If
        
        'ﬁÌœ «ÃÊ— «·⁄„«·
        SngTemp = Round(val(TxtworkerTotal), 2)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(38, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã «ÃÊ— ⁄„«·… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã «ÃÊ— ⁄„«·…
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
        End If

        'ﬁÌœ „’—Ê›«  ’‰«⁄Ì…
        SngTemp = Round(val(TXTFactoryExpenses.Text), 2) + Round(val(TXTLineExpenses.Text), 2) + Round(val(TxtIndirectCostForProduction.Text), 2)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then

                Account_Code_dynamic = get_account_code_branch(39, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã , „’—Ê›«  ’‰«⁄Ì… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '  „’«—Ì› «·«‰ «Ã „’—Ê›«  ’‰«⁄Ì…
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            End If
        End If
 
        'ﬁÌœ «·„’—Ê›« 
        Dim Account_code As String
        Dim Note_Value As Variant

        For i = 1 To Grid.Rows - 1

            If val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = 1 Or val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = -1 Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1
                Account_code = Grid.TextMatrix(i, Grid.ColIndex("Account_code"))
                Note_Value = Round(Grid.TextMatrix(i, Grid.ColIndex("Note_value")), 2)

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
        
        Next

        'ﬁÌœ «·›Ê« Ì—
        For i = 1 To Grid4.Rows - 1

            If val(Grid4.TextMatrix(i, Grid4.ColIndex("select"))) = 1 Or val(Grid4.TextMatrix(i, Grid4.ColIndex("select"))) = -1 Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1
                Account_code = Grid4.TextMatrix(i, Grid4.ColIndex("Account_code"))
                Note_Value = Round(Grid4.TextMatrix(i, Grid4.ColIndex("Note_value")), 2)

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
        
        Next
 
        ' «·ﬁÌœ «·„’—Ê›«  «· ﬁœÌ—Ì…
  
        Dim LineDes As String

        For i = 1 To GridEstimatedCost.Rows - 1

            If (GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountCode"))) <> "" Then
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " „‰  «„— «‰ «Ã —ﬁ„" & TxtTransSerial & CHR(13) & txtRemark.Text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & TxtNoteSerial1V & TxtNoteSerial1V & " From PO NO:" & TxtTransSerial & CHR(13) & txtRemark.Text
                End If
            
                LngDevNO = LngDevNO + 1
                LineDes = GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountName"))
                Account_code = GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("AccountCode"))
                Note_Value = Round(GridEstimatedCost.TextMatrix(i, GridEstimatedCost.ColIndex("Total")), 2)

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes + LineDes, general_noteid, , , , Me.ReciveDate.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
        
        Next
        
    End If

ErrTrap:
End Function

Private Sub CmdTemplate_Click()
    Dim Frm  As FrmBuySearch
    On Error GoTo ErrTrap
    Set Frm = New FrmBuySearch

    With Frm
        .DealingForm = InsertTemplate
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        '    .MDIChild = True
        .BorderStyle = 0
        '  .MinButton = True
        .show vbModeless, mdifrmmain
        .Visible = True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Command3_Click()

    Dim Transaction_ID As Double
    Transaction_ID = get_transaction_id(TxtIssueSerial, 27, 27)

    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmOutProductionOrder.show
    FrmOutProductionOrder.Retrive (Transaction_ID)
End Sub

Public Function get_transaction_NoteSerial2(Transaction_ID As Long) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where Transaction_ID=" & Transaction_ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_NoteSerial2 = ""
    Else
        get_transaction_NoteSerial2 = IIf(IsNull(rs("NoteSerial").value), 0, rs("NoteSerial").value)
    End If

End Function

Public Function get_transaction_NoteSerial(NoteSerial1 As String, _
                                           Transaction_Type As Integer, _
                                           Transaction_Type_Sub As Integer) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where NoteSerial1='" & NoteSerial1 & "' and  Transaction_Type= " & Transaction_Type & " And Transaction_Type_Sub = " & Transaction_Type_Sub
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_NoteSerial = ""
    Else
        get_transaction_NoteSerial = IIf(IsNull(rs("NoteSerial").value), 0, rs("NoteSerial").value)
    End If

End Function

Public Function get_transaction_id(NoteSerial1 As String, _
                                   Transaction_Type As Integer, _
                                   Transaction_Type_Sub As Integer) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select Transaction_ID,Transaction_Type,NoteSerial1,Transaction_Type_Sub from Transactions where NoteSerial1='" & NoteSerial1 & "' and  Transaction_Type= " & Transaction_Type '& " And Transaction_Type_Sub = " & Transaction_Type_Sub
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_id = 0
    Else
        get_transaction_id = IIf(IsNull(rs("Transaction_ID").value), 0, rs("Transaction_ID").value)
    End If

End Function

Private Sub Command4_Click()
    Dim Transaction_ID As Double
    Transaction_ID = get_transaction_id(TxtresiveVoucher, 28, 28)

    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmInpoutWorkOrder.show
    FrmInpoutWorkOrder.Retrive (Transaction_ID)
End Sub

Private Sub Command5_Click()
    Dim NoteSerial As String
    NoteSerial = get_transaction_NoteSerial(TxtIssueSerial, 27, 27)

    If NoteSerial = "" Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)

End Sub

Private Sub Command7_Click()
    Dim NoteSerial As String
    NoteSerial = get_transaction_NoteSerial2(val(Text1.Text))

    If val(NoteSerial) = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
 
    End If
        
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 11815
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 11815
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID1.Text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(Dcbranch.BoundText), 49) = True Then
     TxtTransSerial.Text = ""
    
     End If
     
    End If
    
    
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        
        Dcombos.GetStores Me.DCboStoreName
 
    End If

End Sub

Private Sub DCboStoreName2_Change()
 TxtStoreID1.Text = getStoreCoding(val(DCboStoreName2.BoundText))
 
End Sub

Private Sub DCboStoreName2_KeyUp(KeyCode As Integer, _
                                 Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
 
        Dcombos.GetStores Me.DCboStoreName2
 
    End If

End Sub

Private Sub Dcbranch_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
       TxtTransSerial.Text = ""
    End If
    
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches Dcbranch
    End If

End Sub

Private Sub DcLine_KeyUp(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
 
        Dcombos.GetLine Me.DcLine
 
    End If

End Sub

Private Sub dcShift_Click(Area As Integer)
    Dim sql As String
    Dim rsshift As New ADODB.Recordset
    sql = "select * from TbLSheft where SeftCode=" & val(DcShift.BoundText)
    rsshift.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsshift.RecordCount > 0 Then
        DTFrom.value = rsshift("ShiftFrom").value
        DTTo = rsshift("ShiftTo").value
    End If

    Shifttime.Text = CalculateTimes(Me.DTFrom.value, Me.DTTo.value)
End Sub

Private Sub DTfrom_Change()
    Shifttime.Text = CalculateTimes(Me.DTFrom.value, Me.DTTo.value)
End Sub

Private Sub DtTo_Change()
    Shifttime.Text = CalculateTimes(Me.DTFrom.value, Me.DTTo.value)
End Sub

Private Sub Ele_Click(Index As Integer)

    Select Case Index

        Case 6
            On Error GoTo ErrTrap
            '        If Me.WindowState = vbNormal Then
            '            Me.WindowState = vbMaximized
            '        Else
            '            Me.WindowState = vbNormal
            '        End If
    End Select

    Exit Sub
ErrTrap:
End Sub

Function FillExp()
    'Dim RowNum As Integer
    'Dim unitid As Integer
    '    For RowNum = 1 To FG.Rows - 1
    '        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
    '
    '             unitid = _
    '         IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
    '       End If
    '    Next RowNum
    
    'FillUnitExpenses unitid

End Function

Private Sub FG_AfterDataRefresh()
    'Dim unitid As Integer
    'show_parts
    'FIllEstimatedExpenses

    ' FillExp
End Sub
 
Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    'If Col = 11 Then
    '   With FG
    show_parts
        
    FIllEstimatedExpenses
    cal_expenses
    FG.TextMatrix(Row, FG.ColIndex("PriceNoHours")) = val(TxtHourdippTotal.Text) + val(TxtUsedElectricPriceHTotal.Text) + val(TxtUsedPowerPriceHTotal.Text) + val(TxtworkerTotalPerHour.Text)
    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , , Me.txt_ORDER_NO
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), , Me.txt_ORDER_NO

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    
    '   End With
    'End If
End Sub
Function CheckAccount() As Boolean
CheckAccount = False
    my_branch = val(Me.Dcbranch.BoundText)
   Dim Account_Code_dynamic82 As String
         If val(TxtworkerTotalPerHour.Text) <> 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(38, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·«ÃÊ—", vbCritical
                                                            Else
                                                                MsgBox "Please Select Account Wages", vbCritical
                                                            End If
                                                           CheckAccount = True
                            Exit Function
                                                
                              End If
             End If
            If val(TxtUsedPowerPriceHTotal.Text) <> 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(39, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·ÊﬁÊœ", vbCritical
                                                            Else
                                                                MsgBox "Please Select Account Fuel", vbCritical
                                                            End If
                                                           CheckAccount = True
                            Exit Function
                                               
                              End If
             End If
                If val(TxtUsedElectricPriceHTotal.Text) <> 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(79, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·ﬂÂ—»«¡", vbCritical
                                                            Else
                                                                MsgBox "Please Select  Account of Electricity", vbCritical
                                                            End If
                                                           CheckAccount = True
                            Exit Function
                                                
                              End If
             End If
                          If val(TxtHourdippTotal.Text) <> 0 Then
                            Account_Code_dynamic82 = get_account_code_branch(151, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·«Â·«ﬂ", vbCritical
                                                            Else
                                                                MsgBox "Please Select  Account of Depreciation", vbCritical
                                                            End If
                                                           CheckAccount = True
                            Exit Function
                                                
                              End If
             End If
End Function
Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)
       Dim Transaction_ID As Double
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
        '   FrmAddNewItem.DealingForm = ShowPrice
        '   FrmAddNewItem.Show vbModal
        Else
        Dim NoteSerial As String
        
 With FG
   Select Case .ColKey(Col)
   Case "Voucher"
    CreateIssueVoucher Row
  Case "VoucheRecev"
  If CheckAccount() = True Then Exit Sub
    CreateRecevVoucher Row
  Case "ShowIssue"
    Transaction_ID = get_transaction_id(.TextMatrix(Row, .ColIndex("IssueSerial")), 27, 27)
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmOutProductionOrder.show
    FrmOutProductionOrder.Retrive (Transaction_ID)
  Case "IssuGl"
      
     NoteSerial = get_transaction_NoteSerial(.TextMatrix(Row, .ColIndex("IssueSerial")), 27, 27)
     If NoteSerial = "" Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
     FrmAccEditJournal.show
     FrmAccEditJournal.Retrive (NoteSerial)
 Case "ShowReceiv"
     
    Transaction_ID = get_transaction_id(.TextMatrix(Row, .ColIndex("ReceiveSerial")), 28, 28)
    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmInpoutWorkOrder.show
      FrmInpoutWorkOrder.Retrive (Transaction_ID)
 Case "RecevGl"
     NoteSerial = get_transaction_NoteSerial2(val(.TextMatrix(Row, .ColIndex("ReceivTransID"))))
    If val(NoteSerial) = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmAccEditJournal.show
    FrmAccEditJournal.Retrive (NoteSerial)
   End Select
 End With
      
    End If

End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    With GridWorker

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("code")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub FG_CellChanged(ByVal Row As Long, _
                           ByVal Col As Long)
    'On Error Resume Next
    'If Col = 11 Then
    '        With FG
    '        show_parts
    '       FIllEstimatedExpenses
    '
    '        End With
    'End If

End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value"
                Dim sgl As String
    
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        If .Rows > 1 Then
            TXTFactoryExpenses = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows, .ColIndex("value"))
        Else
            TXTFactoryExpenses = 0
        End If

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    cal_expenses
    ReLineGrid
End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                    Order_no_search.show
                    Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 3
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "AccountName"
                 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts"
                Else
                    StrSQL = "select * from Expenses_accounts_eng "
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
              
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "VoucheRecev"
.ColComboList(.ColIndex("VoucheRecev")) = "..."
Case "Voucher"
.ColComboList(.ColIndex("Voucher")) = "..."
Case "ShowIssue"
.ColComboList(.ColIndex("ShowIssue")) = "..."
Case "IssuGl"
.ColComboList(.ColIndex("IssuGl")) = "..."
Case "ShowReceiv"
.ColComboList(.ColIndex("ShowReceiv")) = "..."
Case "RecevGl"
.ColComboList(.ColIndex("RecevGl")) = "..."

End Select
End With
End Sub

Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

Private Sub ISButton1_Click()
    'Frame3.Visible = True
End Sub

Function fillExpensesFactoryGrid()
 
    '  «·’‰«⁄Ì…   ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    With Me.Fg_Journal
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(Me.XPTxtBillID.Text)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Fg_Journal
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                .TextMatrix(i, .ColIndex("LineNo")) = i
                
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("AccountName").value), "", RsExp.Fields("AccountName").value)
               
                .TextMatrix(i, .ColIndex("value")) = IIf(Not IsNumeric(RsExp.Fields("value").value), 0, RsExp.Fields("value").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsExp.Fields("des").value), "", RsExp.Fields("des").value)
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300

        If .Rows > 1 Then
            TXTFactoryExpenses = .Aggregate(flexSTSum, .FixedRows, .ColIndex("value"), .Rows, .ColIndex("value"))
        Else
            TXTFactoryExpenses = 0
        End If

    End With

    Grid.Visible = True
 
End Function

Function fillExpensesGrid()

    'Exit Function
    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    With Me.Grid
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

If TxtTransSerial.Text = "" Then
Exit Function
End If

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.Notes.NoteSerial1,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name,  dbo.ExpensesType.Namee ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID "
    'My_SQL = My_SQL & "  Where (dbo.Notes.NoteType = 3   and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ")  )  "
    'My_SQL = My_SQL + " WHERE     dbo.Notes.NoteType = 3 and    dbo.Notes.order_no='" & TxtTransSerial.text & "'"

    'My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID   Where ((dbo.Notes.NoteType = 3 ) and (buy is null))"



My_SQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.Buy, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.ItemID, dbo.Notes.Note_Value, dbo.ExpensesType.Name, "
My_SQL = My_SQL + " dbo.ExpensesType.namee , dbo.ExpensesType.Account_Code, dbo.notes_all.BasedONID"
My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
My_SQL = My_SQL + " dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
My_SQL = My_SQL + "  dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 3) AND (dbo.Notes.ORDER_NO = '" & TxtTransSerial.Text & "') AND (dbo.notes_all.BasedONID = 3)"
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Name").value), "", RsExp.Fields("Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Namee").value), "", RsExp.Fields("Namee").value)
                End If
               
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
            
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
           
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Note_Value").value), "", RsExp.Fields("Note_Value").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
            
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If

                .TextMatrix(i, .ColIndex("Select")) = 1
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid.Visible = True
    ' Expenses_update_total
 
End Function

Private Sub grid4_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    TXTFinacilaTotal.Text = fINANCIALiNVOICE_update_total
    cal_expenses
End Sub

Function fINANCIALiNVOICE_update_total() As Long
    Dim i As Integer
    On Error Resume Next

    If Grid4.Rows = 1 Then Exit Function
    fINANCIALiNVOICE_update_total = 0

    For i = 1 To Grid4.Rows - 1
        
        If Grid4.Cell(flexcpChecked, i, Grid4.ColIndex("select")) = flexChecked Then
            fINANCIALiNVOICE_update_total = fINANCIALiNVOICE_update_total + val(Grid4.TextMatrix(i, Grid4.ColIndex("note_value")))
        End If

    Next i
   
End Function

Private Sub GridIssueVoucer_Click()

    With GridIssueVoucer

        Select Case .Col

            Case 2

            Case 4
                FrmOutProductionOrder.Retrive val(.TextMatrix(.Row, 3))

            Case 5
                ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200

        End Select

    End With

End Sub

Private Sub Label10_Click()
    Frame3.Visible = False
End Sub

Private Sub ProkerId_Change()
If Me.TxtModFlg = "R" Then Exit Sub
   
   If val(TxtResProductionNo) <> 0 Then
        RetriveOrder TxtResProductionNo, 61, val(ProkerId.Text)
         
       
    End If
End Sub

Private Sub ReciveDate_Change()
    
          If Me.TxtModFlg = "E" Then
        If Month(rs("ReciveDate").value) = Month(ReciveDate.value) Then Exit Sub
    End If
  
    TxtresiveVoucher.Text = ""
    TxtNoteSerialV = ""
    TxtNoteSerial1V = ""
    
    




End Sub



Private Sub Txt_order_no_Change()

    If val(txt_ORDER_NO) <> 0 And val(CBoBasedON.ListIndex) >= 1 Then
    If val(CBoBasedON.ListIndex) = 1 Then
        RetriveOrder txt_ORDER_NO, , , 42
     ElseIf val(CBoBasedON.ListIndex) = 2 Then
        RetriveOrder txt_ORDER_NO, , , 6
    End If
        show_parts
        
        cal_expenses
        FIllEstimatedExpenses
    End If

End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)
If val(CBoBasedON.ListIndex) = 2 Then
    If KeyCode = vbKeyF3 Then
       Order_no_search.show
       If SystemOptions.UserInterface = ArabicInterface Then
         Order_no_search.Label1(2).Caption = "»ÕÀ «Ê«„— «·»Ì⁄"
        Else
         Order_no_search.Label1(2).Caption = "Sales Orders Search"
        End If
        Order_no_search.Caption = Order_no_search.Label1(2).Caption
         Order_no_search.RetrunType = 6
         Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
    End If
ElseIf val(CBoBasedON.ListIndex) = 1 Then
    If KeyCode = vbKeyF3 Then
    FrmBuySearch.DealingForm = GridTransType.salespricelist
    FrmBuySearch.Index = 222
    If SystemOptions.UserInterface = ArabicInterface Then
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄—÷ ”⁄—"
    Else
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄—÷ ”⁄—"
    End If
         FrmBuySearch.show vbModal
    End If
 End If
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtMIxCode_Change()
'If txtMIxCode.text = "" Then FG1.Rows = 1
 Me.txtMixID = ""
Me.txtMixID = GetMixIdFormCode(txtMIxCode)

add_item_to_parts_grid2
add_item_to_parts_grid1
End Sub

Private Sub txtMIxCode_KeyPress(KeyAscii As Integer)
'TxtMixID.text = ""
  'Fg1.Clear flexClearScrollable, flexClearEverything
  '  Fg1.Rows = 2
  '  Fg1.Clear flexClearScrollable, flexClearEverything
  '  Fg1.Refresh
End Sub

Private Sub TxtMIxCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchDevComItem.lbltype = 2
FrmSearchDevComItem.show

End If
End Sub

Private Sub txtMixID_Change()
add_item_to_parts_grid1
End Sub

Private Sub TxtResProductionNo_Change()

 

If Me.TxtModFlg = "R" Then Exit Sub
   
   If val(TxtResProductionNo) <> 0 Then
        RetriveOrder TxtResProductionNo, 61
         
       
    End If
End Sub

Private Sub TxtResProductionNo_KeyPress(KeyAscii As Integer)
txtMixID.Text = ""
  FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 2
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Refresh
End Sub

Private Sub TxtResProductionNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 61
       Order_no_search.lblSpecificsearch = 61
        'Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        Order_no_search.Caption = "»ÕÀ ”‰œ«  ÕÃ“ «·«‰ «Ã "
        Order_no_search.Label1(2).Caption = Order_no_search.Caption
    End If
    
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName2.BoundText = StoreId
    End If
End Sub

Private Sub TxtStoreID1_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID1)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub TxtTransSerial_Change()
    Retrive_orders_data (val(TxtTransSerial.Text))
End Sub

Sub RetriveSalesMixItems()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    FgMix.Clear flexClearScrollable, flexClearEverything
    FgMix.Rows = 1
sql = " SELECT     dbo.TblProductMixItems.ID, dbo.TblProductMixItems.TransectionID, dbo.TblProductMixItems.ItemID, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, "
sql = sql & "                      dbo.TblItems.ItemNamee, dbo.TblProductMixItems.MianItemID, TblItems_1.ItemName AS MainItemName, TblItems_1.ItemNamee AS MainItemNameE,"
sql = sql & "                      TblItems_1.Fullcode AS MainFullcode, dbo.TblProductMixItems.[Count], dbo.TblProductMixItems.QtyMix, dbo.TblProductMixItems.Qty, dbo.TblProductMixItems.Cost,"
sql = sql & "                      dbo.TblProductMixItems.Valu, dbo.TblProductMixItems.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblProductMixItems.UnitId,"
sql = sql & "                      dbo.TblUnites.UnitName , dbo.TblUnites.UnitNamee ,dbo.TblProductMixItems.MixCode"
sql = sql & " FROM         dbo.TblProductMixItems LEFT OUTER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.TblProductMixItems.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStore ON dbo.TblProductMixItems.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems TblItems_1 ON dbo.TblProductMixItems.MianItemID = TblItems_1.ItemID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TblProductMixItems.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TblProductMixItems.TransectionID = " & val(XPTxtBillID.Text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.FgMix
rs2.MoveFirst
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
 .TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("MixCode")) = IIf(IsNull(rs2("MixCode").value), "", rs2("MixCode").value)
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("MianItemID")) = IIf(IsNull(rs2("MianItemID").value), "", rs2("MianItemID").value)
.TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs2("StoreID").value), "", rs2("StoreID").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(rs2("UnitId").value), 0, rs2("UnitId").value)
.TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(rs2("Count").value), 0, rs2("Count").value)
.TextMatrix(i, .ColIndex("QtyMix")) = IIf(IsNull(rs2("QtyMix").value), 0, rs2("QtyMix").value)
.TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs2("Qty").value), 0, rs2("Qty").value)
.TextMatrix(i, .ColIndex("Cost")) = IIf(IsNull(rs2("Cost").value), 0, rs2("Cost").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(rs2("MainItemName").value), "", rs2("MainItemName").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
.TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
Else
.TextMatrix(i, .ColIndex("MainName")) = IIf(IsNull(rs2("MainItemNameE").value), "", rs2("MainItemNameE").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
.TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub SaveSalesMixItems(Optional TransID As Double)
Dim Rs3 As ADODB.Recordset
Dim i As Integer
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblProductMixItems where 1=-1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With FgMix
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("MianItemID"))) <> 0 Then
Rs3.AddNew

Rs3("TransectionID").value = TransID
Rs3("MixCode").value = (.TextMatrix(i, .ColIndex("MixCode")))
Rs3("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
Rs3("MianItemID").value = val(.TextMatrix(i, .ColIndex("MianItemID")))
Rs3("StoreID").value = val(.TextMatrix(i, .ColIndex("StoreID")))
Rs3("UnitId").value = val(.TextMatrix(i, .ColIndex("UnitId")))
Rs3("Count").value = val((.TextMatrix(i, .ColIndex("Count"))))
Rs3("Qty").value = val((.TextMatrix(i, .ColIndex("Qty"))))
Rs3("QtyMix").value = val((.TextMatrix(i, .ColIndex("QtyMix"))))
Rs3("Cost").value = val((.TextMatrix(i, .ColIndex("Cost"))))
Rs3("Valu").value = val((.TextMatrix(i, .ColIndex("Valu"))))
Rs3.update
End If
Next i
End With
End Sub
Sub FillMixItems()
    Dim RowNum As Integer
    FgMix.Clear flexClearScrollable, flexClearEverything
    FgMix.Rows = 1
    If SystemOptions.ProductionRawMaterMix = True Then
    For RowNum = 1 To FG.Rows - 1
        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) <> 0 Then
          AddRowFg RowNum
        End If

    Next RowNum
   End If
   cal_expenses
End Sub
Sub AddRowFg2(l As Integer)
Dim i As Integer
Dim k As Integer
Dim StoreName As String
Dim item_cost As Variant
Dim ItemName As String
Dim FllCode As String
If SystemOptions.UserInterface = ArabicInterface Then
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), FllCode, ItemName
Else
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), , StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), FllCode, ItemName
End If

With FgMix
k = .Rows
.Rows = .Rows + 1
        For i = k To .Rows - 1
            .TextMatrix(i, .ColIndex("MianItemID")) = val(FG.TextMatrix(l, FG.ColIndex("Code")))
            .TextMatrix(i, .ColIndex("MainName")) = ItemName
            .TextMatrix(i, .ColIndex("Count")) = val(FG.TextMatrix(l, FG.ColIndex("Count")))
            .TextMatrix(i, .ColIndex("StoreName")) = StoreName
            .TextMatrix(i, .ColIndex("StoreID")) = val(FG.TextMatrix(l, FG.ColIndex("StoreID2")))
             item_cost = ModItemCostPrice.GetCostItemPrice(val(FG.TextMatrix(l, FG.ColIndex("Code"))), 0, , , SystemOptions.SysMainStockCostMethod, , , , , val(FG.TextMatrix(l, FG.ColIndex("UnitId"))))
            .TextMatrix(i, .ColIndex("ItemID")) = val(FG.TextMatrix(l, FG.ColIndex("Code")))
            .TextMatrix(i, .ColIndex("Code")) = FllCode
            .TextMatrix(i, .ColIndex("Name")) = ItemName
            .TextMatrix(i, .ColIndex("UnitName")) = FG.TextMatrix(l, FG.ColIndex("UnitId"))
            .TextMatrix(i, .ColIndex("UnitId")) = GetItemUnitsId(.TextMatrix(i, .ColIndex("UnitName")))
            .TextMatrix(i, .ColIndex("QtyMix")) = 1
            If val(.TextMatrix(i, .ColIndex("QtyMix"))) <> 0 Then
            .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("Count"))) / val(.TextMatrix(i, .ColIndex("QtyMix")))
            End If
            .TextMatrix(i, .ColIndex("Cost")) = item_cost
            .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Cost"))) * val(.TextMatrix(i, .ColIndex("Qty")))
            
        Next i
 End With
    
End Sub
Sub AddRowFg3(l As Integer)
  Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
Dim k As Integer
Dim StoreName As String
Dim ItemName As String

If SystemOptions.UserInterface = ArabicInterface Then
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), , ItemName
Else
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), , StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), , ItemName
End If
  
               StrSQL = " SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + "      dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
                StrSQL = StrSQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
                StrSQL = StrSQL + "  FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + "      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
                StrSQL = StrSQL + "      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & val(FG.TextMatrix(l, FG.ColIndex("Code"))) & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
    Dim item_cost As Variant
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs2.RecordCount > 0 Then
With FgMix
If .Rows > 2 Then
k = .Rows - 1
Else
k = .Rows
End If
k = .Rows

.Rows = .Rows + rs2.RecordCount
        For i = k To .Rows - 1
If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(i, .ColIndex("StoreID")) = val(FG.TextMatrix(l, FG.ColIndex("StoreID2")))
getStorenames val(.TextMatrix(i, .ColIndex("StoreID"))), StoreName
Else
getStorenames val(.TextMatrix(i, .ColIndex("StoreID"))), , StoreName
End If
        .TextMatrix(i, .ColIndex("MixCode")) = (FG.TextMatrix(l, FG.ColIndex("MixNo")))
        .TextMatrix(i, .ColIndex("MianItemID")) = val(FG.TextMatrix(l, FG.ColIndex("Code")))
        .TextMatrix(i, .ColIndex("MainName")) = ItemName
        .TextMatrix(i, .ColIndex("Count")) = val(FG.TextMatrix(l, FG.ColIndex("Count")))
        .TextMatrix(i, .ColIndex("StoreName")) = StoreName
        
         item_cost = ModItemCostPrice.GetCostItemPrice(IIf(IsNull(rs2("PartItemID").value), 0, rs2("PartItemID").value), 0, , , SystemOptions.SysMainStockCostMethod, , , , , IIf(IsNull(rs2("Unitid").value), 0, rs2("Unitid").value))
            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("PartItemID").value), 0, rs2("PartItemID").value)
            .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
            Else
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
            End If
            .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(rs2("Unitid").value), 0, rs2("Unitid").value)
            .TextMatrix(i, .ColIndex("Qnty")) = IIf(IsNull(rs2("PartItemQty").value), 0, rs2("PartItemQty").value)
            .TextMatrix(i, .ColIndex("QtyMix")) = IIf(IsNull(rs2("PartItemQty").value), 0, rs2("PartItemQty").value)
            .TextMatrix(i, .ColIndex("QtyMix")) = IIf(IsNull(rs2("PartItemQty").value), 0, rs2("PartItemQty").value)
            If val(.TextMatrix(i, .ColIndex("QtyMix"))) <> 0 Then
            .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("QtyMix")))
            End If
            .TextMatrix(i, .ColIndex("Cost")) = item_cost
            .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Cost"))) * val(.TextMatrix(i, .ColIndex("Qty")))
            rs2.MoveNext
        Next i
 End With
 Else
 AddRowFg2 l
    End If
 End Sub
Sub AddRowFg(l As Integer)
  Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
Dim k As Integer
Dim StoreName As String
Dim ItemName As String

If SystemOptions.UserInterface = ArabicInterface Then
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), , ItemName
Else
getStorenames val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), , StoreName
GetItemData val(FG.TextMatrix(l, FG.ColIndex("Code"))), , ItemName
End If
  
    StrSQL = " SELECT     dbo.TblDefComItem.ID, dbo.TblDefComItem.MaxNo, dbo.TblDefComItemDet.ItemID, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee, "
    StrSQL = StrSQL + "                  dbo.TblDefComItemDet.UnitID , dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblDefComItemDet.Qty, dbo.TblDefComItemDet.cost ,dbo.TblDefComItem.Qty1,dbo.TblDefComItemDet.FlgX,dbo.TblDefComItem.StoreID2"
    StrSQL = StrSQL + " FROM         dbo.TblUnites RIGHT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblDefComItemDet ON dbo.TblUnites.UnitID = dbo.TblDefComItemDet.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblItems ON dbo.TblDefComItemDet.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblDefComItem ON dbo.TblDefComItemDet.IDDefCIT = dbo.TblDefComItem.ID"
    StrSQL = StrSQL + "         Where ( dbo.TblDefComItem.ItemNameID = " & val(FG.TextMatrix(l, FG.ColIndex("Code"))) & " and dbo.TblDefComItem.MaxNo='" & FG.TextMatrix(l, FG.ColIndex("MixNo")) & "') "
    Dim item_cost As Variant
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs2.RecordCount > 0 Then
With FgMix
If .Rows > 2 Then
k = .Rows - 1
Else
k = .Rows
End If
k = .Rows

.Rows = .Rows + rs2.RecordCount
        For i = k To .Rows - 1
If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs2("StoreID2").value), val(FG.TextMatrix(l, FG.ColIndex("StoreID2"))), rs2("StoreID2").value)
getStorenames val(.TextMatrix(i, .ColIndex("StoreID"))), StoreName
Else
getStorenames val(.TextMatrix(i, .ColIndex("StoreID"))), , StoreName
End If
        .TextMatrix(i, .ColIndex("MixCode")) = (FG.TextMatrix(l, FG.ColIndex("MixNo")))
        .TextMatrix(i, .ColIndex("MianItemID")) = val(FG.TextMatrix(l, FG.ColIndex("Code")))
        .TextMatrix(i, .ColIndex("MainName")) = ItemName
        .TextMatrix(i, .ColIndex("Count")) = val(FG.TextMatrix(l, FG.ColIndex("Count")))
        .TextMatrix(i, .ColIndex("StoreName")) = StoreName
        
         item_cost = ModItemCostPrice.GetCostItemPrice(IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value), 0, , , SystemOptions.SysMainStockCostMethod, , , , , IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value))
            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
            Else
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
            .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
            End If
            .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
            .TextMatrix(i, .ColIndex("Qnty")) = IIf(IsNull(rs2("Qty1").value), 0, rs2("Qty1").value)
            .TextMatrix(i, .ColIndex("QtyMix")) = IIf(IsNull(rs2("FlgX").value), 0, rs2("FlgX").value)
            .TextMatrix(i, .ColIndex("QtyMix")) = IIf(IsNull(rs2("Qty").value), 0, rs2("Qty").value)
            If val(.TextMatrix(i, .ColIndex("QtyMix"))) <> 0 Then
            .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("Count"))) / .TextMatrix(i, .ColIndex("Qnty")), 2) * val(.TextMatrix(i, .ColIndex("QtyMix")))
            End If
            .TextMatrix(i, .ColIndex("Cost")) = item_cost
            .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Cost"))) * val(.TextMatrix(i, .ColIndex("Qty")))
            rs2.MoveNext
        Next i
 End With
 Else
 AddRowFg3 l
    End If
 End Sub
Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    'On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim ShowTax As Boolean
    Dim Dcombos As ClsDataCombos

'    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    ScreenNameArabic = "«„— «·‘€· / «·«‰ «Ã "
    ScreenNameEnglish = "Production Order"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    ReciveDate.value = Date
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    ELe(4).Visible = ShowTax
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = GridTransType.ProductionOrder
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
        Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.StoreName = DCboStoreName
    
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = cmdAdd
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.txtPrice = txtPrice
    '//////////////////////////
    '/////////////////////////

    ' Resize_Form Me, TransactionSize
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStoreName2
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetLine Me.DcLine
    Dcombos.GetShift Me.DcShift
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmployees Me.DCEmp1
    Dcombos.get«hay Me.dcHey
Dcombos.GetEmployees Me.DCDriver, , True
Dcombos.get«hay Me.dcHey



    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    NewGrid.fillgrid
With FG
Command5.Visible = False
Command7.Visible = False
Command3.Visible = False
Command4.Visible = False
ReciveDate.Visible = False
Label27.Visible = False
TxtresiveVoucher.Visible = False
TxtIssueSerial.Visible = False
Label16.Visible = False
Label20.Visible = False
CmdIssueVoucher.Visible = False
CmdResiveVoucher.Visible = False

.ColHidden(.ColIndex("IssuTransID")) = True
.ColHidden(.ColIndex("ResiveNoteID")) = True
.ColHidden(.ColIndex("IssuNoteID")) = True
.ColHidden(.ColIndex("ReceivTransID")) = True

If SystemOptions.ProductionRawMaterMix = True Then
.ColHidden(.ColIndex("Voucher")) = False
.ColHidden(.ColIndex("RecevGl")) = False
.ColHidden(.ColIndex("ShowReceiv")) = False
.ColHidden(.ColIndex("ReceiveSerial")) = False
.ColHidden(.ColIndex("VoucheRecev")) = False
.ColHidden(.ColIndex("IssuGl")) = False
.ColHidden(.ColIndex("ShowIssue")) = False
.ColHidden(.ColIndex("IssueSerial")) = False
Else
Command5.Visible = True
Command7.Visible = True
Command3.Visible = True
Command4.Visible = True
ReciveDate.Visible = True
Label27.Visible = True
TxtresiveVoucher.Visible = True
TxtIssueSerial.Visible = True
Label16.Visible = True
Label20.Visible = True
CmdIssueVoucher.Visible = True
CmdResiveVoucher.Visible = True
.ColHidden(.ColIndex("Voucher")) = True
.ColHidden(.ColIndex("RecevGl")) = True
.ColHidden(.ColIndex("ShowReceiv")) = True
.ColHidden(.ColIndex("ReceiveSerial")) = True
.ColHidden(.ColIndex("VoucheRecev")) = True
.ColHidden(.ColIndex("IssuGl")) = True
.ColHidden(.ColIndex("ShowIssue")) = True
.ColHidden(.ColIndex("IssueSerial")) = True
End If
End With
If SystemOptions.UserInterface = ArabicInterface Then
       With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "⁄—÷ «”⁄«—"
        .AddItem "√„— »Ì⁄"
       ' .AddItem "›« Ê—… „‘ —Ì« "
       ' .AddItem "”‰œ  ÕÊÌ·"
       ' .AddItem "”‰œ ÕÃ“"
    End With
Else
      With Me.CBoBasedON
        .Clear
        .AddItem "NA"
        .AddItem "Quotation"
        .AddItem "Sales Order"
      '  .AddItem "Purchase Invoice"
      '  .AddItem "Transfer"
      '  .AddItem "Booking"
    End With
End If
    With Me.CboPriceType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "⁄—÷ √”⁄«— ·›Ê« Ì— «·»Ì⁄"
            .AddItem "«„— «·‘€· / «·«‰ «Ã"
        Else
            .AddItem "Sales  Order"
            .AddItem "Purchases   Order"
        End If

        .ListIndex = 0
    End With

    With CboPayMentType
        .Clear
        .AddItem "‰ﬁœ«"
        .AddItem "«Ã·"
    End With
    
  
CurrentTransactionType = 26

    StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=26 )" 'OR Transaction_Type=17
     StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
     'StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
     
    
    
    StrSQL = StrSQL + " Order By Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim My_SQL As String
    My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    fill_combo Me.DataCombo11, My_SQL

    My_SQL = " select code,account_name from markaas_taklefa"
 
    fill_combo Me.DataCombo1, My_SQL

    My_SQL = " select id,Project_name from projects"
 
    fill_combo Me.DataCombo2, My_SQL

    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL

    My_SQL = " select id,name from Shipment_mode"
 
    fill_combo Me.DcshipmentMethod, My_SQL

    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ   " & TxtTransSerial.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & "«·⁄„Ì·  " & DBCboClientName.Text & CHR(13) & " »‰«¡ ⁄·Ï ÿ·»Ì… —ﬁ„   " & txt_ORDER_NO & CHR(13) & "  „Œ“‰ «·„Ê«œ «·Œ«„  " & DCboStoreName2.Text & CHR(13) & " „Œ“‰  «·«‰ «Ã «· «„   " & DCboStoreName.Text & CHR(13) & " „·«ÕŸ«    " & txtRemark.Text & CHR(13) & "  «—ÌŒ  »œ«Ì… «·«‰ «Ã   " & startDate.value & " " & startTime.value & CHR(13) & "  «—ÌŒ  ‰Â«Ì… «·«‰ «Ã   " & EndDate.value & " " & EndTime.value
        LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr No.   " & TxtTransSerial.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Customer  " & DBCboClientName.Text & CHR(13) & " Basd On Sales Order No   " & txt_ORDER_NO & CHR(13) & "  R.M. Inventory " & DCboStoreName2.Text & CHR(13) & "F.G.  Inventory  " & DCboStoreName.Text & CHR(13) & " Remar;s   " & txtRemark.Text & CHR(13) & " Production Start at   " & startDate.value & " " & startTime.value & CHR(13) & " Production End at  " & EndDate.value & " " & EndTime.value
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, "", TxtTransSerial
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D", "", TxtTransSerial
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
FG.Enabled = True
    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«„— «·‘€· / «·«‰ «Ã"
            Else
                Me.Caption = "Production Order"
            End If

            Frame4.Enabled = False
            ELe(11).Enabled = False
            
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = True
        
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
           ' Fg.Editable = flexEDNone
           FG.Editable = flexEDKbdMouse
        
            CmdConvert.Enabled = True
            ' CmdConvert.Visible = True
            CmdTemplate.Visible = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdConvert.Enabled = False
            End If

            ELe(2).Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«„— «·‘€· / «·«‰ «Ã"
            Else
                Me.Caption = "Production Order"
            End If
   
            Frame4.Enabled = True
            ELe(11).Enabled = True
         
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            ' Me.XPBtnMove(0).Enabled = False
            ' Me.XPBtnMove(1).Enabled = False
            ' Me.XPBtnMove(2).Enabled = False
            ' Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
        
            '   CmdConvert.Visible = False
            CmdTemplate.Enabled = True
            CmdTemplate.Visible = True
            ELe(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«„— «·‘€· / «·«‰ «Ã"
            Else
                Me.Caption = "Production Order"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Frame4.Enabled = True
            ELe(11).Enabled = True
   
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
    
            FG.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
            XPBtnNewClients.Enabled = True
       
            ' CmdConvert.Visible = False
            CmdTemplate.Visible = False
            ELe(2).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
            
    Dim Num As Long

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    Me.dcHey.BoundText = IIf(IsNull(rs.Fields("Neighborhoodid").value), "", rs.Fields("Neighborhoodid").value)



TxtBatchNo.Text = IIf(IsNull(rs.Fields("BatchNo").value), "", rs.Fields("BatchNo").value)
DCEmp1.BoundText = IIf(IsNull(rs.Fields("empID1").value), "", rs.Fields("empID1").value)
    

    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If
  
    'Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)
   
    TxtIssueSerial.Text = IIf(IsNull(rs("Product_Issue_voucher_Serial").value), "", (rs("Product_Issue_voucher_Serial").value))
    TxtresiveVoucher.Text = IIf(IsNull(rs("Product_Receive_voucher_Serial").value), "", (rs("Product_Receive_voucher_Serial").value))
    Text1.Text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
      
      Txtnots2.Text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
      
       TxtStation.Text = IIf(IsNull(rs("Station").value), "", (rs("Station").value))
    CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), 0, (rs("CBoBasedON").value))
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
 txtMIxCode.Text = IIf(IsNull(rs("MIxCode").value), "", (rs("MIxCode").value))
    txtMixID.Text = IIf(IsNull(rs("MixID").value), "", (rs("MixID").value))
    
ProkerId.Text = IIf(IsNull(rs("ProkerId").value), "", (rs("ProkerId").value))
    TxtResProductionNo.Text = IIf(IsNull(rs("ResProductionNo").value), "", (rs("ResProductionNo").value))
    
    
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
Me.dcHey.BoundText = IIf(IsNull(rs.Fields("Neighborhoodid").value), "", rs.Fields("Neighborhoodid").value)


    Me.DcshipmentMethod.BoundText = IIf(IsNull(rs("shipmentMethod").value), "", rs("shipmentMethod").value)
    txtShipmentPrice.Text = IIf(Not IsNumeric(rs("ShipmentPrice").value), 0, (rs("ShipmentPrice").value))
    TxtWorkHour.Text = IIf(Not IsNumeric(rs("WorkHour").value), 0, (rs("WorkHour").value))

    startDate.value = IIf(IsNull(rs("startDate").value), Date, (rs("startDate").value))
    EndDate.value = IIf(IsNull(rs("EndDate").value), Date, (rs("EndDate").value))
    Dim timevalue As Data

    If Not IsNull(rs("startTime").value) Then
        'timevalue = rs("startTime").value
        '  Me.startTime.value = rs("startTime").value 'timevalue
   
    End If

    If Not IsNull(rs("EndTime").value) Then
        ' timevalue = rs("EndTime").value
        '   Me.EndTime.value = rs("EndTime").value ' timevalue
        '
    End If
 ReciveDate.value = IIf(IsNull(rs("ReciveDate").value), rs("Transaction_Date").value, (rs("ReciveDate").value))
    
       TxtManualNo1.Text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    TxtProductionPlanno.Text = IIf(IsNull(rs("ProductionPlanno").value), "", (rs("ProductionPlanno").value))
 
    
    TxtShipmentArae.Text = IIf(IsNull(rs("ShipmentArae").value), "", (rs("ShipmentArae").value))
    txtRemark.Text = IIf(IsNull(rs("Remark").value), "", (rs("Remark").value))
    'Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    'If rs("Transaction_Type").value = 6 Then
    '    Me.CboPriceType.ListIndex = 1
    'ElseIf rs("Transaction_Type").value = 17 Then '17
    '    Me.CboPriceType.ListIndex = 0
    'End If
    TXTLineExpenses.Text = IIf(IsNull(rs("LineExpenses").value), 0, rs("LineExpenses").value)
    TxtHourdippTotal.Text = IIf(IsNull(rs("HourdippTotal").value), 0, rs("HourdippTotal").value)
    TxtUsedPowerPriceHTotal.Text = IIf(IsNull(rs("UsedPowerPriceHTotal").value), 0, rs("UsedPowerPriceHTotal").value)
    TxtUsedElectricPriceHTotal.Text = IIf(IsNull(rs("UsedElectricPriceHTotal").value), 0, rs("UsedElectricPriceHTotal").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DCboStoreName2.BoundText = IIf(IsNull(rs("StoreID1").value), "", rs("StoreID1").value)
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    TxtworkerTotalPerHour.Text = IIf(IsNull(rs("WorkerTotalPerHour").value), 0, rs("WorkerTotalPerHour").value)
    txt_ORDER_NO.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)

    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)

    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
    LblTotalQty.Caption = 0

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
             FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
             FG.TextMatrix(Num, FG.ColIndex("ResiveNoteID")) = IIf(IsNull(RsDetails("ResiveNoteID")), "", (RsDetails("ResiveNoteID").value))
             FG.TextMatrix(Num, FG.ColIndex("IssuNoteID")) = IIf(IsNull(RsDetails("IssuNoteID")), "", (RsDetails("IssuNoteID").value))
             FG.TextMatrix(Num, FG.ColIndex("IssuTransID")) = IIf(IsNull(RsDetails("IssuTransID")), "", (RsDetails("IssuTransID").value))
             FG.TextMatrix(Num, FG.ColIndex("ReceivTransID")) = IIf(IsNull(RsDetails("ReceivTransID")), "", (RsDetails("ReceivTransID").value))
             FG.TextMatrix(Num, FG.ColIndex("IssueSerial")) = IIf(IsNull(RsDetails("IssueSerial")), "", (RsDetails("IssueSerial").value))
             FG.TextMatrix(Num, FG.ColIndex("ReceiveSerial")) = IIf(IsNull(RsDetails("ReceiveSerial")), "", (RsDetails("ReceiveSerial").value))
             FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
             FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
             FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
             FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
             FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
             FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
             FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
             FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
             FG.TextMatrix(Num, FG.ColIndex("MixNo")) = IIf(IsNull(RsDetails("MixNo")), "", (RsDetails("MixNo").value))
             FG.TextMatrix(Num, FG.ColIndex("L")) = IIf(IsNull(RsDetails("L")), "", (RsDetails("L").value))
             FG.TextMatrix(Num, FG.ColIndex("W")) = IIf(IsNull(RsDetails("W")), "", (RsDetails("W").value))
             FG.TextMatrix(Num, FG.ColIndex("H1")) = IIf(IsNull(RsDetails("H1")), "", (RsDetails("H1").value))
             FG.TextMatrix(Num, FG.ColIndex("H2")) = IIf(IsNull(RsDetails("H2")), "", (RsDetails("H2").value))
             FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
             FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
             FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
             FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
             ''/////////
             FG.TextMatrix(Num, FG.ColIndex("NoHours")) = IIf(IsNull(RsDetails("NoHours")), 0, (RsDetails("NoHours").value))
             FG.TextMatrix(Num, FG.ColIndex("PriceNoHours")) = IIf(IsNull(RsDetails("PriceNoHours")), 0, (RsDetails("PriceNoHours").value))
             FG.TextMatrix(Num, FG.ColIndex("TotalPriceNoHours")) = IIf(IsNull(RsDetails("TotalPriceNoHours")), 0, (RsDetails("TotalPriceNoHours").value))
             
            FG.TextMatrix(Num, FG.ColIndex("DistibutePercentage")) = IIf(IsNull(RsDetails("DistibutePercentage")), "", (RsDetails("DistibutePercentage").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
           FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))

            FG.TextMatrix(Num, FG.ColIndex("CorrectionID")) = IIf(IsNull(RsDetails("CorrectionID")), 1, (RsDetails("CorrectionID").value))

            FG.TextMatrix(Num, FG.ColIndex("StoreID2")) = IIf(IsNull(RsDetails("StoreID2")), 1, (RsDetails("StoreID2").value))

            LngCurItemID = val(FG.TextMatrix(Num, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Dim RsUnitData As New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                 
                LblTotalQty = LblTotalQty + val(FG.TextMatrix(Num, FG.ColIndex("Count"))) * val(RsUnitData("UnitFactor").value)
                RsUnitData.Close
            End If
            
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    '«” œ⁄«¡ «·ŒÿÊÿ «·⁄«„·… ›Ì «·Œÿ
    Dim WorkLines As New ADODB.Recordset
    Dim LngRow As Long
    StrSQL = "Select * from TblProductOrderLines where Transaction_ID=" & val(XPTxtBillID.Text)
    WorkLines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FGLine.Clear flexClearScrollable, flexClearEverything
         
    Dim RowNum As Integer
          
    If WorkLines.RecordCount > 0 Then
        FGLine.Rows = 2
        Me.FGLine.Rows = Me.FGLine.Rows + WorkLines.RecordCount - 1

        For RowNum = 1 To WorkLines.RecordCount
       
            LngRow = RowNum
           
            With Me.FGLine
                .TextMatrix(LngRow, .ColIndex("Ser")) = RowNum
                .TextMatrix(LngRow, .ColIndex("id")) = IIf(IsNull(WorkLines("lineid").value), "", WorkLines("lineid"))
                .TextMatrix(LngRow, .ColIndex("name")) = IIf(IsNull(WorkLines("name").value), "", WorkLines("name"))
                .TextMatrix(LngRow, .ColIndex("code")) = IIf(IsNull(WorkLines("code").value), "", WorkLines("code").value)
                .TextMatrix(LngRow, .ColIndex("UsedPowerPriceH")) = IIf(Not IsNumeric(WorkLines("UsedPowerPriceH").value), 0, WorkLines("UsedPowerPriceH").value)
                .TextMatrix(LngRow, .ColIndex("UsedElectricPriceH")) = IIf(Not IsNumeric(WorkLines("UsedElectricPriceH").value), 0, WorkLines("UsedElectricPriceH").value)
                '.TextMatrix(LngRow, .ColIndex("WorkerPriceH")) = IIf(Not IsNumeric(WorkLines("WorkerPriceH").value), 0, WorkLines("WorkerPriceH").value)
                .TextMatrix(LngRow, .ColIndex("Hourdipp")) = IIf(IsNull(WorkLines("Hourdipp").value), 0, WorkLines("Hourdipp").value)
                .TextMatrix(LngRow, .ColIndex("from")) = IIf(IsNull(WorkLines("fromt").value), "", WorkLines("fromt").value)
                .TextMatrix(LngRow, .ColIndex("to")) = IIf(IsNull(WorkLines("tot").value), "", WorkLines("tot").value)
                .TextMatrix(LngRow, .ColIndex("shift")) = IIf(IsNull(WorkLines("shift").value), "", WorkLines("shift").value)
                .TextMatrix(LngRow, .ColIndex("hour")) = IIf(Not IsNumeric(WorkLines("hour").value), 0, WorkLines("hour").value)
                .TextMatrix(LngRow, .ColIndex("shiftname")) = IIf(IsNull(WorkLines("shiftname").value), "", WorkLines("shiftname").value)
                .TextMatrix(LngRow, .ColIndex("total")) = (val(.TextMatrix(LngRow, .ColIndex("UsedPowerPriceH"))) + val(.TextMatrix(LngRow, .ColIndex("UsedElectricPriceH")))) * .TextMatrix(LngRow, .ColIndex("hour"))
            End With

            WorkLines.MoveNext
        Next RowNum

        '       Me.FGLine.Rows = Me.FGLine.Rows + 1
        CalculateNets
    End If
         
    '«” œ⁄«¡   «·⁄„«·… ›Ì «·Œÿ
    Dim WorkWorker As New ADODB.Recordset
     
    StrSQL = "Select * from TblProductOrderWorker where Transaction_ID=" & val(XPTxtBillID.Text)
    WorkWorker.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    GridWorker.Clear flexClearScrollable, flexClearEverything
    GridWorker.Rows = 2
          
    If WorkWorker.RecordCount > 0 Then
        Me.GridWorker.Rows = Me.GridWorker.Rows + WorkWorker.RecordCount - 1

        For RowNum = 1 To WorkWorker.RecordCount
       
            LngRow = RowNum
           
            With Me.GridWorker
                .TextMatrix(LngRow, .ColIndex("LineNo")) = RowNum
                .TextMatrix(LngRow, .ColIndex("Emp_id")) = IIf(IsNull(WorkWorker("Emp_id").value), "", WorkWorker("Emp_id"))
                .TextMatrix(LngRow, .ColIndex("name")) = IIf(IsNull(WorkWorker("name").value), "", WorkWorker("name"))
                .TextMatrix(LngRow, .ColIndex("code")) = IIf(IsNull(WorkWorker("code").value), "", WorkWorker("code").value)
                .TextMatrix(LngRow, .ColIndex("hourprice")) = IIf(Not IsNumeric(WorkWorker("hourprice").value), 0, WorkWorker("hourprice").value)
                .TextMatrix(LngRow, .ColIndex("from")) = IIf(IsNull(WorkWorker("fromt").value), "", WorkWorker("fromt").value)
                .TextMatrix(LngRow, .ColIndex("to")) = IIf(IsNull(WorkWorker("tot").value), "", WorkWorker("tot").value)
                .TextMatrix(LngRow, .ColIndex("shift")) = IIf(Not IsNumeric(WorkWorker("shift").value), 0, WorkWorker("shift").value)
                .TextMatrix(LngRow, .ColIndex("hour")) = IIf(Not IsNumeric(WorkWorker("hour").value), 0, WorkWorker("hour").value)
                .TextMatrix(LngRow, .ColIndex("total")) = (val(.TextMatrix(LngRow, .ColIndex("hour")))) * .TextMatrix(LngRow, .ColIndex("hourprice"))
            End With

            WorkWorker.MoveNext
        Next RowNum

        '       Me.GridWorker.Rows = Me.GridWorker.Rows + 1
        'CalculateNets
        With GridWorker
            TxtworkerTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
            TxtworkerTotalPerHour.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("hourprice"), .Rows - 1, .ColIndex("hourprice"))
 
        End With

    End If
         
    ' ⁄»… «–Ê‰«  «·’—›
    fillExpensesGrid
    ' ⁄»…   «·›Ê« Ì— «·„«·Ì…
    fillFinancialInvoiceGrid

    TXTFinacilaTotal.Text = fINANCIALiNVOICE_update_total
    Me.Txt_EXport.Text = Expenses_update_total

    ' ⁄»∆… «–‰Ê‰«  «·’—› «·’‰«⁄Ì…
    fillExpensesFactoryGrid
 
    show_parts
    '⁄— ÷ «· ﬂ«·Ì› «·’‰«⁄ÌÂ «· ﬁœÌ—Ì…
    RetriveSalesMixItems
    FIllEstimatedExpenses
    add_item_to_parts_grid1
    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    'Õ”«» «·„’—Ê›«  Ê «· ﬂ·›… «·‰Â«∆Ì…
    cal_expenses
       
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub RetriveOrder(Optional order_no As String = "", Optional Transaction_Type As Integer = 0, Optional Transaction_ID As Double = 0, Optional Trans As Integer)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
        FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
  If Transaction_Type = 0 Then
    StrSQL = "Select * from transactions  where  Transaction_Type=" & Trans & " and   order_no='" & order_no & "'"
Else
    StrSQL = "Select * from transactions  where  Transaction_Type=" & Transaction_Type & " and   NoteSerial1='" & order_no & "'"

End If
If Transaction_ID <> 0 Then
StrSQL = "Select * from transactions  where  Transaction_ID=" & Transaction_ID
End If

    
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
         DCDriver.BoundText = "" ' IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
        dcHey.BoundText = "" 'IIf(IsNull(rs("Neighborhoodid").value), "", rs("Neighborhoodid").value)
        
txtMixID.Text = "" ' IIf(IsNull(rs("MixID").value), "", rs("MixID").value)
        txtMIxCode.Text = "" ' IIf(IsNull(rs("MIxCode").value), "", rs("MIxCode").value)
                txtRemark.Text = "" ' IIf(IsNull(rs("TransactionComment").value), "", rs("TransactionComment").value)
 Me.dcHey.BoundText = ""

    If rs.RecordCount < 1 Then
 
 

        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
       
txtMixID.Text = IIf(IsNull(rs("MixID").value), "", rs("MixID").value)

        txtMIxCode.Text = IIf(IsNull(rs("MIxCode").value), "", rs("MIxCode").value)
                txtRemark.Text = IIf(IsNull(rs("TransactionComment").value), "", rs("TransactionComment").value)
                
     Me.dcHey.BoundText = IIf(IsNull(rs.Fields("Neighborhoodid").value), "", rs.Fields("Neighborhoodid").value)
           
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = 0 ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("L")) = IIf(IsNull(RsDetails("L")), "", (RsDetails("L").value))
             FG.TextMatrix(Num, FG.ColIndex("W")) = IIf(IsNull(RsDetails("W")), "", (RsDetails("W").value))
             FG.TextMatrix(Num, FG.ColIndex("H1")) = IIf(IsNull(RsDetails("H1")), "", (RsDetails("H1").value))
             FG.TextMatrix(Num, FG.ColIndex("H2")) = IIf(IsNull(RsDetails("H2")), "", (RsDetails("H2").value))
             FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
             FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
             FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
             FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
If txtMixID <> "" Then
add_item_to_parts_grid1
End If
    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If XPTxtBillID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì…  —ﬁ„ " & CHR(13)
            Msg = Msg + (TxtTransSerial.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø" & CHR(13)
             Msg = Msg + " ”Ì‰ Ã ⁄‰ Â–… «·⁄„·Ì… Õ–› ﬂ· ”‰œ«  «·«‰ «Ã «· «„ «·Œ«’… »Â«" & CHR(13)
        Else
            Msg = " Delete Order NO  " & CHR(13)
            Msg = Msg + (TxtTransSerial.Text) & CHR(13)
            Msg = Msg + " Confrim Delete?" & CHR(13)
            Msg = Msg + " it Will Delete All Production Recive Voucher" & CHR(13)
            
    
        End If

        Dim StrSQL As String
Dim i As Integer
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
    
                StrSQL = "update Notes set  Transaction_ID1=Null , ItemID=NUll, buy = null Where   (Transaction_ID1=" & val(Me.XPTxtBillID.Text) & ")"
                Cn.Execute StrSQL
            
                StrSQL = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=Null ,  ItemID=NUll, buy = null Where  ( Transaction_ID1=" & val(Me.XPTxtBillID.Text) & ")"
                Cn.Execute StrSQL
            With FG
            For i = 1 To .Rows - 1
            DeleteTransactiomsVoucher val(.TextMatrix(i, .ColIndex("IssuTransID")))
            DeleteTransactiomsVoucher val(.TextMatrix(i, .ColIndex("ReceivTransID")))
            Next i
            End With
            Cn.Execute " delete from   TblProductMixItems where  TransectionID=" & val(XPTxtBillID.Text)
                DeleteTransactiomsVoucher val(Text1.Text)
                DeleteTransactiomsVoucher val(Txtnots2.Text)
                
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "«„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «„— «·«‰ «Ã   ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "«„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «„— «·«‰ «Ã «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " «„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «„— «·«‰ «Ã   «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " «„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰    «„— «·«‰ «Ã" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«„— «·«‰ «Ã  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "  «„— «·«‰ «Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    'Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
    'On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If DBCboClientName.Text = "" Then
  ''          If SystemOptions.UserInterface = ArabicInterface Then
   '             Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
   '         Else
   '             Msg = "Please Select Vendor"
   '         End If
'
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DBCboClientName.SetFocus
'            SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
        End If
    
        If DCboStoreName2.Text = "" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = "ÌÃ»  ÕœÌœ „Œ“‰ «·„Ê«œ «·Œ«„"
'            Else
'                Msg = "Select Inventory  For ROM"
'            End If
'
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            '        DCboStoreName2.SetFocus
'            SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
        End If
    
        If DCboStoreName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ „Œ“‰ «·«‰ «Ã «· «„"
            Else
                Msg = "Select Inventory For Finished GoodS"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    
    
        If DCboStoreName2.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ „Œ“‰ «·„Ê«œ «·Œ«„"
            Else
                Msg = "Select Inventory For RM"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        DCboStoreName.SetFocus
         SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            Else
                Msg = "Specify Payment Method"
            End If
        End If
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

 Dim Sanad_No As Integer
Sanad_No = 49

        my_branch = val(Dcbranch.BoundText)

        If TxtTransSerial.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , CurrentTransactionType, , val(DCboStoreName.BoundText)) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , CurrentTransactionType, , val(DCboStoreName.BoundText)) = "" Then
                    TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=26"))
                Else
                    TxtTransSerial.Text = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , CurrentTransactionType, , val(DCboStoreName.BoundText))
                End If
            End If
        End If
 
 
 
        Set RSTransDetails = New ADODB.Recordset
       ' RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
        Cn.BeginTrans
        BeginTrans = True

        If Me.TxtModFlg.Text = "N" Then
            rs.AddNew
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            
            
            
                
        End If

        Screen.MousePointer = vbArrowHourglass
        rs("Transaction_ID").value = val(XPTxtBillID.Text)
        rs.Fields("empID1").value = IIf(DCEmp1.BoundText <> "", val(DCEmp1.BoundText), Null)

        rs("Transaction_Serial").value = (TxtTransSerial.Text)
        rs("NoteSerial1").value = (TxtTransSerial.Text)
        
            rs("ManualNo1").value = IIf(TxtManualNo1.Text = "", Null, val(TxtManualNo1.Text))
    rs("ProductionPlanno").value = IIf(TxtProductionPlanno.Text = "", Null, val(TxtProductionPlanno.Text))
  
      rs("Station").value = IIf(TxtStation.Text = "", Null, (TxtStation.Text))

        rs("order_no").value = txt_ORDER_NO.Text
    
        If chkshipped.value = vbChecked Then
            rs("shipped").value = 1
        Else
            rs("shipped").value = 0
        End If

        rs("Transaction_Date").value = XPDtbBill.value
     
        rs("startDate").value = startDate.value
        rs("Transaction_Date").value = XPDtbBill.value
        rs("EndDate").value = EndDate.value
        rs("startTime").value = FormatDateTime(Me.startTime.value, vbLongTime)
        rs("EndTime").value = FormatDateTime(Me.EndTime.value, vbLongTime)
        rs("MixID").value = val(txtMixID.Text)
        rs("MIxCode").value = txtMIxCode.Text
        rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
        rs.Fields("Neighborhoodid").value = IIf(dcHey.BoundText <> "", val(dcHey.BoundText), Null)
        FillMixItems
        rs("BranchId").value = val(Me.Dcbranch.BoundText)
 
        rs("Transaction_Type").value = 26

        If CboPayMentType.ListIndex = -1 Then
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPayMentType.ListIndex)
        End If
        rs("LineExpenses").value = val(TXTLineExpenses.Text)
        rs("HourdippTotal").value = val(TxtHourdippTotal.Text)
        rs("UsedPowerPriceHTotal").value = val(TxtUsedPowerPriceHTotal.Text)
        rs("UsedElectricPriceHTotal").value = val(TxtUsedElectricPriceHTotal.Text)
        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "" Or DBCboClientName.Text = "", Null, val(DBCboClientName.BoundText))
        rs("shipmentMethod").value = IIf(DcshipmentMethod.BoundText = "", Null, val(DcshipmentMethod.BoundText))
        rs("ShipmentPrice").value = IIf(txtShipmentPrice.Text = "", 0, val(txtShipmentPrice.Text))
        rs("ShipmentArae").value = IIf(TxtShipmentArae.Text = "", Null, TxtShipmentArae.Text)
        rs("Product_Issue_voucher_Serial").value = IIf(TxtIssueSerial.Text = "", Null, TxtIssueSerial.Text)
        rs("Product_Receive_voucher_Serial").value = IIf(TxtresiveVoucher.Text = "", Null, TxtresiveVoucher.Text)
        rs.Fields("Neighborhoodid").value = IIf(dcHey.BoundText <> "", val(dcHey.BoundText), Null)
        rs("Remark").value = IIf(txtRemark.Text = "", Null, txtRemark.Text)
        rs("ProkerId").value = IIf(ProkerId.Text = "", Null, ProkerId.Text)
        rs("ResProductionNo").value = IIf(TxtResProductionNo.Text = "", Null, TxtResProductionNo.Text)
     If CBoBasedON.ListIndex = -1 Then
        rs("CBoBasedON").value = 0
    Else
        rs("CBoBasedON").value = val(CBoBasedON.ListIndex)
    End If
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("StoreID1").value = IIf(DCboStoreName2.BoundText = "", Null, val(DCboStoreName2.BoundText))
     
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
        rs("total").value = IIf(XPTxtSum.Text = "", Null, val(XPTxtSum.Text))
        rs("WorkHour").value = IIf(TxtWorkHour.Text = "", Null, val(TxtWorkHour.Text))
   
        rs("LineExpenses").value = IIf(Not IsNumeric(TXTLineExpenses.Text), 0, val(TXTLineExpenses.Text))
        rs("workerTotal").value = IIf(Not IsNumeric(TxtworkerTotal.Text), 0, val(TxtworkerTotal.Text))
        rs("Expenses").value = IIf(Not IsNumeric(Txt_EXport.Text), 0, val(Txt_EXport.Text))
        rs("FinacilaTotal").value = IIf(Not IsNumeric(TXTFinacilaTotal.Text), 0, val(TXTFinacilaTotal.Text))
        rs("FactoryExpenses").value = IIf(Not IsNumeric(TXTFactoryExpenses.Text), 0, val(TXTFactoryExpenses.Text))
        rs("TotalMaterials").value = IIf(Not IsNumeric(TxtTotalMaterials.Text), 0, val(TxtTotalMaterials.Text))
        rs("WorkerTotalPerHour").value = val(TxtworkerTotalPerHour.Text)
   
        rs("IndirectCostForProduction").value = IIf(Not IsNumeric(TxtIndirectCostForProduction.Text), 0, val(TxtIndirectCostForProduction.Text))
        rs("TotalEstimatedCost").value = IIf(Not IsNumeric(TxtTotalEstimatedCost.Text), 0, val(TxtTotalEstimatedCost.Text))
    rs("ReciveDate").value = ReciveDate.value
        rs.update
        CuurentLogdata

        If Me.TxtModFlg.Text = "E" Then
        Dim i As Integer
            With FG
            For i = 1 To .Rows - 1
            DeleteTransactiomsVoucher val(.TextMatrix(i, .ColIndex("IssuTransID")))
            DeleteTransactiomsVoucher val(.TextMatrix(i, .ColIndex("ReceivTransID")))
            Next i
            End With
            Cn.Execute " delete from   TblProductMixItems where  TransectionID=" & val(XPTxtBillID.Text)
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        Retrive_orders_data (val(TxtTransSerial.Text))
        cal_expenses

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("order_id").value = val(XPTxtBillID.Text)
                RSTransDetails("ColorID").value = 1
                'RSTransDetails("order_no").value = Txt_order_no.text
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("Lineexpenses").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Expenses")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Expenses"))))
            '''///////////
                RSTransDetails("NoHours").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoHours")) = ""), 0, val((FG.TextMatrix(RowNum, FG.ColIndex("NoHours")))))
                RSTransDetails("PriceNoHours").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PriceNoHours")) = ""), 0, val((FG.TextMatrix(RowNum, FG.ColIndex("PriceNoHours")))))
                RSTransDetails("TotalPriceNoHours").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TotalPriceNoHours")) = ""), 0, val((FG.TextMatrix(RowNum, FG.ColIndex("TotalPriceNoHours")))))
                RSTransDetails("L").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("L")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("L"))))
                RSTransDetails("W").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("W")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("W"))))
                RSTransDetails("H1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("H1")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("H1"))))
                RSTransDetails("H2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("H2")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("H2"))))
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NoCount"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Width"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Area"))))
                RSTransDetails("MixNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MixNo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("MixNo"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("DistibutePercentage").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DistibutePercentage")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DistibutePercentage"))))
            
                RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("StoreID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("StoreID2")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("StoreID2"))))
            
                RSTransDetails("CorrectionID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CorrectionID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("CorrectionID"))))
                        
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    '                RSTransDetails("ShowPrice").value = Val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) * RSTransDetails("Quantity").value
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If

                RSTransDetails("ShowPrice").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))))
             Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost
'      RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
'       RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       

                RSTransDetails.update
            End If

        Next RowNum
    
        'Õ›Ÿ «·ŒÿÊÿ «·⁄«„·… ›Ì «·Œÿ
        Dim WorkLines As New ADODB.Recordset

        If Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete TblProductOrderLines where Transaction_ID=" & val(XPTxtBillID.Text)
            Cn.Execute StrSQL
        End If

        StrSQL = "Select * from TblProductOrderLines where Transaction_ID=" & val(XPTxtBillID.Text)
        WorkLines.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        For RowNum = 1 To FGLine.Rows - 1

            If FGLine.TextMatrix(RowNum, FGLine.ColIndex("id")) <> "" Then
                WorkLines.AddNew
                WorkLines("Transaction_ID").value = val(XPTxtBillID.Text)
                WorkLines("LineID").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("id"))
                WorkLines("code").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("code"))
                WorkLines("name").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("name"))
                WorkLines("Hourdipp").value = val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("Hourdipp")))
                WorkLines("UsedPowerPriceH").value = val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("UsedPowerPriceH")))
                WorkLines("UsedElectricPriceH").value = val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("UsedElectricPriceH")))
                WorkLines("fromt").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("from"))
                WorkLines("tot").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("to"))
                WorkLines("Hour").value = val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("Hour")))
                WorkLines("shiftname").value = FGLine.TextMatrix(RowNum, FGLine.ColIndex("shiftname"))
                WorkLines("Shift").value = val(FGLine.TextMatrix(RowNum, FGLine.ColIndex("Shift")))
                WorkLines.update
            End If
         
        Next RowNum
 
        'Õ›Ÿ «·⁄„«·…   ›Ì «·Œÿ
        Dim WorkWorker As New ADODB.Recordset

        If Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete TblProductOrderWorker where Transaction_ID=" & val(XPTxtBillID.Text)
            Cn.Execute StrSQL
        End If

        StrSQL = "Select * from TblProductOrderWorker where Transaction_ID=" & val(XPTxtBillID.Text)
        WorkWorker.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        For RowNum = 1 To GridWorker.Rows - 1

            If GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("Emp_id")) <> "" Then
                WorkWorker.AddNew
                WorkWorker("Transaction_ID").value = val(XPTxtBillID.Text)
                WorkWorker("Emp_id").value = val(GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("Emp_id")))
                WorkWorker("code").value = GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("code"))
                WorkWorker("name").value = GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("name"))
                WorkWorker("hourprice").value = val(GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("hourprice")))
                WorkWorker("fromt").value = GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("from"))
                WorkWorker("tot").value = GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("to"))
                WorkWorker("Hour").value = val(GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("Hour")))
                WorkWorker("Shift").value = GridWorker.TextMatrix(RowNum, GridWorker.ColIndex("Shift"))
                WorkWorker.update
            End If
         
        Next RowNum

        'Õ›Ÿ «·„’—Ê›«  «·’‰«⁄Ì…   ›Ì «·Œÿ
        Dim FactoryExpenses As New ADODB.Recordset

        If Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.Text)
            Cn.Execute StrSQL
        End If

        StrSQL = "Select * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.Text)
        FactoryExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        For RowNum = 1 To Fg_Journal.Rows - 2

            If Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" Then
                FactoryExpenses.AddNew
                FactoryExpenses("Transaction_ID").value = val(XPTxtBillID.Text)
        
                FactoryExpenses("AccountName").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName"))
                FactoryExpenses("value").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("value")))
                FactoryExpenses("des").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("des"))
                FactoryExpenses.update
            End If
         
        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        save_expenses
        Save_Financial_invoice
       SaveSalesMixItems val(XPTxtBillID.Text)
       
        '   CmdIssueVoucher_Click
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… Ê«‰‘«¡ «–‰ ’—› «·Ì" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
                    Msg = Msg + "do you new Operation?"
        
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Save_Financial_invoice()
 
    Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String
  
    With Grid4
 
        For i = 1 To .Rows - 1
      
            Cn.BeginTrans
 
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
        
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=" & val(Me.XPTxtBillID.Text) & " , buy='1'  where Double_Entry_Vouchers_ID=" & val(.TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")))
        
            Else
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=null , buy=Null where Double_Entry_Vouchers_ID=" & val(.TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    '    DoEvents
    '    Command4_Click
End Sub

Private Sub save_expenses()
    Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String
 
    With Grid

        For i = 1 To Grid.Rows - 1
      
            Cn.BeginTrans
 
            If Grid.Cell(flexcpChecked, i, Grid.ColIndex("select")) = flexChecked Then
         
                sql = "update notes set Transaction_ID1=" & val(Me.XPTxtBillID.Text) & " , buy='1' " & " where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        
            Else
                sql = "update notes set Transaction_ID1=null ,  buy=Null  where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    ' Expenses_update_total

End Sub

Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    .DealingForm = ShowPrice
    '    .show vbModal
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    'End With

End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        lbl(4).Enabled = True
    Else
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport2()
    On Error GoTo ErrTrap
    Dim BuyReport As ClsBuyReport

    If Not XPTxtBillID.Text Then
        Set BuyReport = New ClsBuyReport
        BuyReport.ShowBuyData XPTxtBillID.Text, 6, True
    End If

    Exit Sub
ErrTrap:

 

End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap
    Dim BuyReport As ClsBuyReport

    If Not XPTxtBillID.Text Then
        Set BuyReport = New ClsBuyReport
        BuyReport.ShowBuyData XPTxtBillID.Text, 2, True
    End If

    Exit Sub
ErrTrap:

    'On Error GoTo ErrTrap
    'If XPTxtBillID.text <> "" Then
    '    Set SaleReport = New ClsSaleReport
    '    SaleReport.ShowPrice XPTxtBillID.text
    'End If
    'Exit Sub
    'ErrTrap:

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Me.TxtModFlg.Text = "E" Then
        TxtresiveVoucher.Text = ""
        TxtIssueSerial.Text = ""
        TxtTransSerial.Text = ""
        
    End If

End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    chkshipped.Caption = "shipped"
    lbl(36).Caption = "Branch"
    Me.Caption = "Production Order "
    lbl(42).Caption = "Customer"
    Label27.Caption = "Recive Date"
    lbl(53).Caption = "Batch No."
    Label26.Caption = "Indirect Cost According To Percenrage"
lbl(44).Caption = "This Screen Allow to Create Production Order and Calculate Cost Automatically According To Issue Vouchers"
    With CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Credit"
    End With
lbl(55).Caption = "Based On"
'////////mA
lbl(82).Caption = "Select Driver"
lbl(50).Caption = "Location"
lbl(47).Caption = "Receipt"
lbl(45).Caption = "Plan"
lbl(46).Caption = "Manul No."
lbl(49).Caption = "Mix Code"
lbl(52).Caption = "Supervisor"


    ELe(6).Caption = Me.Caption
    lbl(5).Caption = "Order No"
    lbl(32).Caption = "Total Qty"
    lbl(6).Caption = "Date"
'    lbl(17).Caption = "Sales Order No."
    lbl(33).Caption = "ROM Store"
    lbl(34).Caption = "Finish Goods Store"

    Label9.Caption = "Remarks"
    lbl(28).Caption = "Prod Start"
    lbl(35).Caption = "Prod End"

    lbl(27).Caption = "Qty"

    lbl(13).Caption = "Country"
    lbl(14).Caption = "Shipment Mode"
    lbl(21).Caption = "Credit Curr."
 
    lbl(23).Caption = "Value"
    'ISButton1.Caption = "Show Port Data"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "item name"

    lbl(29).Caption = "Status"
    lbl(19).Caption = "Qty"
    lbl(26).Caption = "Price"

    lbl(3).Caption = "Total R.O.M."
    lbl(1).Caption = "By"
    lbl(0).Caption = "Currenr rec."
    lbl(2).Caption = "Total rec."

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Me.XPTab301.TabCaption(0) = "Production Items"
    Me.XPTab301.TabCaption(1) = "ROMl Items"
    Me.XPTab301.TabCaption(2) = "Line Expenese"
    Me.XPTab301.TabCaption(3) = "Worker Expenses"
    Me.XPTab301.TabCaption(4) = "Fn inv  And Expenses VCHR"
 
    Me.XPTab301.TabCaption(5) = "Estimates Expenses"
    Me.XPTab301.TabCaption(6) = " Linked voucher"
    Me.XPTab301.TabCaption(7) = " Issue VCHR "
    Me.XPTab301.TabCaption(8) = " Estimatd Cost"

    Label4.Caption = "Raw Of Material Items"
    Label10.Caption = "Raw Of Material  Total"

    With Me.FG1
        .TextMatrix(0, .ColIndex("Code")) = "Item Code "
        .TextMatrix(0, .ColIndex("Name")) = "Item Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name "

        .TextMatrix(0, .ColIndex("Valu")) = " Value "
        .TextMatrix(0, .ColIndex("TotalQty")) = "TotalQty"

        .TextMatrix(0, .ColIndex("Count")) = "Qty"
        .TextMatrix(0, .ColIndex("Cost")) = "Cost "
        .TextMatrix(0, .ColIndex("Total")) = "Total"
       
    End With
    
    GridWorker.TextMatrix(0, GridWorker.ColIndex("shift")) = "shift"
    
    Label17.Caption = "Hours"
    lbl(41).Caption = "To"

    Label19.Caption = "Estimated Expenses"
    Cmd(9).Caption = "Remove Line"
    Label18.Caption = "Total"

    With Me.FG

        .TextMatrix(0, .ColIndex("EstimatedCost")) = "Estimated Cost "

        .TextMatrix(0, .ColIndex("Expenses")) = "Expenses"
        .TextMatrix(0, .ColIndex("DistibutePercentage")) = "Distibute %"
        
        .TextMatrix(0, .ColIndex("VoucheRecev")) = "Create Resieve Voucher  "
        .TextMatrix(0, .ColIndex("ReceiveSerial")) = "No. Resieve Voucher  "
        .TextMatrix(0, .ColIndex("ShowReceiv")) = "Show Resieve Voucher  "
        .TextMatrix(0, .ColIndex("RecevGl")) = "Show JE Resieve "
        .TextMatrix(0, .ColIndex("Voucher")) = "Create Issue Voucher  "
        .TextMatrix(0, .ColIndex("IssueSerial")) = "No. Issue Voucher  "
        .TextMatrix(0, .ColIndex("ShowIssue")) = "Show Issue Voucher  "
        .TextMatrix(0, .ColIndex("IssuGl")) = "Show JE Issue "
        

    End With

    With Me.GridEstimatedCost
        .TextMatrix(0, .ColIndex("ElementName")) = "ElementName"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name "
        .TextMatrix(0, .ColIndex("Value1")) = "cost "

        .TextMatrix(0, .ColIndex("CurrencyName")) = "CurrencyName"
        .TextMatrix(0, .ColIndex("Rate")) = "Rate "
        .TextMatrix(0, .ColIndex("Count")) = "Count "

        .TextMatrix(0, .ColIndex("Value")) = "unit cost"
        .TextMatrix(0, .ColIndex("Total")) = "Total"

        .TextMatrix(0, .ColIndex("LineNo")) = "Ser"

    End With

    With Me.GridIssueVoucer
  
        .TextMatrix(0, .ColIndex("noteserial1")) = "VCHR NO"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE NO"
        .TextMatrix(0, .ColIndex("code")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
        .TextMatrix(0, .ColIndex("count")) = "Qty"
        .TextMatrix(0, .ColIndex("cost")) = "cost"
        .TextMatrix(0, .ColIndex("total")) = "total"
        .TextMatrix(0, .ColIndex("Ser")) = "S"
 
    End With

    With Me.Fg_Journal
   
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Des")) = "Remarks"
 
    End With

    Label15.Caption = "Financial Invoices And Expenses Vouchers"
    lbl(54).Caption = "Expenses VCHR"
    lbl(38).Caption = "FIN INV."

    lbl(51).Caption = "Expenses VCHR Total"
    lbl(60).Caption = "FIN INV. Total"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "VCHR NO. "
        .TextMatrix(0, .ColIndex("Note_Value")) = "value"
        .TextMatrix(0, .ColIndex("name")) = "Expenses Name"
    End With

    With Me.Grid4
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "INV NO. "
        .TextMatrix(0, .ColIndex("Note_Value")) = "value"
        .TextMatrix(0, .ColIndex("name")) = "Account Name"
    End With

    Label12.Caption = "Worker Expenses"
    Cmd(8).Caption = "Delete Row"
    Label13.Caption = "Total Worker Cost "
    Label32.Caption = "Total Worker Cost Per Hour"
    With Me.GridWorker
        .TextMatrix(0, .ColIndex("LineNo")) = "i"
        .TextMatrix(0, .ColIndex("code")) = "Emp Code "
        .TextMatrix(0, .ColIndex("name")) = "Emp Name "
        .TextMatrix(0, .ColIndex("hourprice")) = "hour price"
        .TextMatrix(0, .ColIndex("from")) = "from"
        .TextMatrix(0, .ColIndex("to")) = "to"
        .TextMatrix(0, .ColIndex("Hour")) = "Hour"
        .TextMatrix(0, .ColIndex("total")) = "total"
    End With

    Check1.Caption = "Work With Product Line"
    lbl(25).Caption = "Select Line"
    Label1(12).Caption = "Shift"
    lbl(40).Caption = "From"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"
    Label11.Caption = "Total Expenses In One Hour"
    Label31.Caption = "Total Used Power"
    Label30.Caption = "Total Used Electricity "
    Label29.Caption = "Total Depreciation"
    With Me.FGLine
        .TextMatrix(0, .ColIndex("Hourdipp")) = "Depreciation Value"
        .TextMatrix(0, .ColIndex("Ser")) = "i"
        .TextMatrix(0, .ColIndex("code")) = "Line Code "
        .TextMatrix(0, .ColIndex("name")) = "Line Name "
        .TextMatrix(0, .ColIndex("UsedPowerPriceH")) = "Used Power Price H"
        .TextMatrix(0, .ColIndex("UsedElectricPriceH")) = "Used Electricity Price H"

        .TextMatrix(0, .ColIndex("from")) = "from"
        .TextMatrix(0, .ColIndex("to")) = "to"
        .TextMatrix(0, .ColIndex("Hour")) = "Hour"
        .TextMatrix(0, .ColIndex("total")) = "total"
    End With
 
    Label15.Caption = "Specify Vouchers"

    Cmd(9).Caption = "Delete Row"
    Label18.Caption = "Total"
    Label28.Caption = "Data of Mix.Items "
    With FgMix
    .TextMatrix(0, .ColIndex("MixCode")) = "Mix Code"
    .TextMatrix(0, .ColIndex("MainName")) = "Main Items"
    .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
    .TextMatrix(0, .ColIndex("Code")) = "Code"
    .TextMatrix(0, .ColIndex("Name")) = "Item Name"
    .TextMatrix(0, .ColIndex("Count")) = "Original Qty"
    .TextMatrix(0, .ColIndex("QtyMix")) = "Mix.Qty"
    .TextMatrix(0, .ColIndex("Qty")) = "Qty"
    .TextMatrix(0, .ColIndex("Cost")) = "Cost"
    .TextMatrix(0, .ColIndex("Valu")) = "Value"
    .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name"
    End With
    Label14.Caption = "Total"
    Label19.Caption = "Estimated Expenses"

    lbl(39).Caption = "Create Issue And Recive Vouchers"
    CmdIssueVoucher.Caption = "Create Issue Voucher"
    CmdResiveVoucher.Caption = "Create Resieve  Voucher"
    Label20.Caption = "NO"
    Label16.Caption = "NO"
    Command3.Caption = "View VCHR"
    Command4.Caption = "View VCHR"

    Command5.Caption = "View JE"
    Command7.Caption = "View JE"
    CmdConvert.Caption = "Convert To Bill"
    CmdTemplate.Caption = "Insert template"

End Sub

Function FillGroupExpenses(GroupID As Integer, Qty As Double)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
 Me.TxtTotalEstimatedCost.Text = 0
    StrSQL = "SELECT     dbo.UnitsIndustrialCost.CurrencyID, dbo.UnitsIndustrialCost.unitid, dbo.UnitsIndustrialCostDetails.TBLProductionElementsId, dbo.UnitsIndustrialCostDetails.Cost, "
    StrSQL = StrSQL + "  dbo.TBLProductionElements.Name, dbo.TBLProductionElements.Namee, dbo.TBLProductionElements.ExpensesID, dbo.ExpensesType.ID,"
    StrSQL = StrSQL + "  dbo.ExpensesType.Name AS ExpensesName, dbo.ExpensesType.Account_Code, dbo.currency.code, dbo.currency.name AS CurrencyName, dbo.currency.rate"
    StrSQL = StrSQL + "  FROM         dbo.UnitsIndustrialCostDetails INNER JOIN"
    StrSQL = StrSQL + "  dbo.UnitsIndustrialCost ON dbo.UnitsIndustrialCostDetails.UnitsIndustrialCostId = dbo.UnitsIndustrialCost.id INNER JOIN"
    StrSQL = StrSQL + "  dbo.TBLProductionElements ON dbo.UnitsIndustrialCostDetails.TBLProductionElementsId = dbo.TBLProductionElements.TBLProductionElementsId INNER JOIN"
    StrSQL = StrSQL + "  dbo.ExpensesType ON dbo.TBLProductionElements.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
    StrSQL = StrSQL + "  dbo.currency ON dbo.UnitsIndustrialCost.CurrencyID = dbo.currency.id"
 
    StrSQL = "SELECT     dbo.UnitsIndustrialCost.CurrencyID, dbo.UnitsIndustrialCost.unitid, dbo.UnitsIndustrialCostDetails.TBLProductionElementsId, dbo.UnitsIndustrialCostDetails.Cost, "
    StrSQL = StrSQL + "   dbo.TBLProductionElements.Name, dbo.TBLProductionElements.Namee, dbo.TBLProductionElements.ExpensesID, dbo.ExpensesType.ID,"
    StrSQL = StrSQL + "   dbo.ExpensesType.Name AS ExpensesName, dbo.ExpensesType.Account_Code, dbo.currency.code, dbo.currency.name AS CurrencyName, dbo.currency.rate,"
    StrSQL = StrSQL + "   dbo.Groups.GroupName"
    StrSQL = StrSQL + "   FROM         dbo.UnitsIndustrialCostDetails INNER JOIN"
    StrSQL = StrSQL + "   dbo.UnitsIndustrialCost ON dbo.UnitsIndustrialCostDetails.UnitsIndustrialCostId = dbo.UnitsIndustrialCost.id INNER JOIN"
    StrSQL = StrSQL + "   dbo.TBLProductionElements ON dbo.UnitsIndustrialCostDetails.TBLProductionElementsId = dbo.TBLProductionElements.TBLProductionElementsId INNER JOIN"
    StrSQL = StrSQL + "   dbo.ExpensesType ON dbo.TBLProductionElements.ExpensesID = dbo.ExpensesType.ID INNER JOIN"
    StrSQL = StrSQL + "   dbo.currency ON dbo.UnitsIndustrialCost.CurrencyID = dbo.currency.id INNER JOIN"
    StrSQL = StrSQL + "   dbo.Groups ON dbo.UnitsIndustrialCost.unitid = dbo.Groups.GroupID"
    StrSQL = StrSQL + "   WHERE     (dbo.UnitsIndustrialCost.unitid = " & GroupID & ")"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
 
        row_count = GridEstimatedCost.Rows
    
        If GridEstimatedCost.TextMatrix(row_count - 1, GridEstimatedCost.ColIndex("ElementId")) = "" Then
            row_count = row_count - 1
        End If
     
        GridEstimatedCost.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To GridEstimatedCost.Rows - 1 'RsDetails.RecordCount
    
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("LineNo")) = Num
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("GroupID")) = IIf(IsNull(RsDetails("unitid")), 0, (RsDetails("unitid").value))
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("GroupName")) = IIf(IsNull(RsDetails("GroupName")), "", (RsDetails("GroupName").value))
           
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("ElementId")) = IIf(IsNull(RsDetails("TBLProductionElementsId")), "", (RsDetails("TBLProductionElementsId").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("ElementName")) = IIf(IsNull(RsDetails("Name")), "", (RsDetails("Name").value))
            Else
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("ElementName")) = IIf(IsNull(RsDetails("Namee")), "", (RsDetails("Namee").value))
            End If

            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("ExpensesID")) = IIf(IsNull(RsDetails("ExpensesID")), "", (RsDetails("ExpensesID").value))
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AccountName")) = IIf(IsNull(RsDetails("ExpensesName")), "", Trim(RsDetails("ExpensesName").value))
        
            '          GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Count")) = items_qty_not_recieved_in_order(GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Code")), GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("order_no")))
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("AccountCode")) = IIf(IsNull(RsDetails("Account_Code")), "", (RsDetails("Account_Code").value))
         
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Value1")) = IIf(IsNull(RsDetails("cost")), "", (RsDetails("cost").value))
        
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("CurrencyId")) = IIf(IsNull(RsDetails("CurrencyId")), "", (RsDetails("CurrencyId").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("CurrencyName")) = IIf(IsNull(RsDetails("CurrencyName")), "", (RsDetails("CurrencyName").value))
            Else
                GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("CurrencyName")) = IIf(IsNull(RsDetails("Code")), "", (RsDetails("Code").value))
            End If

            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Rate")) = IIf(IsNull(RsDetails("Rate")), "", (RsDetails("Rate").value))
 
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("value")) = GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Rate")) * GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Value1"))
  
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Count")) = Qty
            GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("Total")) = Round(val(GridEstimatedCost.TextMatrix(Num, GridEstimatedCost.ColIndex("value"))) * Qty, SystemOptions.SysDefQuantityDecimal)
         
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If GridEstimatedCost.Rows > 10 Then
            '     If Num = 8 Then GridEstimatedCost.Refresh
            ' End If
        Next Num

        With GridEstimatedCost

            If .Rows > 1 Then
                Me.TxtTotalEstimatedCost.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
            Else
                Me.TxtTotalEstimatedCost.Text = 0
            End If

        End With

    End If

End Function

Function Retrive_orders_data(WorkOrderNO As String)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    'StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
    '"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    ''StrSQL = StrSQL + " where Transaction_Type=19  and WorkOrderNO=" & WorkOrderNO

    StrSQL = "SELECT    dbo.Transactions.NoteSerial,  dbo.Transactions.NoteSerial1,    dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, "
    StrSQL = StrSQL + " dbo.Transactions.WorkOrderNO, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL + " dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ShowQty,showPrice, dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL + " dbo.Transaction_Details.Price , dbo.TblUnites.UnitName"
    StrSQL = StrSQL + "  ,ShowQty*showPrice  as Costs FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + "  WHERE     (dbo.Transactions.Transaction_Type = 27) AND (dbo.Transactions.WorkOrderNO = '" & WorkOrderNO & "')"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
    GridIssueVoucer.Rows = 1

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        row_count = GridIssueVoucer.Rows
    
        If GridIssueVoucer.TextMatrix(row_count - 1, GridIssueVoucer.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        GridIssueVoucer.Clear flexClearScrollable, flexClearEverything
        GridIssueVoucer.Rows = 1
        GridIssueVoucer.Enabled = True

        GridIssueVoucer.Rows = RsDetails.RecordCount + 1

        For Num = row_count To GridIssueVoucer.Rows - 1 'RsDetails.RecordCount
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Ser")) = Num
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Transaction_ID")) = IIf(IsNull(RsDetails("Transaction_ID")), "", (RsDetails("Transaction_ID").value))
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("noteserial1")) = IIf(IsNull(RsDetails("noteserial1")), "", (RsDetails("noteserial1").value))
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("noteserial")) = IIf(IsNull(RsDetails("noteserial")), "", (RsDetails("noteserial").value))
       
            '        GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Item_ID")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
        
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
        
            '          GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Count")) = items_qty_not_recieved_in_order(GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Code")), GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("order_no")))
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
           
            '         GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Cost")) = IIf(IsNull(RsDetails("showPrice")), 0, (RsDetails("showPrice").value)) '* IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            '         GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            '         GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("SizeID")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            '          GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
            '          GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Total")) = val(GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Count"))) * val(GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Cost")))
          
            GridIssueVoucer.TextMatrix(Num, GridIssueVoucer.ColIndex("Total")) = IIf(IsNull(RsDetails("Costs")), 0, (RsDetails("Costs").value)) '* IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("Costs").value))
         
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If GridIssueVoucer.Rows > 10 Then
            '     If Num = 8 Then GridIssueVoucer.Refresh
            ' End If
        Next Num

        GridIssueVoucer.AutoSize 0, GridIssueVoucer.Cols - 1, False
    End If
 
    With GridIssueVoucer

        If .Rows > 1 Then
            TXTTotalIssueVouchers = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows, .ColIndex("Total"))
        Else
            TXTTotalIssueVouchers = 0
        End If
    
    End With

End Function

Public Function add_item_to_parts_grid(ItemID As Long, _
                                       itemcode As String, _
                                       ItemName As String, _
                                       cost As Variant, _
                                       Qty As Double, _
                                       productQty As Double, Optional UnitID As Integer)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
    LngNewRow = ModFgLib.SetFgForNewRow(FG1, FG1.ColIndex("Code"))

    StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
    StrSQL = StrSQL + " Where  TblUnites.UnitID=" & val(UnitID)
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    Dim UnitName As String

    If Not (rs.BOF Or rs.EOF) Then
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
        UnitName = IIf(IsNull(rs("UnitName").value), 0, rs("UnitName").value)
    End If

    With Me.FG1
        .TextMatrix(LngNewRow, .ColIndex("id")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("code")) = itemcode
        .TextMatrix(LngNewRow, .ColIndex("Name")) = ItemName
        .TextMatrix(LngNewRow, .ColIndex("count")) = Qty
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = UnitID
        .TextMatrix(LngNewRow, .ColIndex("Unitname")) = UnitName
        .TextMatrix(LngNewRow, .ColIndex("Cost")) = cost
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = cost * Qty
        .TextMatrix(LngNewRow, .ColIndex("TotalQty")) = productQty * Qty
        .TextMatrix(LngNewRow, .ColIndex("Total")) = productQty * cost * Qty
    
        .AutoSize 0, .Cols - 1, False
   
        If .Rows > 1 Then
            Me.TxtTotalMaterials.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
            Me.txtCount.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Count"), .Rows - 1, .ColIndex("Count"))
            Me.TxtTotalQty.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalQty"), .Rows - 1, .ColIndex("TotalQty"))
            
        Else
            Me.TxtTotalMaterials.Text = 0
                 Me.txtCount.Text = 0
                      Me.TxtTotalQty.Text = 0
        End If

    End With

End Function

Public Function add_item_to_parts_grid2(Optional ItemID As Long, _
                                       Optional itemcode As String, _
                                      Optional ItemName As String, _
                                       Optional cost As Long, _
                                       Optional Qty As Long, _
                                       Optional productQty As Long, Optional UnitID As Integer)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
    'LngNewRow = ModFgLib.SetFgForNewRow(Fg1, Fg1.ColIndex("Code"))
    StrSQL = "SELECT     dbo.TblDefComItem.UnitID, dbo.TblDefComItem.ItemNameID, dbo.TblItems.*, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL + "  , dbo.TblDefComItem.CusID,  dbo.TblDefComItem.StoreId "
StrSQL = StrSQL + "   FROM         dbo.TblDefComItem INNER JOIN"
StrSQL = StrSQL + "                       dbo.TblItems ON dbo.TblDefComItem.ItemNameID = dbo.TblItems.ItemID INNER JOIN"
StrSQL = StrSQL + "                       dbo.TblUnites ON dbo.TblDefComItem.UnitID = dbo.TblUnites.UnitID"
                      
 
StrSQL = StrSQL + "  Where (ID = " & val(txtMixID.Text) & ")"

 Dim i As Integer
    Dim RsDetails As New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
   FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
 
    'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
Dim Num As Integer
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1
DBCboClientName.BoundText = IIf(IsNull(RsDetails("CusID")), "", (RsDetails("CusID").value))
'DBCboClientName.BoundText =
 
        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemNameID")), "", (RsDetails("ItemNameID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNameID")), "", Trim(RsDetails("ItemNameID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
            FG.TextMatrix(Num, FG.ColIndex("Price")) = 0
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 1
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = 0
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            If SystemOptions.UserInterface = ArabicInterface Then
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            
            'UnitNamee
            End If
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = 1
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = 1
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = 1
        
            RsDetails.MoveNext
             

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
    
 
End Function

Public Function add_item_to_parts_grid1(Optional ItemID As Long, _
                                       Optional itemcode As String, _
                                      Optional ItemName As String, _
                                       Optional cost As Long, _
                                       Optional Qty As Long, _
                                       Optional productQty As Long, Optional UnitID As Integer)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
  If val(txtMixID.Text) = 0 Then Exit Function
    LngNewRow = ModFgLib.SetFgForNewRow(FG1, FG1.ColIndex("Code"))

  '  StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
  '  StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
  '  StrSQL = StrSQL + " Where  TblUnites.UnitID=" & val(unitid)
    
   StrSQL = "SELECT     dbo.TblDefComItemDet.ItemID, dbo.TblDefComItemDet.UnitID, dbo.TblDefComItemDet.IDDefCIT, dbo.TblDefComItemDet.Qty, dbo.TblItems.ItemID, "
  StrSQL = StrSQL + "  dbo.TblItems.itemcode , dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee"
 StrSQL = StrSQL + " FROM         dbo.TblDefComItemDet INNER JOIN"
 StrSQL = StrSQL + " dbo.TblItems ON dbo.TblDefComItemDet.ItemID = dbo.TblItems.ItemID INNER JOIN"
 StrSQL = StrSQL + " dbo.TblUnites ON dbo.TblDefComItemDet.UnitID = dbo.TblUnites.UnitID"
 StrSQL = StrSQL + " WHERE     (dbo.TblDefComItemDet.IDDefCIT = " & val(txtMixID.Text) & ")"
productQty = val(LblTotalQty.Caption)
Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
       FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 2
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Refresh
    
    Dim UnitName As String
Dim item_cost As Double
    If Not (rs.BOF Or rs.EOF) Then
    FG1.Rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
    LngNewRow = i
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
        UnitName = IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value)
        ItemID = IIf(IsNull(rs("ItemID").value), 0, rs("ItemID").value)
        itemcode = IIf(IsNull(rs("itemcode").value), "", rs("itemcode").value)
        ItemName = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
        Qty = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
    cost = ModItemCostPrice.GetCostItemPrice(rs("ItemID").value, 0, , , SystemOptions.SysMainStockCostMethod, , , , , rs("UnitID").value)
    
       With Me.FG1
        .TextMatrix(LngNewRow, .ColIndex("id")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("code")) = itemcode
        .TextMatrix(LngNewRow, .ColIndex("Name")) = ItemName
        .TextMatrix(LngNewRow, .ColIndex("count")) = Qty
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = UnitID
        .TextMatrix(LngNewRow, .ColIndex("Unitname")) = UnitName
        .TextMatrix(LngNewRow, .ColIndex("Cost")) = cost
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = cost * Qty
        .TextMatrix(LngNewRow, .ColIndex("TotalQty")) = productQty * Qty
        .TextMatrix(LngNewRow, .ColIndex("Total")) = productQty * cost * Qty
    
        .AutoSize 0, .Cols - 1, False
   
        If .Rows > 1 Then
            Me.TxtTotalMaterials.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
        Else
            Me.TxtTotalMaterials.Text = 0
        End If

    End With
    
    rs.MoveNext
    Next i
    End If



End Function
Public Function FIllEstimatedExpenses()
    Dim Item_ID As Integer
    Dim GroupID As Integer
    Dim RowNum As Integer
    Dim EstimatedCost As Double
 
    Dim LngUnitID As Long
    Dim UnitFactor As Double
      
    GridEstimatedCost.Clear flexClearScrollable, flexClearEverything
    GridEstimatedCost.Rows = 1
          
    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            Item_ID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            GetUnitNoOfItems Item_ID, LngUnitID, UnitFactor
            GetItemData Item_ID, , , GroupID
            FillGroupExpenses GroupID, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) * UnitFactor
            EstimatedCost = 0
            GetEstimatedCost , GroupID, EstimatedCost
      
            FG.TextMatrix(RowNum, FG.ColIndex("EstimatedCost")) = EstimatedCost * UnitFactor
        
        End If
        
    Next RowNum

End Function

Public Function show_parts()
 
    Dim RowNum As Integer
    FG1.Clear flexClearScrollable, flexClearEverything
    FG1.Rows = 2
          
    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If add_part_item(val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))) Then
        
            End If
        End If

    Next RowNum

End Function

Public Function add_part_item(LngItemID As Long, _
                              Optional Qty As Double) As Boolean
    '131315
    Dim StrSQL As String
    Dim RsParts As ADODB.Recordset
    Dim i As Integer
  
    StrSQL = "SELECT  dbo.TblItemsParts.Unitid,  dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.TableID   ,dbo.TblItems.ItemName, dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.ItemID, dbo.TblItems.ItemCode"
    StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN "
    StrSQL = StrSQL + " dbo.TblItemsParts ON dbo.TblItems.ItemID = dbo.TblItemsParts.PartItemID"
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID
    StrSQL = StrSQL + " Order By TableID"
    Dim item_cost As Variant
    Set RsParts = New ADODB.Recordset
    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsParts.EOF Or RsParts.BOF) Then

        For i = 0 To RsParts.RecordCount - 1
               
            item_cost = ModItemCostPrice.GetCostItemPrice(RsParts("PartItemID").value, 0, , , SystemOptions.SysMainStockCostMethod, , , , , RsParts("Unitid").value)

            If add_item_to_parts_grid(val(RsParts("PartItemID").value), RsParts("ItemCode").value, RsParts("ItemName").value, item_cost, val(RsParts("PartItemQty").value), Qty, val(RsParts("Unitid").value)) = True Then
            End If
                  
            RsParts.MoveNext
        Next i

    End If

End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Me.Txt_EXport.Text = Expenses_update_total
    cal_expenses
End Sub

Function Expenses_update_total() As Long
    Dim i As Integer
    On Error Resume Next

    If Grid.Rows = 1 Then Exit Function
    Expenses_update_total = 0

    For i = 1 To Grid.Rows - 1
        
        If Grid.Cell(flexcpChecked, i, Grid.ColIndex("select")) = flexChecked Then
            Expenses_update_total = Expenses_update_total + val(Grid.TextMatrix(i, Grid.ColIndex("note_value")))
        End If

    Next i
   
End Function

Function fillFinancialInvoiceGrid()

    With Me.Grid4
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        '
        '    .AutoSize 0, .Cols - 1, False
    End With

If TxtTransSerial.Text = "" Then
Exit Function
End If


    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.Item_id,dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3 and order_no='" & Me.TXT_order_no.text & "' " & "AND (ITEM_ID=" & Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code"))) & " or  ITEM_ID is null)  and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & "))  "
    'My_SQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
    'My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
    'My_SQL = My_SQL + " dbo.Notes.order_no, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID ,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.buy,dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 "
    'My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
    'My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    'My_SQL = My_SQL + "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    'My_SQL = My_SQL + " WHERE      (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 is null or dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ") and  (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) "

  '  My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3   and order_no='" & Me.TXT_order_no.text & "' and(Transaction_ID1 is null or Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")  )  "

  '  My_SQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
  '  My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
  '  My_SQL = My_SQL + " dbo.Notes.order_no, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID ,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.buy,dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 "
  '  My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
  '  My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
  '  My_SQL = My_SQL + "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
  '  'My_SQL = My_SQL + " WHERE      (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 is null or dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ") and  (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) "
  '  My_SQL = My_SQL + " WHERE    dbo.Notes.NoteType = 80 and BasedONID=2  and    dbo.Notes.order_no='" & TxtTransSerial.text & "'"


My_SQL = " SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
 My_SQL = My_SQL + "  dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
     My_SQL = My_SQL + " dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID, dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,"
    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.buy , dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1, dbo.notes_all.BasedONID"
    My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    My_SQL = My_SQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
    My_SQL = My_SQL + " dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
    My_SQL = My_SQL + " WHERE     (dbo.Notes.NoteType = 80) AND (dbo.Notes.ORDER_NO = '" & TxtTransSerial.Text & "') AND (dbo.notes_all.BasedONID = 3)"

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset

    With Me.Grid4
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")) = IIf(IsNull(RsExp.Fields("Double_Entry_Vouchers_ID").value), 0, RsExp.Fields("Double_Entry_Vouchers_ID").value)
           
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
 
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
                End If
 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
 
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
 
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Value").value), "", RsExp.Fields("Value").value)
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
 
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If

                .TextMatrix(i, .ColIndex("Select")) = 1
 
                ' .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("buy").value), _
                  0, RsExp.Fields("buy").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid4.Visible = True

    ' End If
  
    'update_finincial_invoice_total

End Function

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "ItemName" Then
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

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Grid

        Select Case .ColKey(Col)

            Case "ItemName"
       
                StrSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                '     StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

