VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmBarcodePrinting 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
   ClientHeight    =   9060
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   16800
   HelpContextID   =   580
   Icon            =   "FrmBarcodePrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   16800
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
      Height          =   9105
      Left            =   -120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16995
      _cx             =   29977
      _cy             =   16060
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
      AutoSizeChildren=   0
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8550
         Left            =   30
         TabIndex        =   1
         Top             =   -90
         Width           =   17025
         _cx             =   30030
         _cy             =   15081
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
            Height          =   8130
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   16935
            _cx             =   29871
            _cy             =   14340
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
               Height          =   525
               Index           =   5
               Left            =   0
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Width           =   16875
               _cx             =   29766
               _cy             =   926
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
               Picture         =   "FrmBarcodePrinting.frx":038A
               Caption         =   ""
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
                  TabIndex        =   27
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
                  ButtonImage     =   "FrmBarcodePrinting.frx":1064
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
                  ButtonImage     =   "FrmBarcodePrinting.frx":13FE
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
                  ButtonImage     =   "FrmBarcodePrinting.frx":1798
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
                  ButtonImage     =   "FrmBarcodePrinting.frx":1B32
                  ColorHighlight  =   4194304
                  ColorHoverText  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
                  ColorToggledHoverText=   16777215
                  ColorTextShadow =   16777215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
                  Height          =   360
                  Index           =   12
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   120
                  Width           =   5025
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   8115
               Index           =   1
               Left            =   -120
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   17025
               _cx             =   30030
               _cy             =   14314
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
               Begin VB.CheckBox Check17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ / «·€«¡ «·ﬂ·"
                  Height          =   195
                  Left            =   14520
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   4560
                  Width           =   2295
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
                  Height          =   525
                  Left            =   360
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   32
                  Top             =   660
                  Width           =   9960
               End
               Begin VB.TextBox TxtSerial 
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
                  Left            =   14280
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   720
                  Width           =   1800
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
                  TabIndex        =   9
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
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin MSComCtl2.DTPicker dbRecordDate 
                  Height          =   285
                  Left            =   11400
                  TabIndex        =   10
                  Top             =   690
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   93585409
                  CurrentDate     =   38784
               End
               Begin ALLButtonS.ALLButton CmdRemove1 
                  Height          =   255
                  Left            =   15960
                  TabIndex        =   33
                  Tag             =   "Delete Row"
                  Top             =   7800
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
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
                  MICON           =   "FrmBarcodePrinting.frx":1ECC
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                  Height          =   2535
                  Left            =   8640
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   8325
                  _cx             =   14684
                  _cy             =   4471
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
                  Begin VB.CheckBox chkhalf 
                     Alignment       =   1  'Right Justify
                     Caption         =   "‰’ «·ﬂ„ÌÂ"
                     Height          =   255
                     Left            =   6720
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   120
                     Width           =   1215
                  End
                  Begin VB.ListBox ListAllActivity 
                     Height          =   1815
                     ItemData        =   "FrmBarcodePrinting.frx":1EE8
                     Left            =   4620
                     List            =   "FrmBarcodePrinting.frx":1EEF
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   465
                     Width           =   3585
                  End
                  Begin VB.ListBox ListActivitySelected 
                     Height          =   1815
                     ItemData        =   "FrmBarcodePrinting.frx":1EF9
                     Left            =   75
                     List            =   "FrmBarcodePrinting.frx":1F00
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   465
                     Width           =   3585
                  End
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   480
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   1245
                     Width           =   780
                  End
                  Begin VB.Label Label6 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   495
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   1710
                     Width           =   780
                  End
                  Begin VB.Label Label7 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   330
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   930
                     Width           =   780
                  End
                  Begin VB.Label Label8 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   480
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   465
                     Width           =   780
                  End
                  Begin VB.Label Label12 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·„Ã„Ê⁄« "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   285
                     Left            =   3150
                     TabIndex        =   37
                     Top             =   120
                     Width           =   2325
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                  Height          =   2535
                  Left            =   120
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   8325
                  _cx             =   14684
                  _cy             =   4471
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
                  Begin VB.ListBox ListStoreSelected 
                     Height          =   1815
                     ItemData        =   "FrmBarcodePrinting.frx":1F0F
                     Left            =   75
                     List            =   "FrmBarcodePrinting.frx":1F16
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   465
                     Width           =   3585
                  End
                  Begin VB.ListBox ListAllStore 
                     Height          =   1815
                     ItemData        =   "FrmBarcodePrinting.frx":1F2D
                     Left            =   4620
                     List            =   "FrmBarcodePrinting.frx":1F34
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   465
                     Width           =   3585
                  End
                  Begin VB.Label Label10 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õœœ «·„Œ«“‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   285
                     Left            =   2910
                     TabIndex        =   50
                     Top             =   120
                     Width           =   2325
                  End
                  Begin VB.Label Label9 
                     Alignment       =   2  'Center
                     Caption         =   ">"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   480
                     Left            =   3720
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   465
                     Width           =   780
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
                     Caption         =   ">>"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   330
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   930
                     Width           =   780
                  End
                  Begin VB.Label Label3 
                     Alignment       =   2  'Center
                     Caption         =   "<<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   495
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   1710
                     Width           =   780
                  End
                  Begin VB.Label Label2 
                     Alignment       =   2  'Center
                     Caption         =   "<"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   480
                     Left            =   3750
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   1245
                     Width           =   780
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   615
                  Left            =   8640
                  TabIndex        =   51
                  TabStop         =   0   'False
                  Top             =   3960
                  Width           =   8325
                  _cx             =   14684
                  _cy             =   1085
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
                  Begin VB.TextBox TxtCodeAother 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   6510
                     TabIndex        =   53
                     Top             =   120
                     Width           =   1065
                  End
                  Begin MSDataListLib.DataCombo Dcbiteem 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   54
                     Top             =   120
                     Width           =   6375
                     _ExtentX        =   11245
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·’‰›"
                     Height          =   315
                     Index           =   0
                     Left            =   7440
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   120
                     Width           =   720
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2895
                  Left            =   240
                  TabIndex        =   52
                  Top             =   4800
                  Width           =   16680
                  _cx             =   29422
                  _cy             =   5106
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
                  Rows            =   15
                  Cols            =   27
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBarcodePrinting.frx":1F46
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   615
                  Left            =   240
                  TabIndex        =   56
                  TabStop         =   0   'False
                  Top             =   3960
                  Width           =   8325
                  _cx             =   14684
                  _cy             =   1085
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
                  Begin VB.TextBox TxtQtyPrint 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   5400
                     TabIndex        =   59
                     Top             =   120
                     Width           =   1425
                  End
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   615
                     Left            =   0
                     TabIndex        =   57
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   0
                     Width           =   4605
                     _ExtentX        =   8123
                     _ExtentY        =   1085
                     Caption         =   "«÷«›…"
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
                     ButtonImage     =   "FrmBarcodePrinting.frx":2329
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„Ì… «·ÿ»«⁄…"
                     Height          =   345
                     Index           =   2
                     Left            =   6900
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   150
                     Width           =   1185
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   3
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   780
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   525
                  Index           =   5
                  Left            =   13560
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   720
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   480
                  Index           =   7
                  Left            =   14940
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   720
                  Width           =   1785
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
                  TabIndex        =   7
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
         Height          =   600
         Left            =   150
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   8475
         Width           =   16785
         _cx             =   29607
         _cy             =   1058
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
            Height          =   405
            Left            =   21900
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   714
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
            ButtonImage     =   "FrmBarcodePrinting.frx":8B8B
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   405
            Left            =   23535
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   285
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   714
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
            ButtonImage     =   "FrmBarcodePrinting.frx":8F25
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   345
            Left            =   25740
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   195
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   609
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
            ButtonImage     =   "FrmBarcodePrinting.frx":92BF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   630
            Index           =   0
            Left            =   13155
            TabIndex        =   18
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   1
            Left            =   11490
            TabIndex        =   19
            Top             =   0
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   2
            Left            =   9960
            TabIndex        =   20
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   3
            Left            =   8100
            TabIndex        =   21
            Top             =   0
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   4
            Left            =   6195
            TabIndex        =   22
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   6
            Left            =   1485
            TabIndex        =   23
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1111
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
            Height          =   630
            Index           =   5
            Left            =   4485
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1111
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
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   630
            Left            =   3360
            TabIndex        =   60
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1111
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
            Height          =   255
            Left            =   2865
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   285
            Width           =   3225
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
            Height          =   270
            Left            =   9060
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   300
            Width           =   2805
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
      ButtonImage     =   "FrmBarcodePrinting.frx":9659
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmBarcodePrinting"
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
Dim FirstPeriodDateInthisYear As Date
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & TxtSerial.Text & CHR(13) & "   «· «—ÌŒ " & dbRecordDate & CHR(13) & "   „·«ÕŸ«  " & txtRemarks
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "No " & TxtSerial.Text & CHR(13) & "   Date " & dbRecordDate & CHR(13) & "   Remarks " & txtRemarks
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
  '  On Error GoTo ErrTrap
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblBarcodePrintingDet where BrcodID=" & val(Me.TxtSerial.Text)
    End If
    rs("id").value = val(TxtSerial.Text)
    rs("RecordDate").value = dbRecordDate.value
    rs("Remarks").value = IIf(Me.txtRemarks.Text = "", "", Me.txtRemarks.Text)
    rs("QtyPrint").value = val(TxtQtyPrint.Text)
    rs("UserID").value = user_id
    rs("ItemID").value = IIf(val(Dcbiteem.BoundText) = 0, Null, val(Dcbiteem.BoundText))
    rs.update
    ''///////////////////dbo.TblItems.PartNo
    Dim i As Integer
     Set RsDev = New ADODB.Recordset
    RsDev.Open "TblBarcodePrintingDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    With Me.Fg
        For i = 1 To .Rows - 1
            If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
                RsDev.AddNew
                RsDev("BrcodID").value = val(Me.TxtSerial.Text)
                RsDev("TransType").value = 0
                RsDev("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
                RsDev("GroupID").value = val(.TextMatrix(i, .ColIndex("GroupID")))
                RsDev("StoreID").value = val(.TextMatrix(i, .ColIndex("StoreID")))
                RsDev("SizeID").value = val(.TextMatrix(i, .ColIndex("SizeID")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("SortedID").value = val(.TextMatrix(i, .ColIndex("SortedID")))
                RsDev("ColorID").value = val(.TextMatrix(i, .ColIndex("ColorID")))
                RsDev("Qty").value = val(.TextMatrix(i, .ColIndex("Quantity")))
                RsDev("QtyPrint").value = val(.TextMatrix(i, .ColIndex("QtyPrint")))
                RsDev("ItemSerial").value = (.TextMatrix(i, .ColIndex("Serial")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("LotNO").value = (.TextMatrix(i, .ColIndex("LotNO")))
                RsDev("ExpiryDate").value = IIf(.TextMatrix(i, .ColIndex("ExpiryDate")) = "", Null, .TextMatrix(i, .ColIndex("ExpiryDate")))
                RsDev.update
            End If
        Next i
    End With
    
      Set RsDev = New ADODB.Recordset
    RsDev.Open "TblBarcodePrintingDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        For i = 0 To Me.ListActivitySelected.ListCount - 1
             RsDev.AddNew
             RsDev("BrcodID").value = val(Me.TxtSerial.Text)
             RsDev("ItemID").value = val(ListActivitySelected.ItemData(i))
             RsDev("TransType").value = 1
             RsDev.update
       Next i
       
          Set RsDev = New ADODB.Recordset
    RsDev.Open "TblBarcodePrintingDet", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        For i = 0 To Me.ListStoreSelected.ListCount - 1
             RsDev.AddNew
             RsDev("BrcodID").value = val(Me.TxtSerial.Text)
             RsDev("ItemID").value = val(ListStoreSelected.ItemData(i))
             RsDev("TransType").value = 2
             RsDev.update
       Next i
       
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

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Fg
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = True
            Next i

        End With

    Else

        With Me.Fg

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = False
            Next i

        End With

    End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
  
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.TxtSerial.Text = CStr(new_id("TblBarcodePrinting", "id", "", True))
            ListActivitySelected.Clear
            ListStoreSelected.Clear
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
         
        Case 1
  
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            CuurentLogdata

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans


        Case 6
            Unload Me

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap
    
    If TxtSerial.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            Cn.Execute "delete TblBarcodePrintingDet where BrcodID=" & val(Me.TxtSerial.Text)
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg.Clear flexClearScrollable, flexClearEverything
                    Fg.Rows = 2
                    Fg.Enabled = False
                
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

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

     
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


Private Sub CmdPrint_Click()
    Dim str As String

    Dim RowNum As Integer
    Dim ItemCount As Integer
    str = "Delete  TblPrintBarCode"
    Cn.Execute str

    'cBarcode.AddItem
    ' cBarcode.ClearItems
    For RowNum = 1 To Fg.Rows - 1

        If Fg.Cell(flexcpChecked, RowNum, Fg.ColIndex("Print")) = flexChecked Then
            If Not IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("QtyPrint"))) Then
     
          If chkhalf.value = vbChecked Then
      
                addtotable val(Fg.TextMatrix(RowNum, Fg.ColIndex("QtyPrint"))) / 2, Fg.TextMatrix(RowNum, Fg.ColIndex("barCodeNO")), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))), Fg.TextMatrix(RowNum, Fg.ColIndex("PartNo")), Fg.TextMatrix(RowNum, Fg.ColIndex("ItemName")), Fg.TextMatrix(RowNum, Fg.ColIndex("ColorName")), Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize")), Fg.TextMatrix(RowNum, Fg.ColIndex("ClassName")), Fg.TextMatrix(RowNum, Fg.ColIndex("LotNO")), Fg.TextMatrix(RowNum, Fg.ColIndex("ExpiryDate")) _
               , (Fg.TextMatrix(RowNum, Fg.ColIndex("Fullcode"))), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Quantity"))), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))), Fg.TextMatrix(RowNum, Fg.ColIndex("UnitName")), Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")), Fg.TextMatrix(RowNum, Fg.ColIndex("StoreName")), Fg.TextMatrix(RowNum, Fg.ColIndex("GroupName"))
                
          Else
                
                addtotable val(Fg.TextMatrix(RowNum, Fg.ColIndex("QtyPrint"))), Fg.TextMatrix(RowNum, Fg.ColIndex("barCodeNO")), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))), Fg.TextMatrix(RowNum, Fg.ColIndex("PartNo")), Fg.TextMatrix(RowNum, Fg.ColIndex("ItemName")), Fg.TextMatrix(RowNum, Fg.ColIndex("ColorName")), Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize")), Fg.TextMatrix(RowNum, Fg.ColIndex("ClassName")), Fg.TextMatrix(RowNum, Fg.ColIndex("LotNO")), Fg.TextMatrix(RowNum, Fg.ColIndex("ExpiryDate")) _
               , (Fg.TextMatrix(RowNum, Fg.ColIndex("Fullcode"))), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Quantity"))), val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))), Fg.TextMatrix(RowNum, Fg.ColIndex("UnitName")), Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")), Fg.TextMatrix(RowNum, Fg.ColIndex("StoreName")), Fg.TextMatrix(RowNum, Fg.ColIndex("GroupName"))
                
          
          
End If

            End If
        End If

    Next RowNum

    printCodes WindowTarget
    'Unload Me
End Sub
Function addtotable(NoOfRow As Integer, code As String, cost As Double, Optional PartNo As String = "", Optional Name As String = "" _
, Optional Color As String, Optional size As String, Optional Class As String, Optional lotNo As String, Optional ExpiryDate As String _
, Optional Fullcode As String, Optional Quantity As Double = 0, Optional Price As Double = 0, Optional UnitName As String, Optional serial As String, Optional StoreName As String, Optional GroupName As String)
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim i As Integer
    str = "select * from   TblPrintBarCode where 1=-1"
   rs.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 1 To NoOfRow
        rs.AddNew
        rs("PartNo").value = PartNo
        rs("code").value = code
        rs("code128").value = code128$(code)
         
        
        rs("cost").value = val(cost)
        rs("Name").value = Name
        rs("Color").value = Color
        rs("size").value = size
        rs("class").value = Class
        rs("LotNO").value = lotNo
        rs("ExpiryDate").value = IIf(ExpiryDate = "", Null, ExpiryDate)
        rs("Fullcode").value = Fullcode
        rs("Quantity").value = Quantity
        rs("Price").value = Price
        rs("UnitName").value = UnitName
        rs("Serial").value = serial
        rs("StoreName").value = StoreName
        rs("GroupName").value = GroupName
        rs.update
    Next i
'
End Function

Public Function code128$(chaine$)
  'Cette fonction est rÈgie par la Licence GÈnÈrale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'ParamËtres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichÈe avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramËtre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  code128$ = ""
  If Len(chaine$) > 0 Then
  'VÈrifier si caractËres valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(mId$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intÈressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au dÈbut ou ‡ la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'DÈbuter sur table C / Starting with table C
              code128$ = CHR$(210)
            Else 'Commuter sur table C / Switch to table C
              code128$ = code128$ & CHR$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then code128$ = CHR$(209) 'DÈbuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = val(mId$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            code128$ = code128$ & CHR$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            code128$ = code128$ & CHR$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caractËre en table B / Process 1 digit with table B
          code128$ = code128$ & mId$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clÈ de contrÙle / Calculation of the checksum
      For i% = 1 To Len(code128$)
        dummy% = Asc(mId$(code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clÈ / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Ajout de la clÈ et du STOP / Add the checksum and the STOP
      code128$ = code128$ & CHR$(checksum&) & CHR$(211)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caractËres ‡ partir de i% sont numÈriques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(mId$(chaine$, i% + mini%, 1)) < 48 Or Asc(mId$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function

Private Sub Dcbiteem_Change()
Dcbiteem_Click (0)
End Sub
Public Sub printCodes(m_PrintTarget As PrintTarget)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Dim StrFileName As String
StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ScreanBarCode.rpt"
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'MySQL = " select * from TblPrintBarCode"
    MySQL = "SELECT     dbo.TblPrintBarCode.ProductionDate, dbo.TblItems.ItemComment, dbo.TblItems.TotalCalories, dbo.TblItems.shortName, dbo.TblItems.PrintedName, "
   MySQL = MySQL & "                     dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name,"
  MySQL = MySQL & "                      dbo.TblPrintBarCode.Color, dbo.TblPrintBarCode.[size], dbo.TblPrintBarCode.class, dbo.TblPrintBarCode.CodeAnalisys, dbo.TblPrintBarCode.ExpiryDate,"
MySQL = MySQL & "                        dbo.TblPrintBarCode.lotNo , dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.Vat, dbo.TblPrintBarCode.Total, dbo.TblPrintBarCode.code128"
MySQL = MySQL & "  FROM         dbo.TblPrintBarCode LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems ON dbo.TblPrintBarCode.Item_ID = dbo.TblItems.ItemID"
                      
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
  
    Else
       
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        
    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, StrFileName, , MySQL

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub
Private Sub Dcbiteem_Click(Area As Integer)
 Me.TxtCodeAother.Text = GetItemCode(val(Me.Dcbiteem.BoundText))
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg
If .ColKey(Col) <> "QtyPrint" And .ColKey(Col) <> "Print" Then
Cancel = True
Else
.ComboList = ""
End If
End With
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = "     ÿ»«⁄… «·»«—ﬂÊœ  "
    ScreenNameEnglish = " Barcode Printing"

    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
FillMylist
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    With Fg
        Set .WallPaper = GrdBack.Picture
     
    End With
        Dim Dcombo As New ClsDataCombos
        Dcombo.GetItemsNames Me.Dcbiteem
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
         
    End If
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblBarcodePrinting  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
End Sub





Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Me.ListStoreSelected.Clear
    Me.ListActivitySelected.Clear
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtSerial.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
    dbRecordDate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    txtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    TxtQtyPrint.Text = IIf(IsNull(rs("QtyPrint").value), "", rs("QtyPrint").value)
    Me.Dcbiteem.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
    
    StrSQL = " SELECT     dbo.TblBarcodePrintingDet.BrcodID, dbo.TblBarcodePrintingDet.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, "
    StrSQL = StrSQL & "                   dbo.TblBarcodePrintingDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblBarcodePrintingDet.StoreID, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & "                  dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Groups.Fullcode AS GroupFullcode, dbo.TblBarcodePrintingDet.ExpiryDate, dbo.TblBarcodePrintingDet.Qty,"
    StrSQL = StrSQL & "                  dbo.TblBarcodePrintingDet.QtyPrint, dbo.TblBarcodePrintingDet.Price, dbo.TblBarcodePrintingDet.LotNO, dbo.TblBarcodePrintingDet.UnitID, dbo.TblUnites.UnitName,"
    StrSQL = StrSQL & "                  dbo.TblUnites.UnitNamee, dbo.TblBarcodePrintingDet.SortedID, dbo.TblItemsclasses.SizeName AS SortName, dbo.TblItemsclasses.SizeNameE AS SortNameE,"
    StrSQL = StrSQL & "                  dbo.TblBarcodePrintingDet.ColorID, dbo.TblItemsColors.ColorName, dbo.TblBarcodePrintingDet.SizeID, dbo.TblItemsSizes.SizeName, dbo.TblBarcodePrintingDet.ID,"
    StrSQL = StrSQL & "                  dbo.TblBarcodePrintingDet.TransType , dbo.TblBarcodePrintingDet.ItemSerial, dbo.TblItems.shortName, dbo.TblItems.barCodeNO ,dbo.TblItems.PartNo"
    StrSQL = StrSQL & "     FROM         dbo.TblBarcodePrintingDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.TblBarcodePrintingDet.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.TblBarcodePrintingDet.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses ON dbo.TblBarcodePrintingDet.SortedID = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.TblBarcodePrintingDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.TblBarcodePrintingDet.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Groups ON dbo.TblBarcodePrintingDet.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.TblBarcodePrintingDet.ItemID = dbo.TblItems.ItemID"
    StrSQL = StrSQL & "    Where (dbo.TblBarcodePrintingDet.BrcodID = " & val(TxtSerial.Text) & ") And (dbo.TblBarcodePrintingDet.TransType = 0)"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Fg
    
            .Rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .Rows - 1
               .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(RsDev("PartNo").value), "", (RsDev("PartNo").value))
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("ItemID").value), 0, (RsDev("ItemID").value))
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(RsDev("GroupID").value), 0, (RsDev("GroupID").value))
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(RsDev("StoreID").value), 0, (RsDev("StoreID").value))
                .TextMatrix(i, .ColIndex("ExpiryDate")) = IIf(IsNull(RsDev("ExpiryDate").value), "", (RsDev("ExpiryDate").value))
                .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDev("Qty").value), 0, (RsDev("Qty").value))
                .TextMatrix(i, .ColIndex("QtyPrint")) = IIf(IsNull(RsDev("QtyPrint").value), 0, (RsDev("QtyPrint").value))
                .TextMatrix(i, .ColIndex("LotNO")) = IIf(IsNull(RsDev("LotNO").value), "", (RsDev("LotNO").value))
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RsDev("UnitID").value), 0, (RsDev("UnitID").value))
                .TextMatrix(i, .ColIndex("SortedID")) = IIf(IsNull(RsDev("SortedID").value), 0, (RsDev("SortedID").value))
                .TextMatrix(i, .ColIndex("ColorID")) = IIf(IsNull(RsDev("ColorID").value), 0, (RsDev("ColorID").value))
                .TextMatrix(i, .ColIndex("SizeID")) = IIf(IsNull(RsDev("SizeID").value), 0, (RsDev("SizeID").value))
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, (RsDev("Price").value))
                .TextMatrix(i, .ColIndex("ShortName")) = IIf(IsNull(RsDev("ShortName").value), "", (RsDev("ShortName").value))
                .TextMatrix(i, .ColIndex("barCodeNO")) = IIf(IsNull(RsDev("barCodeNO").value), "", (RsDev("barCodeNO").value))
                .TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(RsDev("ItemSerial").value), "", (RsDev("ItemSerial").value))
                .TextMatrix(i, .ColIndex("ColorName")) = IIf(IsNull(RsDev("ColorName").value), "", (RsDev("ColorName").value))
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(RsDev("Fullcode").value), "", (RsDev("Fullcode").value))
                
                .TextMatrix(i, .ColIndex("ItemSize")) = IIf(IsNull(RsDev("SizeName").value), "", (RsDev("SizeName").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("ClassName")) = IIf(IsNull(RsDev("SortName").value), "", (RsDev("SortName").value))
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", (RsDev("UnitName").value))
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreName").value), "", (RsDev("StoreName").value))
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                Else
                .TextMatrix(i, .ColIndex("ClassName")) = IIf(IsNull(RsDev("SortNameE").value), "", (RsDev("SortNameE").value))
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitNamee").value), "", (RsDev("UnitNamee").value))
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsDev("StoreNamee").value), "", (RsDev("StoreNamee").value))
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupNamee").value), "", RsDev("GroupNamee").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)
                End If
                RsDev.MoveNext
            Next i
        End With
    End If
  StrSQL = "  SELECT     dbo.TblBarcodePrintingDet.ItemID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblBarcodePrintingDet.BrcodID, "
  StrSQL = StrSQL & "                     dbo.TblBarcodePrintingDet.TransType"
  StrSQL = StrSQL & "  FROM         dbo.TblBarcodePrintingDet INNER JOIN"
  StrSQL = StrSQL & "                    dbo.Groups ON dbo.TblBarcodePrintingDet.ItemID = dbo.Groups.GroupID"
  StrSQL = StrSQL & "    Where (dbo.TblBarcodePrintingDet.BrcodID = " & val(TxtSerial.Text) & ") And (dbo.TblBarcodePrintingDet.TransType = 1)"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsDev.RecordCount > 0 Then
     RsDev.MoveFirst
     End If
   For i = 0 To RsDev.RecordCount - 1
   If SystemOptions.UserInterface = ArabicInterface Then
   Me.ListActivitySelected.AddItem IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
   Else
   Me.ListActivitySelected.AddItem IIf(IsNull(RsDev("GroupNamee").value), "", RsDev("GroupNamee").value)
   End If
   Me.ListActivitySelected.ItemData(i) = IIf(IsNull(RsDev("ItemID").value), 0, RsDev("ItemID").value)
   RsDev.MoveNext
  Next i
  '''/////////////
    StrSQL = " SELECT     dbo.TblBarcodePrintingDet.BrcodID, dbo.TblBarcodePrintingDet.TransType, dbo.TblBarcodePrintingDet.ItemID, dbo.TblStore.StoreName, "
  StrSQL = StrSQL & "                     dbo.TblStore.storenamee"
  StrSQL = StrSQL & "       FROM         dbo.TblBarcodePrintingDet LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblStore ON dbo.TblBarcodePrintingDet.ItemID = dbo.TblStore.StoreID"
  StrSQL = StrSQL & "    Where (dbo.TblBarcodePrintingDet.BrcodID = " & val(TxtSerial.Text) & ") And (dbo.TblBarcodePrintingDet.TransType = 2)"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsDev.RecordCount > 0 Then
     RsDev.MoveFirst
     End If
   For i = 0 To RsDev.RecordCount - 1
   If SystemOptions.UserInterface = ArabicInterface Then
   Me.ListStoreSelected.AddItem IIf(IsNull(RsDev("StoreName").value), "", RsDev("StoreName").value)
   Else
   Me.ListStoreSelected.AddItem IIf(IsNull(RsDev("storenamee").value), "", RsDev("storenamee").value)
   End If
   Me.ListStoreSelected.ItemData(i) = IIf(IsNull(RsDev("ItemID").value), 0, RsDev("ItemID").value)
   RsDev.MoveNext
  Next i
    
    Exit Sub
ErrTrap:
End Sub
 Function FillMylist()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Set rs2 = New ADODB.Recordset
    sql = " SELECT     GroupID, GroupName, GroupNamee, ParentID"
    sql = sql & "          From dbo.Groups"
    sql = sql & "        Where (Not (ParentID Is Null))"
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Me.ListAllActivity.Clear
    Me.ListActivitySelected.Clear
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllActivity.AddItem IIf(IsNull(rs2("GroupName").value), "", rs2("GroupName").value)
            Else
                ListAllActivity.AddItem IIf(IsNull(rs2("GroupNamee").value), "", rs2("GroupNamee").value)
            End If
            ListAllActivity.ItemData(ListAllActivity.NewIndex) = IIf(IsNull(rs2("GroupID").value), 0, rs2("GroupID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
    '''//////////
       Set rs2 = New ADODB.Recordset
    sql = " SELECT     StoreID, StoreName, StoreNamee"
    sql = sql & "    From dbo.TblStore"
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Me.ListAllStore.Clear
    Me.ListStoreSelected.Clear
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllStore.AddItem IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
            Else
                ListAllStore.AddItem IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
            End If
            ListAllStore.ItemData(ListAllStore.NewIndex) = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
End Function

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
FillGrid
End If
End Sub
Sub FillGrid()
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim StrGroup As String
Dim StrStore As String
Dim UnitID As Double
Dim UnitName As String
Dim Price As Double
Dim i As Integer
StrGroup = "0,"
StrStore = "0,"
For i = 0 To Me.ListActivitySelected.ListCount - 1
StrGroup = StrGroup & Me.ListActivitySelected.ItemData(i)
If i <> Me.ListActivitySelected.ListCount - 1 Then
StrGroup = StrGroup & ","
End If
Next i
For i = 0 To Me.ListStoreSelected.ListCount - 1
StrStore = StrStore & Me.ListStoreSelected.ItemData(i)
If i <> Me.ListStoreSelected.ListCount - 1 Then
StrStore = StrStore & ","
End If
Next i
   getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
   StrSQL = " SELECT     dbo.Transaction_Details.ItemSerial, SUM(dbo.Transaction_Details.ShowQty * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.TblStore.StoreName, "
   StrSQL = StrSQL + "                   dbo.TblUnites.UnitName, dbo.TblItemsColors.ColorName, dbo.TblStore.StoreID, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitID, dbo.TblItemsColors.ColorID,"
   StrSQL = StrSQL + "                   dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.TblItems.ItemID,"
   StrSQL = StrSQL + "                   dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblStore.StoreNamee, dbo.Transaction_Details.ClassId,"
   StrSQL = StrSQL + "                   dbo.TblItemsclasses.SizeId AS SortedID, dbo.TblItemsclasses.SizeName AS SortSizeName, dbo.TblItemsclasses.SizeNameE AS SortSizeNameE,"
   StrSQL = StrSQL + "                   dbo.Transaction_Details.itemsize , dbo.TblItemsSizes.sizename, dbo.TblItemsSizes.sizeid"
   StrSQL = StrSQL + "   , dbo.TblItems.Fullcode ,dbo.TblItems.barCodeNO,dbo.TblItems.shortName ,dbo.TblItems.PartNo"
   StrSQL = StrSQL + "     FROM         dbo.Transactions INNER JOIN"
   StrSQL = StrSQL + "                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
   StrSQL = StrSQL + "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
   StrSQL = StrSQL + "                   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
   StrSQL = StrSQL + "                   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
   StrSQL = StrSQL + "                   dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
   StrSQL = StrSQL + "                   dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
   StrSQL = StrSQL + "                   dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
   StrSQL = StrSQL + "                   dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   StrSQL = StrSQL + "                   dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
   StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
   StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
  If val(Dcbiteem.BoundText) <> 0 Then
    StrSQL = StrSQL + " and  dbo.TblItems.ItemID =" & val(Dcbiteem.BoundText) & ""
  End If

  If StrStore <> "0," Then
   StrSQL = StrSQL + " and dbo.TblStore.StoreId in(" & StrStore & ")"
  End If
If StrGroup <> "0," Then
   StrSQL = StrSQL + " and   dbo.TblItems.GroupID in(" & StrGroup & ")"
  End If
   StrSQL = StrSQL & " GROUP BY dbo.TblStore.StoreName, dbo.TblUnites.UnitName, dbo.TblItemsColors.ColorName, dbo.Transaction_Details.ItemSerial, dbo.TblStore.StoreID, "
   StrSQL = StrSQL & "                    dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitID, dbo.TblItemsColors.ColorID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
   StrSQL = StrSQL & "                    dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
   StrSQL = StrSQL & "                    dbo.TblStore.StoreNamee, dbo.Transaction_Details.ClassId, dbo.TblItemsclasses.SizeId, dbo.TblItemsclasses.SizeName, dbo.TblItemsclasses.SizeNameE,"
   StrSQL = StrSQL & "                    dbo.Transaction_Details.itemsize , dbo.TblItemsSizes.sizename, dbo.TblItemsSizes.sizeid"
   StrSQL = StrSQL + "   , dbo.TblItems.Fullcode ,dbo.TblItems.barCodeNO,dbo.TblItems.shortName ,dbo.TblItems.PartNo"
   StrSQL = StrSQL & "    Having (SUM(dbo.Transaction_Details.ShowQty * dbo.TransactionTypes.StockEffect) <> 0)"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Fg
  .Clear flexClearScrollable, flexClearEverything
 .Rows = 2
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount - 1
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("NumIndex")) = i
.TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(Rs3("PartNo").value), "", Rs3("PartNo").value)
.TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(Rs3("ItemSerial").value), "", Rs3("ItemSerial").value)
.TextMatrix(i, .ColIndex("ShortName")) = IIf(IsNull(Rs3("shortName").value), "", Rs3("shortName").value)
.TextMatrix(i, .ColIndex("barCodeNO")) = IIf(IsNull(Rs3("barCodeNO").value), "", Rs3("barCodeNO").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs3("ItemID").value), "", Rs3("ItemID").value)
  GetUnitInformation val(.TextMatrix(i, .ColIndex("ItemID"))), UnitID, UnitName, Price
  .TextMatrix(i, .ColIndex("UnitName")) = UnitName
.TextMatrix(i, .ColIndex("UnitID")) = UnitID
.TextMatrix(i, .ColIndex("Price")) = Price

.TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(Rs3("SUMQTY").value), "", Rs3("SUMQTY").value)
.TextMatrix(i, .ColIndex("LotNO")) = IIf(IsNull(Rs3("LotNO").value), "", Rs3("LotNO").value)
.TextMatrix(i, .ColIndex("ExpiryDate")) = IIf(IsNull(Rs3("ExpiryDate").value), "", Rs3("ExpiryDate").value)
.TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs3("GroupID").value), "", Rs3("GroupID").value)
.TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(Rs3("StoreID").value), "", Rs3("StoreID").value)
.TextMatrix(i, .ColIndex("SizeID")) = IIf(IsNull(Rs3("SizeId").value), "", Rs3("SizeId").value)

.TextMatrix(i, .ColIndex("SortedID")) = IIf(IsNull(Rs3("SortedID").value), "", Rs3("SortedID").value)
.TextMatrix(i, .ColIndex("ColorID")) = IIf(IsNull(Rs3("ColorID").value), "", Rs3("ColorID").value)
.TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(Rs3("SUMQTY").value), "", Rs3("SUMQTY").value)

If chkhalf.value = vbChecked Then
.TextMatrix(i, .ColIndex("Quantity")) = val(.TextMatrix(i, .ColIndex("Quantity"))) / 2
End If
.TextMatrix(i, .ColIndex("ItemSize")) = IIf(IsNull(Rs3("SizeName").value), "", Rs3("SizeName").value)
.TextMatrix(i, .ColIndex("ColorName")) = IIf(IsNull(Rs3("ColorName").value), "", Rs3("ColorName").value)
.TextMatrix(i, .ColIndex("QtyPrint")) = IIf(val(TxtQtyPrint.Text) = 0, .TextMatrix(i, .ColIndex("Quantity")), val(TxtQtyPrint.Text))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs3("StoreName").value), "", Rs3("StoreName").value)
.TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3("ItemName").value), "", Rs3("ItemName").value)
.TextMatrix(i, .ColIndex("ClassName")) = IIf(IsNull(Rs3("SortSizeName").value), "", Rs3("SortSizeName").value)
Else
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs3("StoreNamee").value), "", Rs3("StoreNamee").value)
.TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3("ItemNamee").value), "", Rs3("ItemNamee").value)
.TextMatrix(i, .ColIndex("ClassName")) = IIf(IsNull(Rs3("SortSizeNameE").value), "", Rs3("SortSizeNameE").value)
End If
Rs3.MoveNext
Next i
End If
End With
End Sub



Sub GetUnitInformation(Optional ItemID As Double, Optional ByRef UnitID As Double = 0, Optional ByRef UnitName As String = "", Optional ByRef UnitSalesPrice As Double = 0)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT     dbo.TblItemsUnits.ItemID, dbo.TblItemsUnits.UnitFactor, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitSalesPrice, dbo.TblUnites.UnitName, "
sql = sql & "                      dbo.TblUnites.UnitNamee"
sql = sql & " FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
sql = sql & " WHERE     (dbo.TblItemsUnits.UnitFactor = " & GetMaxUnitFactor(ItemID) & ") AND (dbo.TblItemsUnits.ItemID = " & ItemID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
UnitID = IIf(IsNull(Rs3("UnitID").value), 0, Rs3("UnitID").value)
UnitSalesPrice = IIf(IsNull(Rs3("UnitSalesPrice").value), 0, Rs3("UnitSalesPrice").value)
If SystemOptions.UserInterface = ArabicInterface Then
UnitName = IIf(IsNull(Rs3("UnitName").value), "", Rs3("UnitName").value)
Else
UnitName = IIf(IsNull(Rs3("UnitNamee").value), "", Rs3("UnitNamee").value)
End If
Else
End If
End Sub
Private Sub Label1_Click()
If Me.TxtModFlg.Text <> "R" Then
If ListActivitySelected.ListIndex > -1 Then
ListActivitySelected.RemoveItem (ListActivitySelected.ListIndex)
End If
End If
End Sub

Private Sub Label2_Click()
If Me.TxtModFlg.Text <> "R" Then
If ListStoreSelected.ListIndex > -1 Then
ListStoreSelected.RemoveItem (ListStoreSelected.ListIndex)
End If
End If
End Sub

Private Sub Label3_Click()
If Me.TxtModFlg.Text <> "R" Then
ListStoreSelected.Clear
End If
End Sub

Private Sub Label4_Click()
If Me.TxtModFlg.Text <> "R" Then
    Dim i As Integer
    For i = 0 To Me.ListAllStore.ListCount - 1
        Me.ListStoreSelected.AddItem ListAllStore.List(i)
        ListStoreSelected.ItemData(i) = ListAllStore.ItemData(i)
    Next i
   End If
End Sub

Private Sub Label6_Click()
If Me.TxtModFlg.Text <> "R" Then
ListActivitySelected.Clear
End If
End Sub
Private Sub Label7_Click()
If Me.TxtModFlg.Text <> "R" Then
    Dim i As Integer
    For i = 0 To Me.ListAllActivity.ListCount - 1
        Me.ListActivitySelected.AddItem ListAllActivity.List(i)
        ListActivitySelected.ItemData(i) = ListAllActivity.ItemData(i)
    Next i
   End If
End Sub

Private Sub Label8_Click()
If Me.TxtModFlg.Text <> "R" Then
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.ListAllActivity.ListIndex > -1 Then
    Me.ListActivitySelected.AddItem ListAllActivity.List(ListAllActivity.ListIndex)
    ListActivitySelected.ItemData(ListActivitySelected.NewIndex) = ListAllActivity.ItemData(ListAllActivity.ListIndex)
End If
End If
End Sub

Private Sub Label9_Click()
If Me.TxtModFlg.Text <> "R" Then
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.ListAllStore.ListIndex > -1 Then
    Me.ListStoreSelected.AddItem ListAllStore.List(ListAllStore.ListIndex)
    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListAllStore.ItemData(ListAllStore.ListIndex)
End If
End If
End Sub

Private Sub TxtCodeAother_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtCodeAother.Text = "" Then
            Me.Dcbiteem.BoundText = ""
        Else
            Me.Dcbiteem.BoundText = GetItemID(Trim$(Me.TxtCodeAother.Text))
        End If
    End If
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        
        CmdRemove1.Enabled = True
        
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
       
        CmdRemove1.Enabled = True
        'Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
      '  Ele(1).Enabled = False
        CmdRemove1.Enabled = False
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
