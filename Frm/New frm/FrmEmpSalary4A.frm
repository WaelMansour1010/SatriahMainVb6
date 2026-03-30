VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalary4A 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ‰ﬁ· «·„⁄œ«  Ê «·«·«   »Ì‰ «·„‘«—Ì⁄"
   ClientHeight    =   7560
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15885
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
   Icon            =   "FrmEmpSalary4A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   15885
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
      _GridInfo       =   $"FrmEmpSalary4A.frx":038A
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
               Height          =   885
               Index           =   5
               Left            =   0
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   0
               Width           =   15795
               _cx             =   27861
               _cy             =   1561
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
               Picture         =   "FrmEmpSalary4A.frx":0410
               Caption         =   "  ‰ﬁ· «·„⁄œ«  Ê «·«·«   »Ì‰ «·„‘«—Ì⁄  "
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
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   53
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   48
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
                  ButtonImage     =   "FrmEmpSalary4A.frx":10EA
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
                  TabIndex        =   49
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
                  ButtonImage     =   "FrmEmpSalary4A.frx":1484
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
                  TabIndex        =   50
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
                  ButtonImage     =   "FrmEmpSalary4A.frx":181E
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
                  TabIndex        =   51
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
                  ButtonImage     =   "FrmEmpSalary4A.frx":1BB8
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
               Left            =   120
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   15705
               _cx             =   27702
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
               Begin MSDataListLib.DataCombo DCoperationType 
                  Height          =   315
                  Left            =   12240
                  TabIndex        =   52
                  Top             =   1680
                  Width           =   2160
                  _ExtentX        =   3810
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
               Begin VB.Frame Frame2 
                  Caption         =   "«·Ï"
                  Height          =   1575
                  Left            =   0
                  TabIndex        =   36
                  Top             =   840
                  Width           =   5595
                  Begin MSDataListLib.DataCombo dcproject1 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   37
                     Top             =   240
                     Width           =   4485
                     _ExtentX        =   7911
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
                  Begin MSDataListLib.DataCombo Dcterm1 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   38
                     Top             =   600
                     Width           =   4485
                     _ExtentX        =   7911
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
                  Begin MSDataListLib.DataCombo dcopr1 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   39
                     Top             =   960
                     Width           =   4485
                     _ExtentX        =   7911
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
                     Caption         =   " «·„‘—Ê⁄"
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
                     Index           =   9
                     Left            =   4605
                     TabIndex        =   42
                     Top             =   240
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·»‰œ"
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
                     Index           =   6
                     Left            =   4485
                     TabIndex        =   41
                     Top             =   600
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄„·Ì…"
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
                     Index           =   3
                     Left            =   4560
                     TabIndex        =   40
                     Top             =   960
                     Width           =   750
                  End
               End
               Begin VB.Frame Frame1 
                  Caption         =   "„‰ "
                  Height          =   1575
                  Left            =   5580
                  TabIndex        =   29
                  Top             =   840
                  Width           =   6435
                  Begin ALLButtonS.ALLButton ALLButton1 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   43
                     Top             =   240
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   1
                     FOCUSR          =   -1  'True
                     BCOL            =   15790320
                     BCOLO           =   15790320
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmEmpSalary4A.frx":1F52
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin MSDataListLib.DataCombo dcproject 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   30
                     Top             =   240
                     Width           =   4485
                     _ExtentX        =   7911
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
                     Left            =   960
                     TabIndex        =   32
                     Top             =   600
                     Width           =   4485
                     _ExtentX        =   7911
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
                  Begin MSDataListLib.DataCombo dcopr 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   34
                     Top             =   960
                     Width           =   4485
                     _ExtentX        =   7911
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
                  Begin ALLButtonS.ALLButton ALLButton2 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   44
                     Top             =   600
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   1
                     FOCUSR          =   -1  'True
                     BCOL            =   15790320
                     BCOLO           =   15790320
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmEmpSalary4A.frx":1F6E
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton ALLButton3 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   45
                     Top             =   960
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     BTYPE           =   3
                     TX              =   "⁄—÷"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   1
                     FOCUSR          =   -1  'True
                     BCOL            =   15790320
                     BCOLO           =   15790320
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmEmpSalary4A.frx":1F8A
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄„·Ì…"
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
                     Index           =   4
                     Left            =   5400
                     TabIndex        =   35
                     Top             =   960
                     Width           =   750
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·»‰œ"
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
                     Index           =   0
                     Left            =   5325
                     TabIndex        =   33
                     Top             =   600
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «·„‘—Ê⁄"
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
                     Left            =   5445
                     TabIndex        =   31
                     Top             =   240
                     Width           =   720
                  End
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6435
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   525
               End
               Begin VB.TextBox xptxtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12195
                  Locked          =   -1  'True
                  TabIndex        =   26
                  Top             =   960
                  Width           =   2235
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«— ﬂ· «·„⁄œ« "
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
                  Left            =   13170
                  TabIndex        =   18
                  Top             =   2160
                  Width           =   2385
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   -4050
                  TabIndex        =   9
                  Top             =   9360
                  Width           =   2235
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3075
                  Left            =   135
                  TabIndex        =   6
                  Top             =   2490
                  Width           =   15465
                  _cx             =   27279
                  _cy             =   5424
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
                  Rows            =   2
                  Cols            =   32
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmEmpSalary4A.frx":1FA6
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
               Begin MSComCtl2.DTPicker XPDtbTrans 
                  Height          =   315
                  Left            =   12195
                  TabIndex        =   11
                  Top             =   1320
                  Width           =   2250
                  _ExtentX        =   3969
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   100532225
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã—«¡ «·Õ«·Ì"
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
                  Index           =   2
                  Left            =   14355
                  TabIndex        =   28
                  Top             =   1680
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·⁄„·Ì…"
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
                  TabIndex        =   10
                  Top             =   1320
                  Width           =   1230
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·⁄„·Ì…"
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
                  Left            =   13680
                  TabIndex        =   8
                  Top             =   960
                  Width           =   1845
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   375
                  Left            =   14220
                  TabIndex        =   7
                  Top             =   960
                  Width           =   900
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
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   12
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
            Left            =   16440
            TabIndex        =   13
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
            ButtonImage     =   "FrmEmpSalary4A.frx":2490
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   15885
            TabIndex        =   14
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
            ButtonImage     =   "FrmEmpSalary4A.frx":282A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   17085
            TabIndex        =   15
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
            ButtonImage     =   "FrmEmpSalary4A.frx":2BC4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   10140
            TabIndex        =   19
            Top             =   120
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
            Left            =   9240
            TabIndex        =   20
            Top             =   150
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
            Left            =   8430
            TabIndex        =   21
            Top             =   120
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
            Left            =   7275
            TabIndex        =   22
            Top             =   150
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
            Left            =   3840
            TabIndex        =   23
            Top             =   150
            Visible         =   0   'False
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
            Left            =   6120
            TabIndex        =   24
            Top             =   150
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
            Left            =   4950
            TabIndex        =   25
            Top             =   150
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
            Left            =   13680
            TabIndex        =   46
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
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
            MICON           =   "FrmEmpSalary4A.frx":2F5E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
      ButtonImage     =   "FrmEmpSalary4A.frx":2F7A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalary4A"
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
                                  ByVal x As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
Dim current_project As Integer
Dim current_term As String

Public Sub YearMonth()

End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    Sql = "Select * from notes where salary=" & year & Month
 
    rs.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    Sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim I As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
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

        For I = .FixedRows To .Rows - 2

            If .TextMatrix(I, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(I, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(I, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(I, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(I, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(I, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next I

    End With
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Function Create_dev1()
    Dim I As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
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

        For I = .FixedRows To .Rows - 2

            If .TextMatrix(I, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(I, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(I, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(I, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(I, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(I, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next I

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

Private Sub ALLButton1_Click()

    Dim sgl As String

    If dcproject.BoundText = "" Then Exit Sub
    'sgl = "select * from opr_employee_details where opr_type=0 and Project_id=" & Val(dcproject.BoundText)
    sgl = "select * from opr_employee_details where opr_type=0 and Ended=0 and  Project_id=" & val(dcproject.BoundText)

    get_all_employee sgl
End Sub

Private Sub ALLButton2_Click()
    Dim sgl As String

    If Dcterm.BoundText = "" Then Exit Sub
    'sgl = "select * from opr_employee_details where opr_type=0 and term_Fullcode='" & Dcterm.BoundText & "'"
    sgl = "select * from opr_employee_details where opr_type=0 and  Ended=0 and  term_Fullcode='" & Dcterm.BoundText & "'"

    get_all_employee sgl
End Sub

Private Sub CboPayMentType_Change()
 
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
    Dim sgl As String

    If dcopr.BoundText = "" Then Exit Sub
    'sgl = "select * from opr_employee_details where opr_type=0 and opr_Fullcode='" & dcopr.BoundText & "'"
    sgl = "select * from opr_employee_details where opr_type=0 and  Ended=0 and   opr_Fullcode='" & dcopr.BoundText & "'"

    get_all_employee sgl

End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
'        get_all_employee
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

    FillGridWithData

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
 
        '   If Trim(Me.dcproject.BoundText) = "" Then
        '       Msg = "ÌÃ» ≈Œ Ì«— «·„‘—Ê⁄..!!"
        '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '       dcproject.SetFocus
        '       SendKeys "{F4}"
        '       Exit Sub
        '   End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
            Me.xptxtid.Text = CStr(new_id("opr_Employee", "id", "", True))
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete opr_employee_details where pk_id=" & val(Me.xptxtid.Text)
   
    End If
    
    rs("id").value = xptxtid.Text
    rs("Start_date").value = XPDtbTrans.value
   
    'from
    rs("Project_id").value = IIf(Me.dcproject.BoundText = "", Null, Me.dcproject.BoundText)
    rs("opr_type").value = IIf(Me.txtType.Text = "", 1, Me.txtType.Text)
     
    If Me.Dcterm.BoundText <> "" Then
        rs("term_Fullcode").value = IIf(Me.Dcterm.BoundText = "", Null, Me.Dcterm.BoundText)
    End If
     
    If Me.dcopr.BoundText <> "" Then
        rs("opr_Fullcode").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
    End If
     
    'to
    rs("Project_id1").value = IIf(Me.dcproject1.BoundText = "", Null, Me.dcproject1.BoundText)
     
    If Me.Dcterm1.BoundText <> "" Then
        rs("term_Fullcode1").value = IIf(Me.Dcterm1.BoundText = "", Null, Me.Dcterm1.BoundText)
    End If
     
    If Me.dcopr1.BoundText <> "" Then
        rs("opr_Fullcode1").value = IIf(Me.dcopr1.BoundText = "", Null, Me.dcopr1.BoundText)
    End If
     
    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim I As Integer

    If Me.DCoperationType.BoundText <> "" Then

        With Me.Grid
        
            For I = .FixedRows To .Rows - 1
                        
                If .TextMatrix(I, .ColIndex("typeid")) = "" Then
                    .TextMatrix(I, .ColIndex("typeid")) = Me.DCoperationType.BoundText
                    .TextMatrix(I, .ColIndex("type")) = Me.DCoperationType.Text
                                  
                End If
                        
            Next I
              
        End With
    
    End If
    
    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If val(.TextMatrix(I, .ColIndex("typeid"))) = 4 Then
                If .TextMatrix(I, .ColIndex("to_projectid")) = "" Then
                    .TextMatrix(I, .ColIndex("to_projectid")) = IIf(dcproject1.BoundText = "", "", dcproject1.BoundText)
                    .TextMatrix(I, .ColIndex("to_project")) = IIf(dcproject1.Text = "", "", dcproject1.Text)
                End If

                If .TextMatrix(I, .ColIndex("to_termid")) = "" Then
                    .TextMatrix(I, .ColIndex("to_termid")) = IIf(Dcterm1.BoundText = "", "", Dcterm1.BoundText)
                End If

                If .TextMatrix(I, .ColIndex("to_oprid")) = "" Then
                    .TextMatrix(I, .ColIndex("to_oprid")) = IIf(dcopr1.BoundText = "", "", dcopr1.BoundText)
                End If
            
            End If

        Next I
        
    End With
        
    With Me.Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Emp_id")) = "" Then
            .Rows = .Rows - 1
        End If

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Emp_id")) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    
                    Msg = "   ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                Else
                    Msg = "  must Select Employee - error in line:" & Chr(13)
                End If
    
                Msg = Msg + CStr(I) & Chr(13)
                Grid.Row = I
                Grid.Col = Grid.ColIndex("Emp_Name")
                Grid.ShowCell I, Grid.ColIndex("Emp_Name")
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Grid.SetFocus
                Screen.MousePointer = vbDefault
                GoTo ErrTrap
            End If

        Next I

    End With
       
    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("typeid")) <> "" Then
            
                Select Case val(.TextMatrix(I, .ColIndex("typeid")))

                    Case 1 '«‰Â«¡

                        If .TextMatrix(I, .ColIndex("Emp_id")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «‰Â«¡ «·⁄„·Ì… ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "During  Operation Finish you must Select Employee - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 2 '«” »œ«·

                        If .TextMatrix(I, .ColIndex("Emp_id")) = .TextMatrix(I, .ColIndex("Emp_id1")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = "·«Ì„ﬂ‰ «” »œ«· «·„ÊŸ› Ê‰›”… —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "you must select two different employee - error in line:" & Chr(13)
                            End If

                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If

                        If .TextMatrix(I, .ColIndex("Emp_id")) = "" Or .TextMatrix(I, .ColIndex("Emp_id1")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «·«” »œ«· ·«»œ „‰  ÕœÌœ «·„ÊŸ› Ê «·„ÊŸ› «·„” »œ· —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "During Exchange Employee with another , you must select Employee and change with employee - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 3

                        If .TextMatrix(I, .ColIndex("Emp_id")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ⁄·Ìﬁ «·⁄„·Ì… ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "During  pending you must Select Employee - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 4

                        If .TextMatrix(I, .ColIndex("to_project")) = .TextMatrix(I, .ColIndex("current_project")) And .TextMatrix(I, .ColIndex("to_termid")) = .TextMatrix(I, .ColIndex("current_term")) And .TextMatrix(I, .ColIndex("to_oprid")) = .TextMatrix(I, .ColIndex("current_opr")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "«·„ÊŸ› Ì⁄„· »«·›⁄· ›Ì Â–« «·„Êﬁ⁄ ﬁ„ » €Ì— «·„Êﬁ⁄ —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "this Employee Already work in this location Select another location if you want - error in line:" & Chr(13)
                            End If

                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If
                        
                        If .TextMatrix(I, .ColIndex("to_projectid")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ‰ﬁ· „ÊŸ› „‰ „‘—Ê⁄ ·«Œ— ·«»œ „‰ «Œ Ì«— «·„‘—Ê⁄ «·ÃœÌœ ⁄·Ï «·«ﬁ· —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "in case change employee project you must select at least new project name - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("to_project")
                            Grid.ShowCell I, Grid.ColIndex("to_project")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
              
                    Case 5

                        If .TextMatrix(I, .ColIndex("to_project")) = .TextMatrix(I, .ColIndex("current_project")) And .TextMatrix(I, .ColIndex("to_termid")) = .TextMatrix(I, .ColIndex("current_term")) And .TextMatrix(I, .ColIndex("to_oprid")) = .TextMatrix(I, .ColIndex("current_opr")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "«·„ÊŸ› Ì⁄„· »«·›⁄· ›Ì Â–« «·„Êﬁ⁄ ﬁ„ » €Ì— «·„Êﬁ⁄ —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "this Employee Already work in this location Select another location if you want - error in line:" & Chr(13)
                            End If

                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If
                        
                        If .TextMatrix(I, .ColIndex("to_projectid")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ‰ﬁ·  Ê«” »œ«· „ÊŸ› „‰ „‘—Ê⁄ ·«Œ— ·«»œ „‰ «Œ Ì«— «·„‘—Ê⁄ «·ÃœÌœ ⁄·Ï «·«ﬁ· —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "in case change employee project you must select at least new project name - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("to_project")
                            Grid.ShowCell I, Grid.ColIndex("to_project")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
                        
                        If .TextMatrix(I, .ColIndex("Emp_id")) = "" Or .TextMatrix(I, .ColIndex("Emp_id1")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «·‰ﬁ· «·«” »œ«· ·«»œ „‰  ÕœÌœ «·„ÊŸ› Ê «·„ÊŸ› «·„” »œ· —«Ã⁄ «·”ÿ— —ﬁ„ " & Chr(13)
                            Else
                                Msg = "During Exchange Employee with another , you must select Employee and change with employee - error in line:" & Chr(13)
                            End If
    
                            Msg = Msg + CStr(I) & Chr(13)
                            Grid.Row = I
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell I, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
                        
                End Select
            
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    Msg = " ·«»œ „‰  ÕœÌœ  Õ«·… «·⁄„·Ì… ›Ì «·”ÿ— —ﬁ„ " & Chr(13)
                Else
                    Msg = " must Specify Operation Type  in row no: " & Chr(13)
                End If

                Msg = Msg + CStr(I) & Chr(13)
                Grid.Row = I
                Grid.Col = Grid.ColIndex("typeid")
                Grid.ShowCell I, Grid.ColIndex("type")
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Grid.SetFocus
                Screen.MousePointer = vbDefault
                GoTo ErrTrap
          
            End If

        Next I
        
    End With

    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("PK")) <> "" Then
         
                RsDev.AddNew
                RsDev("pk_id").value = Me.xptxtid.Text
                RsDev("Emp_id").value = .TextMatrix(I, .ColIndex("Emp_id"))
                RsDev("emp_code").value = .TextMatrix(I, .ColIndex("Emp_Code"))
                RsDev("emp_name").value = .TextMatrix(I, .ColIndex("Emp_Name"))
                RsDev("JobTypeName").value = .TextMatrix(I, .ColIndex("JobTypeName"))
                RsDev("opr_type").value = 1
                RsDev("start_date").value = Date
                
                RsDev("opreration_type").value = IIf(.TextMatrix(I, .ColIndex("typeid")) = "", Null, .TextMatrix(I, .ColIndex("typeid")))
                
                RsDev("Project_id").value = IIf(.TextMatrix(I, .ColIndex("current_project")) = "", 0, .TextMatrix(I, .ColIndex("current_project")))
                RsDev("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_term")) = "", Null, .TextMatrix(I, .ColIndex("current_term")))
                RsDev("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", Null, .TextMatrix(I, .ColIndex("current_opr")))
                
                RsDev("to_project").value = IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", Null, .TextMatrix(I, .ColIndex("to_projectid")))
                RsDev("to_project_name").value = IIf(.TextMatrix(I, .ColIndex("to_project")) = "", Null, .TextMatrix(I, .ColIndex("to_project")))
                
                RsDev("to_term").value = IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", Null, .TextMatrix(I, .ColIndex("to_termid")))
                RsDev("to_opr").value = IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", Null, .TextMatrix(I, .ColIndex("to_oprid")))
              
                RsDev("person_id").value = IIf(.TextMatrix(I, .ColIndex("Emp_id1")) = "", Null, .TextMatrix(I, .ColIndex("Emp_id1")))
                RsDev("person_code").value = IIf(.TextMatrix(I, .ColIndex("emp_code1")) = "", Null, .TextMatrix(I, .ColIndex("emp_code1")))
                RsDev("person_name").value = IIf(.TextMatrix(I, .ColIndex("Emp_Name1")) = "", Null, .TextMatrix(I, .ColIndex("Emp_Name1")))
                RsDev("person_joptypename").value = IIf(.TextMatrix(I, .ColIndex("JobTypeName1")) = "", Null, .TextMatrix(I, .ColIndex("JobTypeName1")))
                RsDev.update

                Select Case val(.TextMatrix(I, .ColIndex("typeid")))

                    Case 1 ' «‰Â«¡

                        ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’
                        If .TextMatrix(I, .ColIndex("pk")) <> "" Then
                            Dim Sql As String
                            Sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(I, .ColIndex("pk")))
                            Cn.Execute Sql, , adExecuteNoRecords
                        End If
             
                        'Õ–› «”„ «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›Ì‰
                        If .TextMatrix(I, .ColIndex("Emp_id")) <> "" Then
            
                            Sql = "update  TblEmployee  set project_id=0  ,term_fullcode=Null,opr_fullcode=Null where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id")))
                            Cn.Execute Sql, , adExecuteNoRecords
                        End If
             
                        ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·Ï «‰Â«¡ Ê ”ÃÌ·  «—ÌŒ «·«‰Â«¡
                        If .TextMatrix(I, .ColIndex("current_opr")) <> "" Then
                            Sql = "update  terms_operations  set ended=1, end_date='" & SQLDate(Date) & "' where fullcode ='" & .TextMatrix(I, .ColIndex("current_opr")) & "'"
                            Cn.Execute Sql, , adExecuteNoRecords
                        End If
         
                    Case 2   '«” »œ«· „ÊŸ› »„ÊŸ› «Œ— ›Ì ‰›” «·„‘—Ê⁄
         
                        If .TextMatrix(I, .ColIndex("Emp_id")) <> "" And .TextMatrix(I, .ColIndex("Emp_id1")) <> "" Then
                            'Õ–› »Ì«‰«   «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›
                            Sql = "update  TblEmployee  set project_id=0  ,term_fullcode=Null,opr_fullcode=Null where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id")))
                            Cn.Execute Sql, , adExecuteNoRecords
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄ ··„ÊŸ› «·ÃœÌœ ›Ì „·› «·„ÊŸ›
           
                            Sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(I, .ColIndex("current_project")) = "", 0, .TextMatrix(I, .ColIndex("current_project"))) & "  ,term_fullcode=" & IIf(.TextMatrix(I, .ColIndex("current_term")) = "", "''", "'" & .TextMatrix(I, .ColIndex("current_term")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", "''", "'" & .TextMatrix(I, .ColIndex("current_opr")) & "'") & " where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id1")))
                            Cn.Execute Sql, , adExecuteNoRecords
          
                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ› «·«”«”Ì
                            If .TextMatrix(I, .ColIndex("pk")) <> "" Then
              
                                Sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(I, .ColIndex("pk")))
                                Cn.Execute Sql, , adExecuteNoRecords
                            End If
             
                            ' ”ÃÌ·  Œ’Ì’ «·„ÊŸ› «·ÃœÌœ ⁄·Ï «·⁄„·Ì…
                            Dim Rs1 As ADODB.Recordset
                            Dim Rs2 As ADODB.Recordset
                            Dim ID As Integer
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(I, .ColIndex("current_project")) = "", 0, .TextMatrix(I, .ColIndex("current_project")))
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_term")) = "", Null, .TextMatrix(I, .ColIndex("current_term")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", Null, .TextMatrix(I, .ColIndex("current_opr")))
                            Rs1.update
                            
                            Set Rs2 = New ADODB.Recordset
        
                            Rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            Rs2.AddNew
                            Rs2("pk_id").value = ID
                            Rs2("Emp_id").value = .TextMatrix(I, .ColIndex("Emp_id1"))
                            Rs2("emp_code").value = .TextMatrix(I, .ColIndex("Emp_Code1"))
                            Rs2("emp_name").value = .TextMatrix(I, .ColIndex("Emp_Name1"))
                            Rs2("JobTypeName").value = .TextMatrix(I, .ColIndex("JobTypeName1"))
                            Rs2("opr_type").value = 0
                            Rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            Rs2("Project_id").value = IIf(.TextMatrix(I, .ColIndex("current_project")) = "", Null, .TextMatrix(I, .ColIndex("current_project")))
                            Rs2("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_term")) = "", Null, .TextMatrix(I, .ColIndex("current_term")))
                            Rs2("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", Null, .TextMatrix(I, .ColIndex("current_opr")))
                            Rs2.update
                        End If
                    
                    Case 3 ' ⁄·Ìﬁ „ÊŸ›

                        If .TextMatrix(I, .ColIndex("Emp_id")) <> "" Then
                            'Õ–› »Ì«‰«   «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›
                            Sql = "update  TblEmployee  set project_id=0  ,term_fullcode='',opr_fullcode='' where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id")))
                            Cn.Execute Sql, , adExecuteNoRecords
                            
                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ›
                            If .TextMatrix(I, .ColIndex("pk")) <> "" Then
              
                                Sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(I, .ColIndex("pk")))
                                Cn.Execute Sql, , adExecuteNoRecords
                            End If

                            ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·Ï  ⁄·Ìﬁ
                            '«Ê÷«⁄ «·⁄„·Ì«  0 ﬁÌœ «· ‰›Ì– Ê 1 «‰ Â  Ê 2 „⁄·ﬁ…
                            If .TextMatrix(I, .ColIndex("current_opr")) <> "" Then
                                Sql = "update  terms_operations  set ended=2  where fullcode ='" & .TextMatrix(I, .ColIndex("current_opr")) & "'"
                                Cn.Execute Sql, , adExecuteNoRecords
                            End If
                        End If

                    Case 4 '  ‰ﬁ· „ÊŸ› „‰ „‘—Ê⁄ «Ê»‰œ «Ê ⁄„·Ì… «·Ï „‘—Ê⁄ «Ê »‰œ «Ê ⁄„·Ì… «Œ—Ï
                    
                        If .TextMatrix(I, .ColIndex("Emp_id")) <> "" Then

                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ›
                            If .TextMatrix(I, .ColIndex("pk")) <> "" Then
              
                                Sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(I, .ColIndex("pk")))
                                Cn.Execute Sql, , adExecuteNoRecords
                            End If
                           
                            ' ”ÃÌ·  Œ’Ì’ «·„ÊŸ› ›Ì «·⁄„·Ì… «·ÃœÌœ…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", 0, .TextMatrix(I, .ColIndex("to_projectid")))
                            
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", Null, .TextMatrix(I, .ColIndex("to_termid")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", Null, .TextMatrix(I, .ColIndex("to_oprid")))
                            Rs1.update
                            
                            Set Rs2 = New ADODB.Recordset
        
                            Rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            Rs2.AddNew
                            Rs2("pk_id").value = ID
                            Rs2("Emp_id").value = .TextMatrix(I, .ColIndex("Emp_id"))
                            Rs2("emp_code").value = .TextMatrix(I, .ColIndex("Emp_Code"))
                            Rs2("emp_name").value = .TextMatrix(I, .ColIndex("Emp_Name"))
                            Rs2("JobTypeName").value = .TextMatrix(I, .ColIndex("JobTypeName"))
                            Rs2("opr_type").value = 0
                            Rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            Rs2("Project_id").value = IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", Null, .TextMatrix(I, .ColIndex("to_projectid")))
                            Rs2("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", Null, .TextMatrix(I, .ColIndex("to_termid")))
                            Rs2("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", Null, .TextMatrix(I, .ColIndex("to_oprid")))
                            Rs2.update
                            
                            ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·ﬁœÌ„… «·Ï «‰Â«¡ Ê ”ÃÌ·  «—ÌŒ «·«‰Â«¡
                            '  If .TextMatrix(I, .ColIndex("current_opr")) <> "" Then
                            '      Sql = "update  terms_operations  set ended=1, end_date='" & SQLDate(Date) & "' where fullcode ='" & .TextMatrix(I, .ColIndex("current_opr")) & "'"
                            '      Cn.Execute Sql, , adExecuteNoRecords
                            '  End If
                        
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄ ··„ÊŸ› «·ÃœÌœ ›Ì „·› «·„ÊŸ›
           
                            Sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", 0, .TextMatrix(I, .ColIndex("to_projectid"))) & "  ,term_fullcode=" & IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", "''", "'" & .TextMatrix(I, .ColIndex("to_termid")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", "''", "'" & .TextMatrix(I, .ColIndex("to_oprid")) & "'") & " where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id")))
                            Cn.Execute Sql, , adExecuteNoRecords
                            
                        End If

                    Case 5 '  ‰ﬁ· Ê «” »œ«· „ÊŸ› „‰ „‘—Ê⁄ «Ê»‰œ «Ê ⁄„·Ì… «·Ï „‘—Ê⁄ «Ê »‰œ «Ê ⁄„·Ì… «Œ—Ï
                    
                        If .TextMatrix(I, .ColIndex("Emp_id")) <> "" Then

                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ› A
                            If .TextMatrix(I, .ColIndex("pk")) <> "" Then
              
                                Sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(I, .ColIndex("pk")))
                                Cn.Execute Sql, , adExecuteNoRecords
                            End If
                           
                            ' ”ÃÌ·  Œ’Ì’ A «·„ÊŸ› ›Ì «·⁄„·Ì… «·ÃœÌœ…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", 0, .TextMatrix(I, .ColIndex("to_projectid")))
                            
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", Null, .TextMatrix(I, .ColIndex("to_termid")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", Null, .TextMatrix(I, .ColIndex("to_oprid")))
                            Rs1.update
                            
                            Set Rs2 = New ADODB.Recordset
        
                            Rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            Rs2.AddNew
                            Rs2("pk_id").value = ID
                            Rs2("Emp_id").value = .TextMatrix(I, .ColIndex("Emp_id"))
                            Rs2("emp_code").value = .TextMatrix(I, .ColIndex("Emp_Code"))
                            Rs2("emp_name").value = .TextMatrix(I, .ColIndex("Emp_Name"))
                            Rs2("JobTypeName").value = .TextMatrix(I, .ColIndex("JobTypeName"))
                            Rs2("opr_type").value = 0
                            Rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            Rs2("Project_id").value = IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", Null, .TextMatrix(I, .ColIndex("to_projectid")))
                            Rs2("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", Null, .TextMatrix(I, .ColIndex("to_termid")))
                            Rs2("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", Null, .TextMatrix(I, .ColIndex("to_oprid")))
                            Rs2.update
                        
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄  «·ÃœÌœ ··„ÊŸ› A ›Ì „·› «·„ÊŸ›
           
                            Sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(I, .ColIndex("to_projectid")) = "", 0, .TextMatrix(I, .ColIndex("to_projectid"))) & "  ,term_fullcode=" & IIf(.TextMatrix(I, .ColIndex("to_termid")) = "", "''", "'" & .TextMatrix(I, .ColIndex("to_termid")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(I, .ColIndex("to_oprid")) = "", "''", "'" & .TextMatrix(I, .ColIndex("to_oprid")) & "'") & " where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id")))
                            Cn.Execute Sql, , adExecuteNoRecords
                            
                            '«·„ÊŸ› b «·„” »œ·
                           
                            ' ”ÃÌ·  Œ’Ì’ B «·„ÊŸ› ›Ì «·⁄„·Ì… «·ﬁœÌ„…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(I, .ColIndex("current_project")) = "", 0, .TextMatrix(I, .ColIndex("current_project")))
                            
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_term")) = "", Null, .TextMatrix(I, .ColIndex("current_term")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", Null, .TextMatrix(I, .ColIndex("current_opr")))
                            Rs1.update
                            
                            Set Rs2 = New ADODB.Recordset
        
                            Rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            Rs2.AddNew
                            Rs2("pk_id").value = ID
                            Rs2("Emp_id").value = .TextMatrix(I, .ColIndex("Emp_id1"))
                            Rs2("emp_code").value = .TextMatrix(I, .ColIndex("Emp_Code1"))
                            Rs2("emp_name").value = .TextMatrix(I, .ColIndex("Emp_Name1"))
                            Rs2("JobTypeName").value = .TextMatrix(I, .ColIndex("JobTypeName1"))
                            Rs2("opr_type").value = 0
                            Rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            Rs2("Project_id").value = IIf(.TextMatrix(I, .ColIndex("current_project")) = "", Null, .TextMatrix(I, .ColIndex("current_project")))
                            Rs2("term_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_term")) = "", Null, .TextMatrix(I, .ColIndex("current_term")))
                            Rs2("opr_Fullcode").value = IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", Null, .TextMatrix(I, .ColIndex("current_opr")))
                            Rs2.update
                        
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄  «·ÃœÌœ ··„ÊŸ› b ›Ì „·› «·„ÊŸ›
           
                            Sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(I, .ColIndex("current_project")) = "", 0, .TextMatrix(I, .ColIndex("current_project"))) & "  ,term_fullcode=" & IIf(.TextMatrix(I, .ColIndex("current_term")) = "", "''", "'" & .TextMatrix(I, .ColIndex("current_term")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(I, .ColIndex("current_opr")) = "", "''", "'" & .TextMatrix(I, .ColIndex("current_opr")) & "'") & " where Emp_id =" & val(.TextMatrix(I, .ColIndex("Emp_id1")))
                            Cn.Execute Sql, , adExecuteNoRecords
                        End If

                End Select
                    
            End If
            
            '
        Next I

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

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

    TxtModFlg.Text = "R"
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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            'Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
 
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me

       
            '  XPDtbTrans.value = Date
       
            '        XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
            Grid.Enabled = True

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            ' Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.searchtype = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
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
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If Rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If Rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If Rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If Rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If Rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If Rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If Rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If Rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If Rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If Rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If Rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If Rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If Rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If Rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If Rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If Rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If Rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If Rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If x = vbNo Then Exit Sub
    
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

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
            
        My_SQL = " select id,Project_name from projects"
        fill_combo dcproject, My_SQL
    End If
        
End Sub

Private Sub dcproject1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
      '  Dim My_SQL As String
            
      '  My_SQL = " select id,Project_name from projects"
      '  fill_combo dcproject1, My_SQL
    
    Dim Dcombos As New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
     Dcombos.GetProjects dcproject1
    End If

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
                 
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT id,name  FROM employee_operations_type "
    Else
        My_SQL = "SELECT id,namee  FROM employee_operations_type "
    End If
                
    fill_combo Me.DCoperationType, My_SQL

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL

    My_SQL = " select  fullcode,des from projects_des"
    fill_combo Dcterm, My_SQL

    My_SQL = " select  fullcode,name from terms_operations"
    fill_combo dcopr, My_SQL

   ' My_SQL = " select id,Project_name from projects"
   ' fill_combo dcproject1, My_SQL

    My_SQL = " select  fullcode,des from projects_des"
    fill_combo Dcterm1, My_SQL

    My_SQL = " select  fullcode,name from terms_operations"
    fill_combo dcopr1, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
     Dcombos.GetProjects dcproject
     Dcombos.GetProjects dcproject1
    Set BKGrndPic = New ClsBackGroundPic

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
    StrSQL = "select * From opr_Employee where opr_type=1 " ' 0 ‘«‘… «·‰Œ’Ì’ 1 ‘«‘… «·‰ﬁ· Ê «·«” »œ«·
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ALLButton1.Caption = "View"
    ALLButton2.Caption = "View"
    ALLButton3.Caption = "View"
Me.Caption = "Equipments Movements Transactions"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    CmdRemove.Caption = "Remove Line"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

     Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(8).Caption = "Date"
    lbl(2).Caption = "Procedure"

    'Ele(3).Caption = "Select Interval"
    lbl(5).Caption = "Project"
    lbl(0).Caption = "Term"
    lbl(4).Caption = "Operation"

    lbl(9).Caption = "Project"
    lbl(6).Caption = "Term"
    lbl(3).Caption = "Operation"
    Frame1.Caption = "From Project"
    Frame2.Caption = "To Project"

    Check1.Caption = "Show All Equipments"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Emp_code")) = "Equip.Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Equip. Name"
        .TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
        .TextMatrix(0, .ColIndex("work_status")) = "work_status"
        .TextMatrix(0, .ColIndex("project_name")) = "project name"
        .TextMatrix(0, .ColIndex("cost_center")) = "cost center"
        .TextMatrix(0, .ColIndex("work_days")) = "work days"
        .TextMatrix(0, .ColIndex("ATTENDANCE")) = "absence"
        .TextMatrix(0, .ColIndex("late")) = "delay"
        .TextMatrix(0, .ColIndex("discount")) = "discount"
        .TextMatrix(0, .ColIndex("net_work_days")) = "net work days"
        .TextMatrix(0, .ColIndex("addition")) = "over time"
        .TextMatrix(0, .ColIndex("remarks")) = "remarks"

        .TextMatrix(0, .ColIndex("type")) = "Procedure"
        .TextMatrix(0, .ColIndex("to_project")) = "to project"
        .TextMatrix(0, .ColIndex("to_termid")) = "to term "
        .TextMatrix(0, .ColIndex("to_oprid")) = "to opr"
        .TextMatrix(0, .ColIndex("Emp_Name1")) = "To Equip"

    End With

End Sub

Public Sub get_all_employee(Optional Sql As String = "")
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim j As Integer
    Dim xx As Integer
    xx = 0
    Dim I As Integer
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
    Grid.Enabled = True
          
    If Sql = "" Then
        Sql = "Select * from emp_all_details "
        xx = 1
    End If
 
    Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For I = 1 To Rs3.RecordCount
                .TextMatrix(I, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                .TextMatrix(I, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                    
                .TextMatrix(I, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                       
                If xx = 1 Then
                    .TextMatrix(I, .ColIndex("PK")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                Else
                    .TextMatrix(I, .ColIndex("PK")) = IIf(IsNull(Rs3.Fields("id").value), "", Rs3.Fields("id").value)
                End If
                       
                .TextMatrix(I, .ColIndex("current_opr")) = IIf(IsNull(Rs3("opr_fullcode").value), "", Rs3("opr_fullcode").value)
                    
                .TextMatrix(I, .ColIndex("current_term")) = IIf(IsNull(Rs3("term_Fullcode").value), "", Rs3("term_Fullcode").value)
                     
                .TextMatrix(I, .ColIndex("current_project")) = IIf(IsNull(Rs3("Project_id").value), 0, Rs3("Project_id").value)
          
                Rs3.MoveNext
            Next I
  
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim I As Integer
    Dim rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
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
    Set Rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For I = 1 To .Rows - 1
        
                .TextMatrix(I, .ColIndex("Ser")) = I
                ',DepartmentID,project_id
            
                .TextMatrix(I, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(I, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(I, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(I, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(I, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(I, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(I, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(I, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(I, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(I, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(I, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(I, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(I, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
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

Public Sub Grid_AfterEdit(ByVal Row As Long, _
                          ByVal Col As Long)
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim I As Integer

    With Me.Grid

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Emp_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
  
            End If

        Next I
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid
 
        .ComboList = ""

        Select Case .ColKey(Col)

            Case "Emp_Code"
                .ComboList = ""

            Case "JobTypeName"
                .ComboList = ""
        
            Case "DepartmentName"
                .ComboList = ""
        
            Case "work_status"
                .ComboList = ""

            Case "work_days"
                .ComboList = ""

            Case "attendance"
                .ComboList = ""

            Case "late"
                .ComboList = ""

            Case "discount"
                .ComboList = ""

            Case "net_work_days"
                .ComboList = ""

            Case "addition"
                .ComboList = ""

            Case "remarks"
                .ComboList = ""

            Case "absence"
                .ComboList = ""
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub Grid_Click()

    With Grid
            
        ' get_Employee_project_information Val(.TextMatrix(.Row, .ColIndex("Emp_id")))
    End With

End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, _
                       Shift As Integer)

    If KeyCode = vbKeyF3 Then
       ' project_search.show
       ' project_search.case_id = 1

    End If

    With Me.Grid

        If KeyCode = 13 Then
            get_Employee_project_information val(.TextMatrix(.Row, .ColIndex("Emp_id")))
        End If

    End With

End Sub

  
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Me.Grid

        Select Case .ColKey(Col)
 
Case "Emp_Name", "Emp_Name1"
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
.Rows = .Rows + 1
 
'            Case "Emp_Name"
        
                'Full Path Display
'                StrSQL = "SELECT *  FROM emp_all_details  where project_id=0 or project_id is null"
'
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                StrComboList = Grid.BuildComboList(rs, "emp_name", "Emp_id")
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'
'                .ComboList = StrComboList
'
        End Select

    End With



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

    'Retrive
    Exit Sub
ErrTrap:
End Sub
