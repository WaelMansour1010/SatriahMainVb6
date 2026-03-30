VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalary4 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁ· «·„ÊŸ›Ì‰ »Ì‰ «·„‘«—Ì⁄ "
   ClientHeight    =   7605
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
   Icon            =   "FrmEmpSalary4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
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
      _GridInfo       =   $"FrmEmpSalary4.frx":038A
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
            Begin VB.ComboBox CboYear 
               Height          =   330
               Left            =   0
               TabIndex        =   56
               Text            =   "Combo1"
               Top             =   0
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.ComboBox CmbMonth 
               Height          =   330
               Left            =   180
               TabIndex        =   55
               Text            =   "Combo1"
               Top             =   5970
               Visible         =   0   'False
               Width           =   795
            End
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
               Picture         =   "FrmEmpSalary4.frx":0410
               Caption         =   "‰ﬁ· «·„ÊŸ›Ì‰ »Ì‰ «·„‘«—Ì⁄ "
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
                  ButtonImage     =   "FrmEmpSalary4.frx":10EA
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
                  ButtonImage     =   "FrmEmpSalary4.frx":1484
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
                  ButtonImage     =   "FrmEmpSalary4.frx":181E
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
                  ButtonImage     =   "FrmEmpSalary4.frx":1BB8
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
                     Style           =   2
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
                     Style           =   2
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
                     Style           =   2
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
                     MICON           =   "FrmEmpSalary4.frx":1F52
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
                     Style           =   2
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
                     Style           =   2
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
                     Style           =   2
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
                     MICON           =   "FrmEmpSalary4.frx":1F6E
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
                     MICON           =   "FrmEmpSalary4.frx":1F8A
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
                  Caption         =   "«ŸÂ«— ﬂ· «·„ÊŸ›Ì‰"
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
                  Cols            =   35
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmEmpSalary4.frx":1FA6
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
                  Format          =   117571585
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
            ButtonImage     =   "FrmEmpSalary4.frx":24FD
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
            ButtonImage     =   "FrmEmpSalary4.frx":2897
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
            ButtonImage     =   "FrmEmpSalary4.frx":2C31
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11280
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
            Left            =   10380
            TabIndex        =   20
            Top             =   120
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
            Left            =   9570
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
            Left            =   8415
            TabIndex        =   22
            Top             =   120
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
            Left            =   4980
            TabIndex        =   23
            Top             =   120
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
            Left            =   7260
            TabIndex        =   24
            Top             =   120
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
            Left            =   6090
            TabIndex        =   25
            Top             =   120
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
            MICON           =   "FrmEmpSalary4.frx":2FCB
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
            Left            =   4080
            TabIndex        =   54
            Top             =   150
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            Width           =   1455
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
      ButtonImage     =   "FrmEmpSalary4.frx":2FE7
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalary4"
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

Dim mToDate As String
Dim NoDay As Long
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
Dim current_project As Integer
Dim current_term As String



Private Sub ChkDetails_Click()
    FillGridWithData
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

Private Sub ALLButton1_Click()

    Dim sgl As String

    If DCPROJECT.BoundText = "" Then Exit Sub
    'sgl = "select * from opr_employee_details where opr_type=0 and Project_id=" & Val(dcproject.BoundText)
    sgl = "select * from opr_employee_details where opr_type=0 and Ended=0 and  Project_id=" & val(DCPROJECT.BoundText)

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
            Me.XPTxtID.Text = CStr(new_id("opr_Employee", "id", "", True))
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete opr_employee_details where pk_id=" & val(Me.XPTxtID.Text)
   
    End If
    
    rs("id").value = XPTxtID.Text
    rs("Start_date").value = XPDtbTrans.value
   
    'from
    rs("Project_id").value = IIf(Me.DCPROJECT.BoundText = "", Null, Me.DCPROJECT.BoundText)
    rs("opr_type").value = 1 ' IIf(Me.txtType.Text = "", 1, Me.txtType.Text)
     
       
  rs("Years").value = val(CboYear.ListIndex)
  rs("Months").value = val(CmbMonth.ListIndex)
  
    If Me.Dcterm.BoundText <> "" Then
        rs("term_Fullcode").value = IIf(Me.Dcterm.BoundText = "", Null, Me.Dcterm.BoundText)
    End If
     
    If Me.dcopr.BoundText <> "" Then
        rs("opr_Fullcode").value = IIf(Me.dcopr.BoundText = "", Null, Me.dcopr.BoundText)
    End If
     
    'to
    rs("Project_id1").value = IIf(Me.DCPROJECT1.BoundText = "", Null, Me.DCPROJECT1.BoundText)
     
    If Me.Dcterm1.BoundText <> "" Then
        rs("term_Fullcode1").value = IIf(Me.Dcterm1.BoundText = "", Null, Me.Dcterm1.BoundText)
    End If
     
    If Me.dcopr1.BoundText <> "" Then
        rs("opr_Fullcode1").value = IIf(Me.dcopr1.BoundText = "", Null, Me.dcopr1.BoundText)
    End If
     
    rs.update
    
    
    Dim s As String
    s = "Delete opr_employee_details Where IDTransfer = " & val(XPTxtID.Text)
    Cn.Execute s
    s = "Delete opr_employee Where IDTransfer = " & val(XPTxtID.Text)
    Cn.Execute s
    
    s = "Delete ProJectMofrdSalar Where pk_id = " & val(Me.XPTxtID.Text)
    Cn.Execute s
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    If Me.DCoperationType.BoundText <> "" Then

        With Me.Grid
        
            For i = .FixedRows To .Rows - 1
                        
                If .TextMatrix(i, .ColIndex("typeid")) = "" Then
                    .TextMatrix(i, .ColIndex("typeid")) = Me.DCoperationType.BoundText
                    .TextMatrix(i, .ColIndex("type")) = Me.DCoperationType.Text
                                  
                End If
                        
            Next i
              
        End With
    
    End If
    
    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("typeid"))) = 4 Then
                If .TextMatrix(i, .ColIndex("to_projectid")) = "" Then
                    .TextMatrix(i, .ColIndex("to_projectid")) = IIf(DCPROJECT1.BoundText = "", "", DCPROJECT1.BoundText)
                    .TextMatrix(i, .ColIndex("to_project")) = IIf(DCPROJECT1.Text = "", "", DCPROJECT1.Text)
                End If

                If .TextMatrix(i, .ColIndex("to_termid")) = "" Then
                    .TextMatrix(i, .ColIndex("to_termid")) = IIf(Dcterm1.BoundText = "", "", Dcterm1.BoundText)
                End If

                If .TextMatrix(i, .ColIndex("to_oprid")) = "" Then
                    .TextMatrix(i, .ColIndex("to_oprid")) = IIf(dcopr1.BoundText = "", "", dcopr1.BoundText)
                End If
            
            End If
                    
        Next i
        
    End With
        
    With Me.Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Emp_id")) = "" Then
            .Rows = .Rows - 1
        End If

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_id")) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    
                    Msg = "   ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                Else
                    Msg = "  must Select Employee - error in line:" & CHR(13)
                End If
    
                Msg = Msg + CStr(i) & CHR(13)
                Grid.Row = i
                Grid.Col = Grid.ColIndex("Emp_Name")
                Grid.ShowCell i, Grid.ColIndex("Emp_Name")
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Grid.SetFocus
                Screen.MousePointer = vbDefault
                GoTo ErrTrap
            End If

        Next i

    End With
       
    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("typeid")) <> "" Then
            
                Select Case val(.TextMatrix(i, .ColIndex("typeid")))

                    Case 1 '«‰Â«¡

                        If .TextMatrix(i, .ColIndex("Emp_id")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «‰Â«¡ «·⁄„·Ì… ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "During  Operation Finish you must Select Employee - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 2 '«” »œ«·

                        If .TextMatrix(i, .ColIndex("Emp_id")) = .TextMatrix(i, .ColIndex("Emp_id1")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = "·«Ì„ﬂ‰ «” »œ«· «·„ÊŸ› Ê‰›”… —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "you must select two different employee - error in line:" & CHR(13)
                            End If

                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If

                        If .TextMatrix(i, .ColIndex("Emp_id")) = "" Or .TextMatrix(i, .ColIndex("Emp_id1")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «·«” »œ«· ·«»œ „‰  ÕœÌœ «·„ÊŸ› Ê «·„ÊŸ› «·„” »œ· —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "During Exchange Employee with another , you must select Employee and change with employee - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 3

                        If .TextMatrix(i, .ColIndex("Emp_id")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ⁄·Ìﬁ «·⁄„·Ì… ·«»œ „‰ «Œ Ì«— „ÊŸ› —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "During  pending you must Select Employee - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If

                    Case 4

                        If .TextMatrix(i, .ColIndex("to_project")) = .TextMatrix(i, .ColIndex("FromProjectName")) And .TextMatrix(i, .ColIndex("to_termid")) = .TextMatrix(i, .ColIndex("current_term")) And .TextMatrix(i, .ColIndex("to_oprid")) = .TextMatrix(i, .ColIndex("current_opr")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "«·„ÊŸ› Ì⁄„· »«·›⁄· ›Ì Â–« «·„Êﬁ⁄ ﬁ„ » €Ì— «·„Êﬁ⁄ —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "this Employee Already work in this location Select another location if you want - error in line:" & CHR(13)
                            End If

                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If
                        
                        If .TextMatrix(i, .ColIndex("to_projectid")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ‰ﬁ· „ÊŸ› „‰ „‘—Ê⁄ ·«Œ— ·«»œ „‰ «Œ Ì«— «·„‘—Ê⁄ «·ÃœÌœ ⁄·Ï «·«ﬁ· —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "in case change employee project you must select at least new project name - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("to_project")
                            Grid.ShowCell i, Grid.ColIndex("to_project")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
              
                    Case 5

                        If .TextMatrix(i, .ColIndex("to_project")) = .TextMatrix(i, .ColIndex("current_project")) And .TextMatrix(i, .ColIndex("to_termid")) = .TextMatrix(i, .ColIndex("current_term")) And .TextMatrix(i, .ColIndex("to_oprid")) = .TextMatrix(i, .ColIndex("current_opr")) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "«·„ÊŸ› Ì⁄„· »«·›⁄· ›Ì Â–« «·„Êﬁ⁄ ﬁ„ » €Ì— «·„Êﬁ⁄ —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "this Employee Already work in this location Select another location if you want - error in line:" & CHR(13)
                            End If

                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                           
                        End If
                        
                        If .TextMatrix(i, .ColIndex("to_projectid")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·…  ‰ﬁ·  Ê«” »œ«· „ÊŸ› „‰ „‘—Ê⁄ ·«Œ— ·«»œ „‰ «Œ Ì«— «·„‘—Ê⁄ «·ÃœÌœ ⁄·Ï «·«ﬁ· —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "in case change employee project you must select at least new project name - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("to_project")
                            Grid.ShowCell i, Grid.ColIndex("to_project")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
                        
                        If .TextMatrix(i, .ColIndex("Emp_id")) = "" Or .TextMatrix(i, .ColIndex("Emp_id1")) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                    
                                Msg = " ›Ì Õ«·… «·‰ﬁ· «·«” »œ«· ·«»œ „‰  ÕœÌœ «·„ÊŸ› Ê «·„ÊŸ› «·„” »œ· —«Ã⁄ «·”ÿ— —ﬁ„ " & CHR(13)
                            Else
                                Msg = "During Exchange Employee with another , you must select Employee and change with employee - error in line:" & CHR(13)
                            End If
    
                            Msg = Msg + CStr(i) & CHR(13)
                            Grid.Row = i
                            Grid.Col = Grid.ColIndex("Emp_Name1")
                            Grid.ShowCell i, Grid.ColIndex("Emp_Name1")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            GoTo ErrTrap
                        End If
                        
                End Select
            
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    Msg = " ·«»œ „‰  ÕœÌœ  Õ«·… «·⁄„·Ì… ›Ì «·”ÿ— —ﬁ„ " & CHR(13)
                Else
                    Msg = " must Specify Operation Type  in row no: " & CHR(13)
                End If

                Msg = Msg + CStr(i) & CHR(13)
                Grid.Row = i
                Grid.Col = Grid.ColIndex("typeid")
                Grid.ShowCell i, Grid.ColIndex("type")
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Grid.SetFocus
                Screen.MousePointer = vbDefault
                GoTo ErrTrap
          
            End If

        Next i
        
    End With

    With Me.Grid

        For i = .FixedRows To .Rows - 1
                            
    
                
                
              '  get_Employee_project_information val(StrAccountCode)

                            
            If .TextMatrix(i, .ColIndex("PK")) <> "" Then
         
                RsDev.AddNew
                RsDev("pk_id").value = Me.XPTxtID.Text
RsDev("toid").value = .TextMatrix(i, .ColIndex("PK"))

                RsDev("Emp_id").value = .TextMatrix(i, .ColIndex("Emp_id"))
                RsDev("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
                RsDev("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
                RsDev("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName"))
                RsDev("opr_type").value = 1
                RsDev("start_date").value = Date
                
                RsDev("opreration_type").value = IIf(.TextMatrix(i, .ColIndex("typeid")) = "", Null, .TextMatrix(i, .ColIndex("typeid")))
                
                RsDev("FromProjectID").value = IIf(.TextMatrix(i, .ColIndex("FromProjectID")) = "", val(DCPROJECT.BoundText), .TextMatrix(i, .ColIndex("FromProjectID")))
                RsDev("Project_id").value = IIf(.TextMatrix(i, .ColIndex("current_project")) = "", 0, .TextMatrix(i, .ColIndex("current_project")))
                RsDev("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_term")) = "", Null, .TextMatrix(i, .ColIndex("current_term")))
                RsDev("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", Null, .TextMatrix(i, .ColIndex("current_opr")))
                
                RsDev("to_project").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", Null, .TextMatrix(i, .ColIndex("to_projectid")))
                RsDev("to_project_name").value = IIf(.TextMatrix(i, .ColIndex("to_project")) = "", Null, .TextMatrix(i, .ColIndex("to_project")))
                
                RsDev("to_term").value = IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", Null, .TextMatrix(i, .ColIndex("to_termid")))
                RsDev("to_opr").value = IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", Null, .TextMatrix(i, .ColIndex("to_oprid")))
              
                RsDev("person_id").value = IIf(.TextMatrix(i, .ColIndex("Emp_id1")) = "", Null, .TextMatrix(i, .ColIndex("Emp_id1")))
                RsDev("person_code").value = IIf(.TextMatrix(i, .ColIndex("emp_code1")) = "", Null, .TextMatrix(i, .ColIndex("emp_code1")))
                RsDev("person_name").value = IIf(.TextMatrix(i, .ColIndex("Emp_Name1")) = "", Null, .TextMatrix(i, .ColIndex("Emp_Name1")))
                RsDev("person_joptypename").value = IIf(.TextMatrix(i, .ColIndex("JobTypeName1")) = "", Null, .TextMatrix(i, .ColIndex("JobTypeName1")))
                RsDev.update

                Select Case val(.TextMatrix(i, .ColIndex("typeid")))

                    Case 1 ' «‰Â«¡

                        ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’
                        If .TextMatrix(i, .ColIndex("pk")) <> "" Then
                            Dim sql As String
                            sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(XPDtbTrans.value) & "' where id =" & val(.TextMatrix(i, .ColIndex("pk")))
                            Cn.Execute sql, , adExecuteNoRecords
                        End If
             
                        'Õ–› «”„ «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›Ì‰
                        If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Then
            
                            sql = "update  TblEmployee  set project_id=0  ,term_fullcode=Null,opr_fullcode=Null where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id")))
                            Cn.Execute sql, , adExecuteNoRecords
                        End If
             
                        ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·Ï «‰Â«¡ Ê ”ÃÌ·  «—ÌŒ «·«‰Â«¡
                        If .TextMatrix(i, .ColIndex("current_opr")) <> "" Then
                            sql = "update  terms_operations  set ended=1, end_date='" & SQLDate(XPDtbTrans.value) & "' where fullcode ='" & .TextMatrix(i, .ColIndex("current_opr")) & "'"
                            Cn.Execute sql, , adExecuteNoRecords
                        End If
         
                    Case 2   '«” »œ«· „ÊŸ› »„ÊŸ› «Œ— ›Ì ‰›” «·„‘—Ê⁄
         
                        If .TextMatrix(i, .ColIndex("Emp_id")) <> "" And .TextMatrix(i, .ColIndex("Emp_id1")) <> "" Then
                            'Õ–› »Ì«‰«   «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›
                            sql = "update  TblEmployee  set project_id=0  ,term_fullcode=Null,opr_fullcode=Null where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id")))
                            Cn.Execute sql, , adExecuteNoRecords
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄ ··„ÊŸ› «·ÃœÌœ ›Ì „·› «·„ÊŸ›
           
                            sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(i, .ColIndex("current_project")) = "", 0, .TextMatrix(i, .ColIndex("current_project"))) & "  ,term_fullcode=" & IIf(.TextMatrix(i, .ColIndex("current_term")) = "", "''", "'" & .TextMatrix(i, .ColIndex("current_term")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", "''", "'" & .TextMatrix(i, .ColIndex("current_opr")) & "'") & " where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id1")))
                            Cn.Execute sql, , adExecuteNoRecords
          
                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ› «·«”«”Ì
                            If .TextMatrix(i, .ColIndex("pk")) <> "" Then
              
                                sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(XPDtbTrans.value) & "' where id =" & val(.TextMatrix(i, .ColIndex("pk")))
                                Cn.Execute sql, , adExecuteNoRecords
                            End If
             
                            ' ”ÃÌ·  Œ’Ì’ «·„ÊŸ› «·ÃœÌœ ⁄·Ï «·⁄„·Ì…
                            Dim Rs1 As ADODB.Recordset
                            Dim rs2 As ADODB.Recordset
                            Dim ID As Integer
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("IDTransfer").value = val(XPTxtID.Text)
                            
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("FromProjectID").value = IIf(.TextMatrix(i, .ColIndex("FromProjectID")) = "", val(DCPROJECT.BoundText), .TextMatrix(i, .ColIndex("FromProjectID")))
                            Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("current_project")) = "", 0, .TextMatrix(i, .ColIndex("current_project")))
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_term")) = "", Null, .TextMatrix(i, .ColIndex("current_term")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", Null, .TextMatrix(i, .ColIndex("current_opr")))
                            Rs1.update
                            
                            Set rs2 = New ADODB.Recordset
        
                            rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            rs2.AddNew
                            rs2("pk_id").value = ID
                            rs2("Emp_id").value = .TextMatrix(i, .ColIndex("Emp_id1"))
                            rs2("IDTransfer").value = val(XPTxtID.Text)
                            rs2("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code1"))
                            rs2("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name1"))
                            rs2("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName1"))
                            rs2("opr_type").value = 0
                            rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            
                            rs2("FromProjectID").value = IIf(.TextMatrix(i, .ColIndex("FromProjectID")) = "", val(DCPROJECT.BoundText), .TextMatrix(i, .ColIndex("FromProjectID")))
                            rs2("Project_id").value = IIf(.TextMatrix(i, .ColIndex("current_project")) = "", Null, .TextMatrix(i, .ColIndex("current_project")))
                            rs2("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_term")) = "", Null, .TextMatrix(i, .ColIndex("current_term")))
                            rs2("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", Null, .TextMatrix(i, .ColIndex("current_opr")))
                            rs2.update
                        End If
                    
                    Case 3 ' ⁄·Ìﬁ „ÊŸ›

                        If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Then
                            'Õ–› »Ì«‰«   «·„‘—Ê⁄ „‰ „·› «·„ÊŸ›
                            sql = "update  TblEmployee  set project_id=0  ,term_fullcode='',opr_fullcode='' where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id")))
                            Cn.Execute sql, , adExecuteNoRecords
                            
                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ›
                            If .TextMatrix(i, .ColIndex("pk")) <> "" Then
              
                                sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "' where id =" & val(.TextMatrix(i, .ColIndex("pk")))
                                Cn.Execute sql, , adExecuteNoRecords
                            End If

                            ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·Ï  ⁄·Ìﬁ
                            '«Ê÷«⁄ «·⁄„·Ì«  0 ﬁÌœ «· ‰›Ì– Ê 1 «‰ Â  Ê 2 „⁄·ﬁ…
                            If .TextMatrix(i, .ColIndex("current_opr")) <> "" Then
                                sql = "update  terms_operations  set ended=2  where fullcode ='" & .TextMatrix(i, .ColIndex("current_opr")) & "'"
                                Cn.Execute sql, , adExecuteNoRecords
                            End If
                        End If

                    Case 4 '  ‰ﬁ· „ÊŸ› „‰ „‘—Ê⁄ «Ê»‰œ «Ê ⁄„·Ì… «·Ï „‘—Ê⁄ «Ê »‰œ «Ê ⁄„·Ì… «Œ—Ï
                    
                        If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Then

                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ›
                            If .TextMatrix(i, .ColIndex("pk")) <> "" Then
              
'                                sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "', ToDate='" & SQLDate(XPDtbTrans.value) & "' "
'                                sql = sql & " ,interval = DATEDIFF(d,FromDate,ToDate) "
'                                sql = sql & " where id =" & val(.TextMatrix(i, .ColIndex("pk")))
'
                                 sql = "Select * from  opr_employee_details  "
                                sql = sql & " where id =" & val(.TextMatrix(i, .ColIndex("pk")))
                                Set Rs1 = New ADODB.Recordset
                                StrSQL = sql
                                Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                                Do While Not Rs1.EOF
                                    GetNoOfDays Rs1!FromDate & "", DateAdd("D", -1, XPDtbTrans.value)
                                    Rs1!ToDate = DateAdd("D", -1, XPDtbTrans.value)
                                    Rs1!end_date = Date
                                  '  Rs1!Fromdate = Date
                                    Rs1!ended = 1
                                    Rs1!interval = NoDay
                                    Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                                    Rs1.update
                                    Rs1.MoveNext
                                Loop
                                
                                'sql & " and
                              '  Cn.Execute sql, , adExecuteNoRecords
                                
                             
                            End If
                           
                            ' ”ÃÌ·  Œ’Ì’ «·„ÊŸ› ›Ì «·⁄„·Ì… «·ÃœÌœ…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                            
                            
                        '    to_projectid
                            
                            Rs1("IDTransfer").value = val(XPTxtID.Text)
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", Null, .TextMatrix(i, .ColIndex("to_termid")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", Null, .TextMatrix(i, .ColIndex("to_oprid")))
                            Rs1.update
                            
                            Set rs2 = New ADODB.Recordset
        
                            rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            rs2.AddNew
                            rs2("pk_id").value = ID
                            rs2("Emp_id").value = .TextMatrix(i, .ColIndex("Emp_id"))
                            rs2("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
                            rs2("IDTransfer").value = val(XPTxtID.Text)
                            rs2("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
                            rs2("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName"))
                            rs2("opr_type").value = 0
                            rs2("start_date").value = XPDtbTrans.value
                            rs2("FromDate").value = XPDtbTrans.value
                            
                            
                            
                            'rs2("opreration_type").value = 2
                            rs2("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", Null, .TextMatrix(i, .ColIndex("to_projectid")))
                            rs2("ProjectID").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                            'Rs2("PrjectCode").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", Null, .TextMatrix(i, .ColIndex("to_projectid")))
                            
                            rs2("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", Null, .TextMatrix(i, .ColIndex("to_termid")))
                            rs2("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", Null, .TextMatrix(i, .ColIndex("to_oprid")))
                            rs2.update
                            
                            If SalaryType(val(.TextMatrix(i, .ColIndex("Emp_id")))) = 4 Then
                                SaveSalaryProject val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("to_projectid"))), , GetTypeEmployee(val(.TextMatrix(i, .ColIndex("Emp_id")))), XPDtbTrans.value
                            Else
                                SaveSalaryCompany val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("to_projectid"))), , XPDtbTrans.value
                            End If
                            SaveSalaryCompany2 val(.TextMatrix(i, .ColIndex("Emp_id"))), val(IIf(.TextMatrix(i, .ColIndex("FromProjectid")) = "", 0, .TextMatrix(i, .ColIndex("FromProjectid"))))
                            ' €ÌÌ— Ê÷⁄ «·⁄„·Ì… «·ﬁœÌ„… «·Ï «‰Â«¡ Ê ”ÃÌ·  «—ÌŒ «·«‰Â«¡
                            '  If .TextMatrix(I, .ColIndex("current_opr")) <> "" Then
                            '      Sql = "update  terms_operations  set ended=1, end_date='" & SQLDate(Date) & "' where fullcode ='" & .TextMatrix(I, .ColIndex("current_opr")) & "'"
                            '      Cn.Execute Sql, , adExecuteNoRecords
                            '  End If
                        
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄ ··„ÊŸ› «·ÃœÌœ ›Ì „·› «·„ÊŸ›
           
                            sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid"))) & "  ,term_fullcode=" & IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", "''", "'" & .TextMatrix(i, .ColIndex("to_termid")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", "''", "'" & .TextMatrix(i, .ColIndex("to_oprid")) & "'") & " where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id")))
                            Cn.Execute sql, , adExecuteNoRecords
                            
                        End If

                    Case 5 '  ‰ﬁ· Ê «” »œ«· „ÊŸ› „‰ „‘—Ê⁄ «Ê»‰œ «Ê ⁄„·Ì… «·Ï „‘—Ê⁄ «Ê »‰œ «Ê ⁄„·Ì… «Œ—Ï
                    
                        If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Then

                            ' ”ÃÌ·  «—ÌŒ «·«‰ Â«¡ ·· Œ’Ì’ ··„ÊŸ› A
                            If .TextMatrix(i, .ColIndex("pk")) <> "" Then
              
'                                  sql = "update  opr_employee_details  set ended=1, end_date='" & SQLDate(Date) & "', ToDate='" & SQLDate(XPDtbTrans.value) & "' "
'                                sql = sql & " ,interval = DATEDIFF(d,FromDate,ToDate) "
'                                sql = sql & " where id =" & val(.TextMatrix(i, .ColIndex("pk")))
'                                'sql & " and
'                                Cn.Execute sql, , adExecuteNoRecords
                                 sql = "Select * from  opr_employee_details  "
                                sql = sql & " where id =" & val(.TextMatrix(i, .ColIndex("pk")))
                                Set Rs1 = New ADODB.Recordset
                                StrSQL = sql
                                Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                                Do While Not Rs1.EOF
                                    GetNoOfDays Rs1!FromDate & "", DateAdd("D", -1, XPDtbTrans.value)
                                    Rs1!ToDate = DateAdd("D", -1, XPDtbTrans.value)
                                   
                                    Rs1!end_date = Date
                                    Rs1!ended = 1
                                    Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                                    Rs1!interval = NoDay
                                    Rs1.update
                                    Rs1.MoveNext
                                Loop
                               
                            End If
                           
                            ' ”ÃÌ·  Œ’Ì’ A «·„ÊŸ› ›Ì «·⁄„·Ì… «·ÃœÌœ…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("IDTransfer").value = val(XPTxtID.Text)
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                            
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", Null, .TextMatrix(i, .ColIndex("to_termid")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", Null, .TextMatrix(i, .ColIndex("to_oprid")))
                            Rs1.update
                            
                            Set rs2 = New ADODB.Recordset
        
                            rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            rs2.AddNew
                            rs2("pk_id").value = ID
                            rs2("Emp_id").value = .TextMatrix(i, .ColIndex("Emp_id"))
                            rs2("IDTransfer").value = val(XPTxtID.Text)
                            rs2("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
                            rs2("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
                            rs2("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName"))
                            rs2("opr_type").value = 0
                            rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            rs2("Project_id").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", Null, .TextMatrix(i, .ColIndex("to_projectid")))
                            rs2("ProjectID").value = IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid")))
                            rs2("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", Null, .TextMatrix(i, .ColIndex("to_termid")))
                            rs2("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", Null, .TextMatrix(i, .ColIndex("to_oprid")))
                            rs2.update
                        
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄  «·ÃœÌœ ··„ÊŸ› A ›Ì „·› «·„ÊŸ›
           
                            sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(i, .ColIndex("to_projectid")) = "", 0, .TextMatrix(i, .ColIndex("to_projectid"))) & "  ,term_fullcode=" & IIf(.TextMatrix(i, .ColIndex("to_termid")) = "", "''", "'" & .TextMatrix(i, .ColIndex("to_termid")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(i, .ColIndex("to_oprid")) = "", "''", "'" & .TextMatrix(i, .ColIndex("to_oprid")) & "'") & " where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id")))
                            Cn.Execute sql, , adExecuteNoRecords
                            
                            '«·„ÊŸ› b «·„” »œ·
                           
                            ' ”ÃÌ·  Œ’Ì’ B «·„ÊŸ› ›Ì «·⁄„·Ì… «·ﬁœÌ„…
                     
                            Set Rs1 = New ADODB.Recordset
                            StrSQL = "select * From opr_Employee  "
                            Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                            Rs1.AddNew
                            ID = CStr(new_id("opr_Employee", "id", "", True))
                            Rs1("id").value = ID
                            Rs1("IDTransfer").value = val(XPTxtID.Text)
                            Rs1("Start_date").value = XPDtbTrans.value
                            Rs1("Project_id").value = IIf(.TextMatrix(i, .ColIndex("current_project")) = "", 0, .TextMatrix(i, .ColIndex("current_project")))
                            
                            Rs1("opr_type").value = 0
                            Rs1("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_term")) = "", Null, .TextMatrix(i, .ColIndex("current_term")))
                            Rs1("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", Null, .TextMatrix(i, .ColIndex("current_opr")))
                            Rs1.update
                            
                            Set rs2 = New ADODB.Recordset
        
                            rs2.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable

                            rs2.AddNew
                            rs2("pk_id").value = ID
                            rs2("IDTransfer").value = val(XPTxtID.Text)
                            rs2("Emp_id").value = .TextMatrix(i, .ColIndex("Emp_id1"))
                            rs2("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code1"))
                            rs2("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name1"))
                            rs2("JobTypeName").value = .TextMatrix(i, .ColIndex("JobTypeName1"))
                            rs2("opr_type").value = 0
                            rs2("start_date").value = Date
                            'rs2("opreration_type").value = 2
                            rs2("Project_id").value = IIf(.TextMatrix(i, .ColIndex("current_project")) = "", Null, .TextMatrix(i, .ColIndex("current_project")))
                            rs2("term_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_term")) = "", Null, .TextMatrix(i, .ColIndex("current_term")))
                            rs2("opr_Fullcode").value = IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", Null, .TextMatrix(i, .ColIndex("current_opr")))
                            rs2.update
                        
             
                            If SalaryType(val(.TextMatrix(i, .ColIndex("Emp_id")))) = 4 Then
                                SaveSalaryProject val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("to_projectid"))), , GetTypeEmployee(val(.TextMatrix(i, .ColIndex("Emp_id")))), XPDtbTrans.value
                            Else
                                SaveSalaryCompany val(.TextMatrix(i, .ColIndex("Emp_id"))), val(.TextMatrix(i, .ColIndex("to_projectid"))), , XPDtbTrans.value
                            End If
                            SaveSalaryCompany2 val(.TextMatrix(i, .ColIndex("Emp_id"))), val(DCPROJECT.BoundText)
                            
                            ' ”ÃÌ·  »Ì«‰«   «·„‘—Ê⁄  «·ÃœÌœ ··„ÊŸ› b ›Ì „·› «·„ÊŸ›
           
                            sql = "update  TblEmployee  set project_id=" & IIf(.TextMatrix(i, .ColIndex("current_project")) = "", 0, .TextMatrix(i, .ColIndex("current_project"))) & "  ,term_fullcode=" & IIf(.TextMatrix(i, .ColIndex("current_term")) = "", "''", "'" & .TextMatrix(i, .ColIndex("current_term")) & "'") & ",opr_fullcode=" & IIf(.TextMatrix(i, .ColIndex("current_opr")) = "", "''", "'" & .TextMatrix(i, .ColIndex("current_opr")) & "'") & " where Emp_id =" & val(.TextMatrix(i, .ColIndex("Emp_id1")))
                            Cn.Execute sql, , adExecuteNoRecords
                        End If

                End Select
                    
            End If
            
            '
        Next i

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
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


Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0
 
            If DoPremis(Do_New, Me.Name, True) = False Then
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

            If DoPremis(Do_Edit, Me.Name, True) = False Then
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

        Case 9
            '   ViewDataList

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim s As String
    Dim i As Integer
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From opr_Employee Where id=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Cn.Execute "delete ProJectMofrdSalar where pk_id=" & val(Me.XPTxtID.Text)
                rs.MoveFirst
                 Cn.Execute "delete opr_employee_details where pk_id=" & val(Me.XPTxtID)
                
                For i = 1 To Grid.Rows - 1
                    If Grid.TextMatrix(i, Grid.ColIndex("pk")) <> "" Then
                    
                        s = "update  opr_employee_details  set ended=0, ToDate=OldToDate "
                        s = s & " ,interval = DATEDIFF(d,FromDate,ToDate) "
                        s = s & " where id =" & val(Grid.TextMatrix(i, Grid.ColIndex("pk")))
                        'sql & " and
                        Cn.Execute s, , adExecuteNoRecords
                    
                    End If
                    
                Next
                s = "Delete opr_employee_details Where IDTransfer = " & val(XPTxtID.Text)
                Cn.Execute s
                s = "Delete opr_employee Where IDTransfer = " & val(XPTxtID.Text)
                Cn.Execute s
                

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
                    clear_all Me
                    TxtModFlg_Change
                   ' XPTxtCurrent.Caption = 0
                   ' XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
     '   XPTxtCurrent.Caption = 0
               '         XPTxtCount.Caption = 0
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available.no records there"
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
  Msg = "Sorry error douring delete data"
  End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim s As String
       s = " select distinct P1.Id  Transaction_ID,IsNull(Pro1.project_name ,p1.Project_name) Emp_Comm , p1.Project_name,P11.Project_name,PDes.[des],PDes1.[des] des2,T1.Start_date Transaction_Date,"
       s = s & "T2.to_opr,t2.to_term "
       s = s & " ,T2.emp_name CusName,T2.[Start_date],T2.person_name  ,T2.to_project_name  , T2.opreration_type PaymentType ,T1.OpraID"
       s = s & "   From opr_Employee T1"
       s = s & " LEFT OUTER JOIN opr_employee_details T2"
       s = s & " ON T2.pk_id = T1.id"
       s = s & " AND T2.opr_type = T1.opr_type"
       s = s & " LEFT OUTER JOIN projects AS p1 ON T1.Project_id = p1.id"
       s = s & " LEFT OUTER JOIN projects AS p11 ON T1.Project_id1  = p11.id"
       s = s & "        LEFT OUTER JOIN projects Pro1"
       s = s & "        ON  Pro1.Id = T2.FromProjectID "
       s = s & " LEFT OUTER JOIN projects_des PDes ON  T1.term_Fullcode = PDes.project_id"
       s = s & " LEFT OUTER JOIN projects_des PDes1 ON  T1.term_Fullcode1  = PDes1.project_id"

       

       s = s & " Where T1.opr_type = 1"
       s = s & " And (T1.id = " & val(Me.XPTxtID.Text) & ")"
       


        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpSalary4.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpSalary4.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(12).AddCurrentValueval (lbTotalMente.Caption)
  xReport.ParameterFields(12).AddCurrentValue (DCPROJECT.Text)
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Private Sub dcproject_Click(Area As Integer)

    If DCPROJECT.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(DCPROJECT.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
            
        My_SQL = " select id,Project_name from projects"
        fill_combo DCPROJECT, My_SQL
    End If
        
End Sub

Private Sub dcproject1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
            
       ' My_SQL = " select id,Project_name from projects"
       ' fill_combo dcproject1, My_SQL
         Dim Dcombos As New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
     Dcombos.GetProjects DCPROJECT1
     
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
YearMonth
    Dim My_SQL As String
    CboYear.Text = year(Date)
    CmbMonth.Text = MonthName(Month(Date))
    CmbMonth.ListIndex = Month(Date) - 1

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT id,name  FROM employee_operations_type "
    Else
        My_SQL = "SELECT id,namee  FROM employee_operations_type "
    End If
                
    fill_combo Me.DCoperationType, My_SQL

   ' My_SQL = " select id,Project_name from projects"
   ' fill_combo dcproject, My_SQL

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
     Dcombos.GetProjects DCPROJECT1
     Dcombos.GetProjects DCPROJECT
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

    Me.Caption = "Projects Labors Transfer  "
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

    Check1.Caption = "Show All Employee"

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("Emp_code")) = "Emp_code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp_Name"
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
        .TextMatrix(0, .ColIndex("Emp_Name1")) = "To Employee"

    End With

End Sub

Public Sub get_all_employee(Optional sql As String = "")
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer
    Dim xx As Integer
    xx = 0
    Dim i As Integer
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
    Grid.Enabled = True
          
    If sql = "" Then
        sql = "Select * from emp_all_details "
        xx = 1
    End If
 
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
                    
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                       
                If xx = 1 Then
                    .TextMatrix(i, .ColIndex("PK")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                Else
                    .TextMatrix(i, .ColIndex("PK")) = IIf(IsNull(Rs3.Fields("id").value), "", Rs3.Fields("id").value)
                End If
                       
                .TextMatrix(i, .ColIndex("current_opr")) = IIf(IsNull(Rs3("opr_fullcode").value), "", Rs3("opr_fullcode").value)
                    
                .TextMatrix(i, .ColIndex("current_term")) = IIf(IsNull(Rs3("term_Fullcode").value), "", Rs3("term_Fullcode").value)
                     
                .TextMatrix(i, .ColIndex("current_project")) = IIf(IsNull(Rs3("Project_id").value), 0, Rs3("Project_id").value)
          
                Rs3.MoveNext
            Next i
  
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

Public Sub Grid_AfterEdit(ByVal Row As Long, _
                          ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
            Case "Emp_Code"
                
               
                
                StrSQL = "SELECT  emp_all_details.*,projects.project_name from emp_all_details "
                StrSQL = StrSQL & " lEFT Outer join projects On projects.Id =  emp_all_details.Project_id"
                StrSQL = StrSQL & " Where Emp_id = " & val(StrAccountCode)
''
                StrSQL = " SELECT TT.Id ,TblEmployee.Emp_Name, TT.project_id,projects.project_name,TT.Emp_id,TT.Emp_Code, TT.Emp_Name ,T1.JobTypeName FROM emp_all_details T1"
                StrSQL = StrSQL & " Right Outer JOIN opr_employee_details TT ON T1.Emp_ID = TT.Emp_ID AND TT.Project_id = T1.project_id"
                StrSQL = StrSQL & " Right Outer join projects On projects.Id =  TT.Project_id"
                StrSQL = StrSQL & " Right Outer join TblEmployee On TblEmployee.Emp_ID =  T1.Emp_ID"
                
                StrSQL = StrSQL & " Where TT.project_id <> 0 And TT.opr_type = 0"
             '   StrSQL = StrSQL & " and  TT.project_id = " & mProjectID
                StrSQL = StrSQL & " and  TT.Emp_Code = N'" & Trim(.TextMatrix(Row, .ColIndex("Emp_Code"))) & "'"
'
                
                
              '  get_Employee_project_information val(StrAccountCode)
                Set rs = Nothing
            
                If Trim(.TextMatrix(Row, .ColIndex("Emp_Code"))) <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        '.TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                        .TextMatrix(Row, .ColIndex("Pk")) = rs!ID & ""
                        'If xx = 1 Then
                           '.TextMatrix(Row, .ColIndex("PK")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
                        'Else
                          ' .TextMatrix(Row, .ColIndex("PK")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                        'End If
                        .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                        .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                        .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                        .TextMatrix(Row, .ColIndex("FromProjectID")) = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
                        .TextMatrix(Row, .ColIndex("FromProjectName")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
                        
                            
                    End If
                End If
         
 
            Case "Emp_Name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                Dim mProjectID As Integer
                mProjectID = IIf(.TextMatrix(Row, .ColIndex("FromProjectID")) = "", val(DCPROJECT.BoundText), val(.TextMatrix(Row, .ColIndex("FromProjectID"))))
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PK"), False, True)
                .TextMatrix(Row, .ColIndex("Pk")) = StrAccountCode
                
                StrSQL = "SELECT  emp_all_details.*,projects.project_name from emp_all_details "
                StrSQL = StrSQL & " lEFT Outer join projects On projects.Id =  emp_all_details.Project_id"
                StrSQL = StrSQL & " Where Emp_id = " & val(StrAccountCode)
''
                StrSQL = " SELECT TT.Id ,TT.project_id,projects.project_name,TT.Emp_id,TT.Emp_Code, TT.Emp_Name ,T1.JobTypeName FROM emp_all_details T1"
                StrSQL = StrSQL & " Right Outer JOIN opr_employee_details TT ON T1.Emp_ID = TT.Emp_ID AND TT.Project_id = T1.project_id"
                StrSQL = StrSQL & " Right Outer join projects On projects.Id =  TT.Project_id"
                StrSQL = StrSQL & " Where TT.project_id <> 0 And TT.opr_type = 0"
             '   StrSQL = StrSQL & " and  TT.project_id = " & mProjectID
                StrSQL = StrSQL & " and  TT.Emp_id = " & val(StrAccountCode)
'
                
                
              '  get_Employee_project_information val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("Emp_Code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                        .TextMatrix(Row, .ColIndex("Pk")) = rs!ID & ""
                        'If xx = 1 Then
                           '.TextMatrix(Row, .ColIndex("PK")) = IIf(IsNull(rs.Fields("Emp_id").value), "", rs.Fields("Emp_id").value)
                        'Else
                          ' .TextMatrix(Row, .ColIndex("PK")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                        'End If
                        .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                        .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                        .TextMatrix(Row, .ColIndex("FromProjectID")) = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
                        .TextMatrix(Row, .ColIndex("FromProjectName")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
                        
                            
                    End If
                End If
         
            Case "Emp_Name1"
       
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id1"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_id1")) = StrAccountCode
             
                StrSQL = "SELECT  * from emp_all_details Where Emp_id=" & val(StrAccountCode)
                get_Employee_project_information val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("Emp_Code1")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                            
                        .TextMatrix(Row, .ColIndex("JobTypeName1")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                            
                    End If
                End If
        
            Case "Emp_Code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from emp_all_details Where Emp_Code=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    .TextMatrix(Row, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                    .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
              
                Else
                    .TextMatrix(Row, .ColIndex("JobTypeName")) = ""
              
                    .TextMatrix(Row, .ColIndex("Emp_Name")) = ""
              
                    .TextMatrix(Row, .ColIndex("Emp_id")) = ""
              
                End If

            Case "type"
                StrAccountCode = .ComboData
                ' LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_id"), False, True)
                .TextMatrix(Row, .ColIndex("typeid")) = StrAccountCode

            Case "to_project"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("to_projectid")) = StrAccountCode

            Case "FromProjectName"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("FromProjectID")) = StrAccountCode

            Case "to_term"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("to_termid")) = StrAccountCode
             
            Case "to_opr"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("to_oprid")) = StrAccountCode
             
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" Then
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
 
        .ComboList = ""

        Select Case .ColKey(Col)

            Case "Emp_Code"
                .ComboList = ""
Case "FromProjectName"
            .ComboList = ""
            Cancel = True
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
'        project_search.show
'        project_search.case_id = 1

    End If

    With Me.Grid

        If KeyCode = 13 Then
        '  get_Employee_project_information val(.TextMatrix(.Row, .ColIndex("Emp_id")))
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
 
    With Me.Grid

        Select Case .ColKey(Col)

            Case "Emp_Name"
                Dim mProjectID As Integer
                mProjectID = IIf(.TextMatrix(Row, .ColIndex("FromProjectID")) = "", val(DCPROJECT.BoundText), val(.TextMatrix(Row, .ColIndex("FromProjectID"))))

                StrSQL = "SELECT *  FROM emp_all_details TT where project_id<>0 "
                StrSQL = StrSQL & "and  TT.project_id <> 0 "
                'And TT.opr_type = 0"
                'StrSQL = StrSQL & " and  TT.project_id = " & val(mProjectID)


                
'                StrSQL = " SELECT TT.id Emp_id,TT.Emp_Name  FROM emp_all_details T1"
'                StrSQL = StrSQL & " RIGHT Outer JOIN opr_employee_details TT ON T1.Emp_ID = TT.Emp_ID AND TT.Project_id = T1.project_id"
'                StrSQL = StrSQL & " Where TT.project_id <> 0 And TT.opr_type = 0"
'                StrSQL = StrSQL & " and  TT.project_id = " & val(dcproject.BoundText)
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "emp_name", "Emp_id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "Emp_Name1"

                StrSQL = "SELECT *  FROM emp_all_details   "
    
                StrSQL = " SELECT TT.id Emp_id,TT.Emp_Name  FROM emp_all_details T1"
                StrSQL = StrSQL & " RIGHT Outer JOIN opr_employee_details TT ON T1.Emp_ID = TT.Emp_ID AND TT.Project_id = T1.project_id"
                StrSQL = StrSQL & " Where TT.project_id <> 0 And TT.opr_type = 0"
                StrSQL = StrSQL & " and  TT.project_id = " & val(DCPROJECT.BoundText)
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "emp_name", "Emp_id")
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "emp_name", "Emp_id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "type"
                StrSQL = "SELECT *  FROM employee_operations_type "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = Grid.BuildComboList(rs, "namee", "id")
                End If
                
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
 
            Case "to_project"
               
                StrSQL = "SELECT id,Project_name From  projects "
                StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "Project_name", "id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            Case "FromProjectName"
                               
'                StrSQL = "SELECT projects.id,projects.Project_name  FROM emp_all_details TT "
'                StrSQL = StrSQL & " Inner join projects On projects.Id = TT.project_id"
'                StrSQL = StrSQL & " where TT.project_id<>0 "
'                StrSQL = StrSQL & "and  TT.project_id <> 0 "
'
'                StrSQL = "SELECT id,Project_name From  projects "
'                StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
'
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                StrComboList = Grid.BuildComboList(rs, "Project_name", "id")
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'
'                .ComboList = StrComboList

            Case "to_term"
                StrSQL = "SELECT *  FROM projects_des "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "des", "fullcode")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "to_opr"
                StrSQL = "SELECT *  FROM terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Grid.BuildComboList(rs, "name", "fullcode")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.XPTxtID.Text = IIf(IsNull(rs("id").value), "", rs("id").value)
 
    XPDtbTrans.value = IIf(IsNull(rs("Start_date").value), Date, rs("Start_date").value)
 
    'from
    DCPROJECT.BoundText = IIf(IsNull(rs("Project_id").value), "", rs("Project_id").value)
    Dcterm.BoundText = IIf(IsNull(rs("term_Fullcode").value), "", rs("term_Fullcode").value)
    dcopr.BoundText = IIf(IsNull(rs("opr_Fullcode").value), "", rs("opr_Fullcode").value)

    'to
     CboYear.ListIndex = IIf(IsNull(rs("Years").value), -1, rs("Years").value)
    CmbMonth.ListIndex = IIf(IsNull(rs("Months").value), -1, rs("Months").value)
    
    DCPROJECT1.BoundText = IIf(IsNull(rs("Project_id1").value), "", rs("Project_id1").value)
    Dcterm1.BoundText = IIf(IsNull(rs("term_Fullcode1").value), "", rs("term_Fullcode1").value)
    dcopr1.BoundText = IIf(IsNull(rs("opr_Fullcode1").value), "", rs("opr_Fullcode1").value)

    txtType.Text = IIf(IsNull(rs("opr_type").value), 1, rs("opr_type").value)
 
    StrSQL = "select opr_employee_details.*,projects.project_name from opr_employee_details "
    StrSQL = StrSQL & " Left Outer join projects"
    StrSQL = StrSQL & " Right Outer join projects On projects.Id =  opr_employee_details.Project_id "
    
    StrSQL = StrSQL & " Where pk_id = " & Me.XPTxtID.Text
    
     
    StrSQL = " SELECT opr_employee_details.*,"
    StrSQL = StrSQL & "        Pro1.project_name,Pro2.project_name Toproject_name"
    StrSQL = StrSQL & " From opr_employee_details"
    StrSQL = StrSQL & "        LEFT OUTER JOIN projects Pro1"
    StrSQL = StrSQL & "        ON  Pro1.Id = opr_employee_details.FromProjectID "
    StrSQL = StrSQL & "        LEFT  OUTER JOIN projects Pro2"
    StrSQL = StrSQL & "            ON  Pro2.Id = opr_employee_details.ProjectID"
    StrSQL = StrSQL & " Where pk_id = " & Me.XPTxtID.Text
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_id").value), "", RsDev("Emp_id").value)
                .TextMatrix(i, .ColIndex("PK")) = IIf(IsNull(RsDev("toid").value), "", RsDev("toid").value)
            
                .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Emp_code").value), "", RsDev("Emp_code").value)
                .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(RsDev("emp_name").value), "", RsDev("emp_name").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("FromProjectID")) = IIf(IsNull(RsDev("FromProjectID").value), val(DCPROJECT.BoundText), RsDev("FromProjectID").value)
                .TextMatrix(i, .ColIndex("FromProjectName")) = IIf(IsNull(RsDev("project_name").value), (DCPROJECT.Text), RsDev("project_name").value)
            
                .TextMatrix(i, .ColIndex("Emp_id1")) = IIf(IsNull(RsDev("person_id").value), "", RsDev("person_id").value)
    
                .TextMatrix(i, .ColIndex("Emp_code1")) = IIf(IsNull(RsDev("person_code").value), "", RsDev("person_code").value)
                .TextMatrix(i, .ColIndex("emp_name1")) = IIf(IsNull(RsDev("person_name").value), "", RsDev("person_name").value)
            
                .TextMatrix(i, .ColIndex("JobTypeName1")) = IIf(IsNull(RsDev("person_joptypename").value), "", RsDev("person_joptypename").value)
             
                .TextMatrix(i, .ColIndex("current_project")) = IIf(IsNull(RsDev("Project_id").value), "", RsDev("Project_id").value)
            
                .TextMatrix(i, .ColIndex("current_term")) = IIf(IsNull(RsDev("term_Fullcode").value), "", RsDev("term_Fullcode").value)
            
                .TextMatrix(i, .ColIndex("current_opr")) = IIf(IsNull(RsDev("opr_Fullcode").value), "", RsDev("opr_Fullcode").value)
            
                .TextMatrix(i, .ColIndex("to_projectid")) = IIf(IsNull(RsDev("to_project").value), "", RsDev("to_project").value)
            
                .TextMatrix(i, .ColIndex("to_project")) = IIf(IsNull(RsDev("to_project_name").value), "", RsDev("to_project_name").value)
            
                .TextMatrix(i, .ColIndex("to_termid")) = IIf(IsNull(RsDev("to_term").value), "", RsDev("to_term").value)
                       
                .TextMatrix(i, .ColIndex("to_oprid")) = IIf(IsNull(RsDev("to_opr").value), "", RsDev("to_opr").value)
            
                .TextMatrix(i, .ColIndex("typeid")) = IIf(IsNull(RsDev("opreration_type").value), "", RsDev("opreration_type").value)

                Select Case val(.TextMatrix(i, .ColIndex("typeid")))
            
                    Case 1

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("type")) = "«‰Â«¡"
                        Else
                            .TextMatrix(i, .ColIndex("type")) = "Finish"
                        End If

                    Case 2

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("type")) = " »œÌ· „ÊŸ›"
                        Else
                            .TextMatrix(i, .ColIndex("type")) = "Change Emp"
                        End If

                    Case 3

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("type")) = " ⁄·Ìﬁ"
                        Else
                            .TextMatrix(i, .ColIndex("type")) = "Pending"
                        End If

                    Case 4

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("type")) = "‰ﬁ· ·⁄„·Ì…"
                        Else
                            .TextMatrix(i, .ColIndex("type")) = "change operation"
                        End If
              
                End Select
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
' fillapprovData
    ReLineGrid
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

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

Sub SaveSalaryProject(Optional empid As Double, Optional ProjID As Double, Optional NoDay2 As Double, Optional TypeEmp As Integer, Optional FromDate As Date, Optional ToDate As Date)
    Dim Rs7 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql As String
    Dim Value1 As Double
    
   
'    If IsDate(ToDate) Then
'        'NoDay =
'        NoDay = DateDiff("d", Fromdate, ToDate)
'    Else
'        NoDay = DateDiff("d", Fromdate, MonthLastDay(Fromdate))
'    End If
   ' GetNoOfDays Fromdate, ToDate
    If IsDate(FromDate) Then

'
                GetNoOfDays FromDate, ""
    End If
    Set Rs7 = New ADODB.Recordset
    Rs7.Open "ProJectMofrdSalar", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    sql = "Select * from ProJectMofrd  where ProjID=" & ProjID & " and  TypeEmp =" & TypeEmp & ""
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs4.RecordCount > 0 Then
    Rs4.MoveFirst
    Dim i As Integer
        For i = 1 To Rs4.RecordCount
                Rs7.AddNew
                Value1 = 0
                Rs7("pk_id").value = Me.XPTxtID.Text
                Rs7("EmpID").value = empid
                Rs7("ProjID").value = ProjID
                Rs7("NoDay").value = NoDay
                Rs7("YearID").value = val(CboYear.ListIndex)
                Rs7("MonthID").value = val(CmbMonth.ListIndex)
                Value1 = IIf(IsNull(Rs4("Valuee").value), 0, Rs4("Valuee").value)
                Value1 = Value1 / 30
                Rs7("Valuee").value = Round(Value1, 2)
                Rs7("MofrdID").value = IIf(IsNull(Rs4("MofrdID").value), 0, Rs4("MofrdID").value)
                Rs7("Total").value = Round((Value1 * NoDay), 0)
                Rs7("TypeSalary").value = 1
                Rs7("FromDate").value = FromDate
                
                If IsDate(ToDate) Then
                    Rs7("ToDate").value = ToDate
                Else
                    
                    Rs7("ToDate").value = Null
                End If
                
                Rs7.update
                Rs4.MoveNext
        Next i
    End If
End Sub
Function SalaryType(Optional Emp_id As Double) As Integer
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     Emp_ID, SalaryType"
sql = sql & " From dbo.TblEmployee"
sql = sql & " Where (Emp_id = " & Emp_id & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
SalaryType = IIf(IsNull(Rs7("SalaryType").value), 0, Rs7("SalaryType").value)
Else
SalaryType = 0
End If
End Function




Sub SaveSalaryCompany2(Optional empid As Double, Optional ProjID As Double, Optional NoDay2 As Double, Optional FromDate As Date, Optional ToDate As String = "")
    Dim Rs7 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql As String
    Dim Value1 As Double
    
    
    Set Rs7 = New ADODB.Recordset
    sql = "Select * from ProJectMofrdSalar Where ProjID =   " & ProjID & " and EmpID = " & empid
    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic
    Do While Not Rs7.EOF
                
        
       ' Rs7("FromDate").value = FromDate
 
         If IsDate(Rs7!FromDate) Then
        'NoDay =
            'NoDay = DateDiff("d", Rs7!Fromdate, XPDtbTrans.value)
             'GetNoOfDays Rs7!Fromdate & "", XPDtbTrans.value
            GetNoOfDays Rs7!FromDate & "", DateAdd("D", -1, XPDtbTrans.value)
                                    
        Else
'            NoDay = DateDiff("d", Rs7!Fromdate, MonthLastDay(Rs7!Fromdate))
            GetNoOfDays Rs7!FromDate & "", ""
        End If
       
        Rs7("ToDate").value = XPDtbTrans.value
        Rs7("NoDay").value = NoDay

        
        Rs7("Total").value = Round((val(Rs7("Valuee").value & "") * NoDay), 0)
        Rs7.update
        Rs7.MoveNext
    Loop
    
End Sub

Sub SaveSalaryCompany(Optional empid As Double, Optional ProjID As Double, Optional NoDay2 As Double, Optional FromDate As Date, Optional ToDate As String = "")
    Dim Rs7 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql As String
    Dim Value1 As Double
    YearMonth
'    If IsDate(ToDate) Then
'        'NoDay =
'        NoDay = DateDiff("d", Fromdate, ToDate)
'    Else
'        NoDay = DateDiff("d", Fromdate, MonthLastDay(Fromdate))
'    End If
            If IsDate(FromDate) Then

'            NoDay = DateDiff("d", Rs7!Fromdate, MonthLastDay(Rs7!Fromdate))
                GetNoOfDays FromDate, ""
             End If
  
    Set Rs7 = New ADODB.Recordset
    Rs7.Open "ProJectMofrdSalar", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    sql = "Select * from dbo.EmpSalaryComponent  where emp_ID=" & empid & ""
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs4.RecordCount > 0 Then
    Rs4.MoveFirst
    Dim i As Integer
        For i = 1 To Rs4.RecordCount
                CboYear.Text = year(XPDtbTrans.value)
                CmbMonth.Text = MonthName(Month(XPDtbTrans.value))
                CmbMonth.ListIndex = Month(XPDtbTrans.value) - 1
                Rs7.AddNew
                Value1 = 0
                
                Rs7("pk_id").value = Me.XPTxtID.Text
                Rs7("EmpID").value = empid
                Rs7("ProjID").value = ProjID
                Rs7("NoDay").value = NoDay
                Rs7("YearID").value = val(CboYear.ListIndex)
                Rs7("MonthID").value = val(CmbMonth.ListIndex)
                
                
                Value1 = IIf(IsNull(Rs4("Value").value), 0, Rs4("Value").value)
                Value1 = Value1 / 30
                
                Rs7("Valuee").value = Round(Value1, 2)
                Rs7("MofrdID").value = IIf(IsNull(Rs4("AccountCode").value), 0, Rs4("AccountCode").value)
                Rs7("Total").value = Round((Value1 * NoDay), 0)
                Rs7("TypeSalary").value = 0
                Rs7("FromDate").value = FromDate
                If IsDate(ToDate) Then
                    Rs7("ToDate").value = ToDate
                Else
                    
                    Rs7("ToDate").value = Null
                End If
                Rs7.update
                Rs4.MoveNext
        Next i
    End If
End Sub
Public Function MonthLastDay(ByVal dCurrDate As Date)
  Dim dFirstDayNextMonth As Date
  
  On Error GoTo lbl_Error
 
  MonthLastDay = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
  MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
  Exit Function
lbl_Error:
  MsgBox Err.description, vbOKOnly + vbExclamation
End Function

Function GetTypeEmployee(Optional empid As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select TypeEmp from TblEmployee where Emp_ID=" & empid & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetTypeEmployee = IIf(IsNull(Rs3("TypeEmp").value), 0, Rs3("TypeEmp").value) + 1
Else
GetTypeEmployee = 0
End If
End Function

Private Sub GetNoOfDays(ByVal mFromDate As String, ByVal mToDate As String)
    If IsDate(mFromDate) And Not IsDate(mToDate) Then
        NoDay = 30 - Day(mFromDate) + 1
    ElseIf IsDate(mToDate) Then
        NoDay = DateDiff("d", mFromDate, mToDate) + 1
        
    End If
    
End Sub
