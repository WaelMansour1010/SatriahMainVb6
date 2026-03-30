VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSerivcesOperations 
   Caption         =   "›« Ê—… Œœ„« "
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   Icon            =   "FrmSerivcesOperations.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   8580
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8580
      _cx             =   15134
      _cy             =   12356
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
      Align           =   5
      AutoSizeChildren=   8
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
      GridRows        =   6
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSerivcesOperations.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   615
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   15
         Visible         =   0   'False
         Width           =   510
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6000
         Width           =   8550
         _cx             =   15081
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
            Height          =   375
            Left            =   6315
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4005
            TabIndex        =   3
            Top             =   45
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·›« Ê—…"
            Height          =   255
            Index           =   0
            Left            =   7335
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   75
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   1
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   765
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   105
            Width           =   750
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   135
            Width           =   630
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   330
            Index           =   4
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   75
            Width           =   915
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1530
         Index           =   0
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   645
         Width           =   8550
         _cx             =   15081
         _cy             =   2699
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   675
            Index           =   7
            Left            =   780
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   840
            Width           =   2955
            _cx             =   5212
            _cy             =   1191
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
            Begin VB.CheckBox XPChkGoOut 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " „ «· ”·Ì„"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   1740
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   0
               Width           =   1170
            End
            Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
               Height          =   315
               Left            =   120
               TabIndex        =   84
               Top             =   210
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               Format          =   100073473
               CurrentDate     =   38784
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”·Ì„ «·›⁄·Ï"
               Height          =   315
               Index           =   7
               Left            =   1485
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   270
               Width           =   1410
            End
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   690
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   1965
         End
         Begin VB.TextBox XPTxtMaintanenceID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5100
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   90
            Width           =   2235
         End
         Begin VB.TextBox TxtTransID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1605
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   420
            Visible         =   0   'False
            Width           =   1065
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   315
            Left            =   4665
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   420
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
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
            ButtonImage     =   "FrmSerivcesOperations.frx":0420
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5100
            TabIndex        =   15
            Top             =   450
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoInDtae 
            Height          =   315
            Left            =   5100
            TabIndex        =   16
            Top             =   795
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   5100
            TabIndex        =   87
            Top             =   1140
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”·Ì„"
            Height          =   315
            Index           =   9
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” ·«„"
            Height          =   315
            Index           =   3
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   255
            Index           =   5
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   90
            Width           =   1065
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   315
            Index           =   6
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   465
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   315
            Index           =   8
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   120
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   3795
         Left            =   15
         TabIndex        =   21
         Top             =   2190
         Width           =   8550
         _cx             =   15081
         _cy             =   6694
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…"
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
         Picture(0)      =   "FrmSerivcesOperations.frx":07BA
         Picture(1)      =   "FrmSerivcesOperations.frx":0B54
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3330
            Index           =   2
            Left            =   45
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Width           =   8460
            _cx             =   14923
            _cy             =   5874
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
            AutoSizeChildren=   8
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
            GridRows        =   2
            GridCols        =   2
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmSerivcesOperations.frx":0EEE
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   705
               Index           =   5
               Left            =   0
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   2625
               Width           =   8460
               _cx             =   14923
               _cy             =   1244
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
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   2160
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   315
                  Width           =   1905
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   480
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   315
                  Width           =   1530
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   4125
                  TabIndex        =   26
                  Top             =   315
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   6405
                  TabIndex        =   27
                  Top             =   315
                  Width           =   2010
                  _ExtentX        =   3545
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   390
                  Left            =   60
                  TabIndex        =   28
                  Top             =   270
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmSerivcesOperations.frx":0F35
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
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·Œœ„…"
                  Height          =   255
                  Index           =   31
                  Left            =   6525
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   0
                  Width           =   1890
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·Œœ„…"
                  Height          =   255
                  Index           =   30
                  Left            =   4380
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   255
                  Index           =   28
                  Left            =   2325
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   0
                  Width           =   1560
               End
               Begin VB.Label Lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   255
                  Index           =   26
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1470
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2625
               Index           =   6
               Left            =   0
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   465
               _cx             =   820
               _cy             =   4630
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
               Begin ImpulseButton.ISButton XPBtnRemove 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   735
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   556
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
                  ButtonImage     =   "FrmSerivcesOperations.frx":12CF
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  LowerToggledContent=   0   'False
               End
               Begin ImpulseButton.ISButton XPBtnAdd 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   210
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   556
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
                  ButtonImage     =   "FrmSerivcesOperations.frx":1669
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
                  LowerToggledContent=   0   'False
               End
               Begin ImpulseButton.ISButton CmdReplace 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   36
                  Top             =   1260
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   14871017
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmSerivcesOperations.frx":1A03
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  Alignment       =   1
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2625
               Left            =   465
               TabIndex        =   37
               Top             =   0
               Width           =   7995
               _cx             =   14102
               _cy             =   4630
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSerivcesOperations.frx":1D9D
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
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3330
            Index           =   4
            Left            =   9195
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   45
            Width           =   8460
            _cx             =   14923
            _cy             =   5874
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
            AutoSizeChildren=   8
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
            GridRows        =   12
            GridCols        =   6
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmSerivcesOperations.frx":1EA3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ"
               Height          =   570
               Index           =   2
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   2085
               Width           =   1125
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "¬Ã· "
               Height          =   360
               Index           =   1
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1140
               Width           =   1125
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœ«"
               Height          =   375
               Index           =   0
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   540
               Width           =   1125
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   1125
               Index           =   2
               Left            =   2205
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   2085
               Width           =   4980
               Begin VB.TextBox XPTxtChqueNum 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2970
                  MaxLength       =   40
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   75
                  Width           =   975
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Index           =   2
                  Left            =   2970
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   465
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DCboBankName 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   56
                  Top             =   90
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker XPDTPDueDate 
                  Height          =   345
                  Left            =   60
                  TabIndex        =   57
                  Top             =   465
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   100073473
                  CurrentDate     =   38784
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   19
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   525
                  Width           =   1155
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   210
                  Index           =   17
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   135
                  Width           =   645
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   210
                  Index           =   18
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   142
                  Width           =   735
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   16
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   532
                  Width           =   465
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   945
               Index           =   1
               Left            =   2205
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1140
               Width           =   4980
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   90
                  Width           =   1545
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   2730
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   90
                  Width           =   1425
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   330
                  Left            =   150
                  TabIndex        =   49
                  Top             =   480
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   100073473
                  CurrentDate     =   38784
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   540
                  Width           =   1155
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   15
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   150
                  Width           =   465
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   150
                  Width           =   525
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   585
               Index           =   0
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   480
               Width           =   7185
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5400
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   120
                  Width           =   1155
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3120
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   120
                  Width           =   1545
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   42
                  Top             =   120
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   180
                  Width           =   465
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   150
                  Width           =   525
               End
               Begin VB.Label Lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   270
                  Index           =   22
                  Left            =   2310
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   150
                  Width           =   765
               End
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   480
               Index           =   20
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   15
         Width           =   8550
         _cx             =   15081
         _cy             =   1085
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
         Caption         =   "›« Ê—… Œœ„« "
         Align           =   0
         AutoSizeChildren=   8
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
         GridRows        =   3
         GridCols        =   10
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmSerivcesOperations.frx":1F87
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1335
            TabIndex        =   67
            Top             =   120
            Width           =   540
            _ExtentX        =   953
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
            ButtonImage     =   "FrmSerivcesOperations.frx":202F
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
            Left            =   735
            TabIndex        =   68
            Top             =   120
            Width           =   540
            _ExtentX        =   953
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
            ButtonImage     =   "FrmSerivcesOperations.frx":23C9
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
            Left            =   1935
            TabIndex        =   69
            Top             =   120
            Width           =   540
            _ExtentX        =   953
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
            ButtonImage     =   "FrmSerivcesOperations.frx":2763
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
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   555
            _ExtentX        =   979
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
            ButtonImage     =   "FrmSerivcesOperations.frx":2AFD
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
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   6450
         Width           =   8550
         _cx             =   15081
         _cy             =   953
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   7650
            TabIndex        =   73
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   6690
            TabIndex        =   74
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   5685
            TabIndex        =   75
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   4800
            TabIndex        =   76
            Top             =   90
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   3795
            TabIndex        =   77
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   2895
            TabIndex        =   78
            Top             =   90
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   30
            TabIndex        =   79
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   1860
            TabIndex        =   80
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   375
            Left            =   960
            TabIndex        =   81
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
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
   End
End
Attribute VB_Name = "FrmSerivcesOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Resize_Form Me, TransactionSize
End Sub
