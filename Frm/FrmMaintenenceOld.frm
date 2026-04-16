VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintenenceOLd 
   Caption         =   "œŒÊ· ··’Ì«‰…"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   HelpContextID   =   80
   Icon            =   "FrmMaintenenceOld.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   8580
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7005
      Left            =   0
      TabIndex        =   22
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
      _GridInfo       =   $"FrmMaintenenceOld.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   55
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
            Left            =   6330
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   30
            Width           =   960
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4020
            TabIndex        =   57
            Top             =   45
            Width           =   1230
            _ExtentX        =   2170
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
            Height          =   330
            Index           =   4
            Left            =   5340
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   75
            Width           =   945
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   135
            Width           =   675
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   105
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   1
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·›« Ê—…"
            Height          =   255
            Index           =   0
            Left            =   7305
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   75
            Width           =   1230
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1530
         Index           =   0
         Left            =   15
         TabIndex        =   23
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
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Text            =   "Combo1"
            Top             =   60
            Width           =   1425
         End
         Begin VB.TextBox TxtTransID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   750
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox CboMaintenanceType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   2205
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   750
            Width           =   1305
         End
         Begin VB.TextBox XPTxtMaintanenceID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6390
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   90
            Width           =   1185
         End
         Begin VB.CheckBox XPChkGoOut 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " „ «· ”·Ì„"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5490
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   60
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   255
            Left            =   4110
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   450
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
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
            ButtonImage     =   "FrmMaintenenceOld.frx":0420
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   4410
            TabIndex        =   1
            Top             =   450
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoInDtae 
            Height          =   315
            Left            =   4410
            TabIndex        =   2
            Top             =   795
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
            Height          =   315
            Left            =   4410
            TabIndex        =   3
            Top             =   1155
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   87
            Top             =   30
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   180
            TabIndex        =   89
            Top             =   1140
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdSearchTrans 
            Height          =   345
            Left            =   900
            TabIndex        =   91
            Top             =   750
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmMaintenenceOld.frx":07BA
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   5
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   90
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   255
            Index           =   24
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1140
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   25
            Left            =   2370
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’Ì«‰…"
            Height          =   255
            Index           =   10
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   420
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ›« Ê—… «·»Ì⁄"
            Height          =   255
            Index           =   9
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   315
            Index           =   8
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Œ—ÊÃ «·„ Êﬁ⁄"
            Height          =   375
            Index           =   7
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   315
            Index           =   6
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   465
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·œŒÊ·"
            Height          =   315
            Index           =   3
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   810
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   3795
         Left            =   15
         TabIndex        =   26
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
         Picture(0)      =   "FrmMaintenenceOld.frx":0B54
         Picture(1)      =   "FrmMaintenenceOld.frx":0EEE
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3330
            Index           =   2
            Left            =   45
            TabIndex        =   27
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
            _GridInfo       =   $"FrmMaintenenceOld.frx":1288
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   705
               Index           =   5
               Left            =   0
               TabIndex        =   77
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   495
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   315
                  Width           =   1560
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   2145
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   315
                  Width           =   1950
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   4155
                  TabIndex        =   9
                  Top             =   315
                  Width           =   2160
                  _ExtentX        =   3810
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   6345
                  TabIndex        =   8
                  Top             =   315
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   390
                  Left            =   60
                  TabIndex        =   78
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
                  ButtonImage     =   "FrmMaintenenceOld.frx":12CF
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
                  Caption         =   "«· ﬂ·›…"
                  Height          =   255
                  Index           =   26
                  Left            =   555
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   0
                  Width           =   1500
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   255
                  Index           =   28
                  Left            =   2325
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   0
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   255
                  Index           =   30
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   0
                  Width           =   2025
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   255
                  Index           =   31
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   0
                  Width           =   1935
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2625
               Index           =   6
               Left            =   0
               TabIndex        =   31
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
                  TabIndex        =   32
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
                  ButtonImage     =   "FrmMaintenenceOld.frx":1669
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
                  TabIndex        =   33
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
                  ButtonImage     =   "FrmMaintenenceOld.frx":1A03
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
                  TabIndex        =   83
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
                  ButtonImage     =   "FrmMaintenenceOld.frx":1D9D
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
               TabIndex        =   7
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
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmMaintenenceOld.frx":2137
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
            TabIndex        =   34
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
            _GridInfo       =   $"FrmMaintenenceOld.frx":2261
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   570
               Index           =   0
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   510
               Width           =   7185
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3120
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   120
                  Width           =   1545
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5400
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   120
                  Width           =   1155
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   84
                  Top             =   120
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   270
                  Index           =   22
                  Left            =   2310
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   150
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   4770
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   150
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   180
                  Width           =   465
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   930
               Index           =   1
               Left            =   2205
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1140
               Width           =   4980
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   2730
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   90
                  Width           =   1425
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   90
                  Width           =   1545
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   330
                  Left            =   150
                  TabIndex        =   16
                  Top             =   480
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   100073473
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   150
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   15
                  Left            =   4260
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   150
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   540
                  Width           =   1155
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   1080
               Index           =   2
               Left            =   2205
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2130
               Width           =   4980
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Index           =   2
                  Left            =   2970
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   465
                  Width           =   975
               End
               Begin VB.TextBox XPTxtChqueNum 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2970
                  MaxLength       =   40
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   75
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DCboBankName 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   19
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
                  TabIndex        =   21
                  Top             =   465
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   100073473
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   16
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   532
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   210
                  Index           =   18
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   142
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   210
                  Index           =   17
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   135
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   19
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   525
                  Width           =   1155
               End
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœ«"
               Height          =   360
               Index           =   0
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   570
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
               TabIndex        =   14
               Top             =   1140
               Width           =   1125
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ"
               Height          =   495
               Index           =   2
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   2130
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   510
               Index           =   20
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   7
         Left            =   15
         TabIndex        =   50
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
         Caption         =   "œŒÊ· ··’Ì«‰…"
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
         _GridInfo       =   $"FrmMaintenenceOld.frx":2346
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1335
            TabIndex        =   51
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
            ButtonImage     =   "FrmMaintenenceOld.frx":23EE
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
            TabIndex        =   52
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
            ButtonImage     =   "FrmMaintenenceOld.frx":2788
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
            TabIndex        =   53
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
            ButtonImage     =   "FrmMaintenenceOld.frx":2B22
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
            TabIndex        =   54
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
            ButtonImage     =   "FrmMaintenenceOld.frx":2EBC
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   615
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   15
         Visible         =   0   'False
         Width           =   750
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   64
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
            Left            =   7635
            TabIndex        =   65
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   6675
            TabIndex        =   66
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   5700
            TabIndex        =   67
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   4770
            TabIndex        =   68
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   3825
            TabIndex        =   69
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   2880
            TabIndex        =   70
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            TabIndex        =   71
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
            Left            =   1890
            TabIndex        =   72
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            TabIndex        =   73
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
Attribute VB_Name = "FrmMaintenenceOLd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim MaintenReport As ClsMaintananceReport
Dim cSearchDcbo(4) As clsDCboSearch

Public BolPrint As Boolean

Private Sub C1Elastic6_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboMaintenanceType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboMaintenanceType.ListIndex = 0 Then
            TxtTransSerial.Enabled = False
            lbl(9).Enabled = False
            CmdReplace.Enabled = False
        Else
            TxtTransSerial.Enabled = True
            lbl(9).Enabled = True
            CmdReplace.Enabled = True
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboMaintenanceType_Click()
    CboMaintenanceType_Change
End Sub

Private Sub cmdAdd_Click()
    '“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer

    If DCboItemsCode.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsCode.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    'If CboItemCase.ListIndex = -1 Then
    '    Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
    '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    CboItemCase.SetFocus
    '    SendKeys "{F4}"
    '    Exit Sub
    'End If
    'If TxtQuantity.Text = "" Then
    '    Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
    '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    TxtQuantity.SetFocus
    '    Exit Sub
    'End If
    'If Not IsNumeric(TxtQuantity.Text) Then
    '    Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
    '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    TxtQuantity.SetFocus
    '    Exit Sub
    'End If
    If TxtPrice.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «· ﬂ·›…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtPrice.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(TxtPrice.text) Then
        Msg = "«· ﬂ·›… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtPrice.SetFocus
        Exit Sub
    End If

    With FG

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
        End If

        .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
        .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
        '    If CboItemCase.ListIndex <> -1 Then
        '        .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = CboItemCase.ListIndex + 1
        '    End If
        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = TxtSerial.text
        '    .TextMatrix(.Rows - 1, .ColIndex("Count")) = TxtQuantity.Text
        .TextMatrix(.Rows - 1, .ColIndex("Cost")) = TxtPrice.text

        If TxtSerial.Tag = "T" Then
            .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
        ElseIf TxtSerial.Tag = "F" Then
            .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
        End If

    End With

    DCboItemsCode.BoundText = ""
    DCboItemsName.BoundText = ""
    TxtSerial.text = ""
    TxtPrice.text = ""
    XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
    FG.SetFocus
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdReplace_Click()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsSerial As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·’‰› «·–Ì  —€» ›Ì «” »œ«·Â "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If DBCboClientName.text = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·" & Chr(13)
        Msg = Msg + "«·–Ì ﬁ«„ »‘—«¡ Â–Â «·ﬁÿ⁄…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If CboMaintenanceType.ListIndex = 1 Then
        If TxtTransSerial.text = "" Then
            Msg = Msg + "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ " & Chr(13)
            Msg = Msg + "«· Ì  „ »Ì⁄ Â–« «·’‰› ›ÌÂ«"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Sub
        End If
    End If

    '«· √ﬂœ √‰ «·ﬁÿ⁄… ﬁœ  „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ… ›Ì Õ«·… «·’Ì«‰…  »⁄ «·÷„«‰
    If CboMaintenanceType.ListIndex = 1 Then
        If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" Then
            If FG.Cell(flexcpChecked, FG.Row, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
                    StrSQL = StrSQL + " AND Transaction_Serial ='" & val(TxtTransSerial.text) & "'"
                    StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
                    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If RsSerial.EOF Or RsSerial.BOF Then
                        Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                        Msg = Msg + FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = FG.Row
                        FG.Col = FG.ColIndex("Name")
                        FG.ShowCell FG.Row, FG.ColIndex("Name")
                        FG.SetFocus
                        Exit Sub
                    End If
                
                    If IsNull(RsSerial("guaranteeTime").value) Then
                        Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                        Msg = Msg + FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "·Ì” ·Â« ÷„«‰"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = FG.Row
                        FG.Col = FG.ColIndex("Name")
                        FG.ShowCell FG.Row, FG.ColIndex("Name")
                        FG.SetFocus
                        Exit Sub
                    End If

                    If (DateDiff("d", XPDtbGoInDtae.value, DateAdd("m", RsSerial("guaranteeTime").value, RsSerial("Transaction_Date").value))) < 0 Then
                        Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
                        Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").value & Chr(13)
                        Msg = Msg + "–«  «·”Ì—Ì«·  " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").value, "yyyy/m/d") & Chr(13)
                        Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").value & Chr(13)
                        Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").value & "  ‘Â—" & Chr(13)
                        Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"

                        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
                            XPTab301.CurrTab = 0
                            FG.Row = FG.Row
                            FG.Col = FG.ColIndex("Name")
                            FG.ShowCell FG.Row, FG.ColIndex("Name")
                            FG.SetFocus
                            Exit Sub
                        End If
                    End If

                    RsSerial.Close
                Else
                    Msg = "ÌÃ»  ÕœÌœ «·”Ì—Ì«· «·Œ«’ »«·ﬁÿ⁄… «· Ì  —€» ›Ì «” »œ«·Â«"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If

            Else
                Msg = "Â–Â «·⁄„·Ì… Œ«’… »«·√’‰«› «· Ì   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If
    End If

    FG.Tag = FG.Row

    With FrmReplace
        .TxtTransID.text = Me.TxtTransID.text
        .TxtTransSerial.text = Me.TxtTransSerial.text
        .XPTxtMaintanenceID.text = XPTxtMaintanenceID.text
        .DCboItemsName.BoundText = FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
        .Tag = FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Code"))
        .TxtItemSerial.text = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
        .Show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DCboItemsCode.BoundText <> "" Then
        DCboItemsName.BoundText = DCboItemsCode.BoundText
    Else
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsCode.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtSerial.Enabled = True
            '        TxtQuantity.Enabled = False
            '        TxtQuantity.Text = "1"
            TxtSerial.Tag = "T"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtSerial.Enabled = False
            '        TxtQuantity.Enabled = True
            '        TxtQuantity.Text = ""
            TxtSerial.Tag = "F"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DCboItemsName.BoundText <> "" Then
        DCboItemsCode.BoundText = DCboItemsName.BoundText
    Else
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtSerial.Enabled = True
            ''        TxtQuantity.Enabled = False
            '        TxtQuantity.Text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtSerial.Enabled = False
            '        TxtQuantity.Enabled = True
            '        TxtQuantity.Text = ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Ele_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 1 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim RsSerial As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim Msg As String
    Dim StrSQL As String

    If XPDtbGoInDtae.value = "" Then
        Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ⁄„·Ì… «·’Ì«‰…" & Chr(13)
        Msg = Msg + "ﬁ»· ≈œŒ«· »Ì«‰«  «·√’‰«›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPDtbGoInDtae.SetFocus
        Exit Sub
    End If

    If Col = FG.ColIndex("Name") Then
        If FG.TextMatrix(Row, FG.ColIndex("Name")) <> "" Then
            FG.TextMatrix(Row, FG.ColIndex("Code")) = FG.TextMatrix(Row, FG.ColIndex("Name"))

            If IsNumeric(FG.TextMatrix(Row, FG.ColIndex("Code"))) Then
                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(Row, FG.ColIndex("Code"))
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.EOF Or RsTemp.BOF Then
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                Else

                    If RsTemp("HaveSerial").value = True Then
                        FG.TextMatrix(Row, FG.ColIndex("HaveSerial")) = True
                    Else
                        FG.TextMatrix(Row, FG.ColIndex("HaveSerial")) = False
                    End If
                End If

            Else
                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If
    End If

    If Col = FG.ColIndex("Code") Then
        If FG.TextMatrix(Row, FG.ColIndex("Code")) <> "" Then
            FG.TextMatrix(Row, FG.ColIndex("Name")) = FG.TextMatrix(Row, FG.ColIndex("Code"))
            StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(Row, FG.ColIndex("Code")) & ""

            If IsNumeric(FG.TextMatrix(Row, FG.ColIndex("Code"))) Then
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.EOF Or RsTemp.BOF Then
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                Else

                    If RsTemp("HaveSerial").value = True Then
                        FG.TextMatrix(Row, FG.ColIndex("HaveSerial")) = True
                    Else
                        FG.TextMatrix(Row, FG.ColIndex("HaveSerial")) = False
                    End If
                End If

            Else
                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If
    End If

    If CboMaintenanceType.ListIndex = 1 Then
        If FG.TextMatrix(Row, FG.ColIndex("Code")) <> "" Then
            If FG.Cell(flexcpChecked, Row, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(Row, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(Row, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(Row, FG.ColIndex("Serial")) & "'"
                    StrSQL = StrSQL + " AND Transaction_Serial='" & val(TxtTransSerial.text) & "'"
                    StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
                    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If RsSerial.EOF Or RsSerial.BOF Then
                        Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                        Msg = Msg + FG.TextMatrix(Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    
                        '»Ì«‰«  «·›« Ê—… «· Ì  „ »Ì⁄ «·ﬁÿ⁄Â ›ÌÂ«
                        StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(Row, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(Row, FG.ColIndex("Serial")) & "'"
                        Set RsTemp = New ADODB.Recordset
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            Msg = "·ﬁœ  „ »Ì⁄ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")) & Chr(13)
                            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(Row, FG.ColIndex("Serial")) & Chr(13)
                            Msg = Msg + "≈·Ï «·⁄„Ì· : " & RsTemp("CusName").value & Chr(13)
                            Msg = Msg + "›Ì «·›« Ê—… —ﬁ„ : " & RsTemp("Transaction_ID").value
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        End If

                        XPTab301.CurrTab = 0
                        FG.Row = Row
                        FG.Col = FG.ColIndex("Name")
                        FG.ShowCell Row, FG.ColIndex("Name")
                        FG.SetFocus
                        Exit Sub
                    End If

                    If IsNull(RsSerial("guaranteeTime").value) Then
                        Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                        Msg = Msg + FG.TextMatrix(Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "·Ì” ·Â« ÷„«‰"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = Row
                        FG.Col = FG.ColIndex("Name")
                        FG.ShowCell Row, FG.ColIndex("Name")
                        FG.SetFocus
                        Exit Sub
                    End If

                    If (DateDiff("d", XPDtbGoInDtae.value, DateAdd("m", RsSerial("guaranteeTime").value, RsSerial("Transaction_Date").value))) < 0 Then
                        Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
                        Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").value & Chr(13)
                        Msg = Msg + "–«  «·”Ì—Ì«·  " & FG.TextMatrix(Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").value, "yyyy/m/d") & Chr(13)
                        Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").value & Chr(13)
                        Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").value & "  ‘Â—" & Chr(13)
                        Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"

                        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
                            XPTab301.CurrTab = 0
                            FG.Row = Row
                            FG.Col = FG.ColIndex("Name")
                            FG.ShowCell Row, FG.ColIndex("Name")
                            FG.SetFocus
                            Exit Sub
                        End If
                    End If

                    RsSerial.Close
                End If
            End If
        End If
    End If

    XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)
    On Error GoTo ErrTrap

    If Col = FG.ColIndex("HaveSerial") Then
        Cancel = True
    End If

    With FG

        If .TextMatrix(Row, .ColIndex("MType")) <> "" Then
            If .TextMatrix(Row, .ColIndex("MType")) = 2 Then
                If Col = .ColIndex("Cost") Then
                    .TextMatrix(Row, .ColIndex("Cost")) = 0
                    Cancel = True
                End If
            End If
        End If

        If .TextMatrix(Row, .ColIndex("HaveSerial")) <> "" Then
            If .TextMatrix(Row, .ColIndex("HaveSerial")) = False Then
                If Col = .ColIndex("Serial") Then
                    Cancel = True
                End If
            End If
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        FrmAddNewItem.DealingForm = Maintenance
        FrmAddNewItem.Show vbModal
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
    '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«·
    Dim StrSQL As String
    Dim RsReplace As ADODB.Recordset

    If Me.TxtModFlg.text <> "R" Then Exit Sub

    With FG

        If .Col = .ColIndex("NewItem") Then
            If .Cell(flexcpData, .Row, .ColIndex("Replace")) <> "" Then
                If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = "select * From ReplacedItems where MaintenanceID=" & XPTxtMaintanenceID.text
                    StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & Chr(13)
                        Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
                        Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").value & Chr(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ﬁÿ⁄…  „ «” »œ«·Â«"
                    End If
                End If
            End If
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim BGround As New ClsBackGroundPic
    Dim RsItems As New ADODB.Recordset
    Dim StrList As String
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    XPTab301.CurrTab = 0
    Me.Height = 7515
    Me.Width = 8700
    Resize_Form Me
    AddTip
    SetDtpickerDate Me.XPDtbGoInDtae
    SetDtpickerDate XPDtbGoOutDtae

    FG.WallPaper = BGround.Picture
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    StrSQL = "SELECT * From TblCustemers"
    fill_combo Me.DBCboClientName, StrSQL
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    StrSQL = "SELECT * From BanksData"
    fill_combo Me.DcboBankName, StrSQL
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboBankName

    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL

    With CboPaymentType
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"
    End With

    With CboMaintenanceType
        .AddItem "»«· ﬂ·›…"
        .AddItem " »⁄ «·÷„«‰"
    End With

    Set rs = New ADODB.Recordset
    rs.Open "[TblMaintenece]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "select * From TblItems"
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    StrList = FG.BuildComboList(RsItems, "ItemName", "ItemID")

    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("Name")) = "|" & StrList
    End If

    StrList = FG.BuildComboList(RsItems, "ItemCode", "ItemID")

    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("Code")) = "|" & StrList
    End If

    FG.ColComboList(FG.ColIndex("MType")) = "#1;»«· ﬂ·›…|#2; »⁄ «·÷„«‰"
    FillItemData
    Retrive
    Me.TxtModFlg.text = "R"
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set MaintenReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            Me.Caption = "»Ì«‰«  ⁄„·Ì«  «·’Ì«‰…"
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
            XPDtbGoInDtae.Enabled = False
            XPDtbGoOutDtae.Enabled = False
            DBCboClientName.locked = True
            XPBtnNewClients.Enabled = False
            '        XPMTxtRemarks.Locked = True
            XPChkGoOut.Enabled = False
            XPBtnAdd.Enabled = False
            XPBtnRemove.Enabled = False
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            TxtTransSerial.locked = True
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            XPTxtChqueNum.Enabled = False
            DcboBankName.Enabled = False
            XPTxtValue(2).Enabled = False
            XPDTPDueDate.Enabled = False
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            CboPaymentType.locked = True
            CboMaintenanceType.locked = True
            DtpDelayDate.Enabled = False
            CmdReplace.Enabled = False
            Ele(5).Enabled = False

        Case "N"
            Me.Caption = "»Ì«‰«  ⁄„·Ì«  «·’Ì«‰…( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPDtbGoInDtae.Enabled = True
            XPDtbGoOutDtae.Enabled = True
            DBCboClientName.locked = False
            XPBtnNewClients.Enabled = True
            '        XPMTxtRemarks.Locked = False
            XPChkGoOut.Enabled = True
            TxtTransSerial.locked = False
            XPBtnNewClients.Enabled = True

            If XPTab301.CurrTab = 0 Then
                XPBtnAdd.Enabled = True
                XPBtnRemove.Enabled = True
            Else
                XPBtnAdd.Enabled = False
                XPBtnRemove.Enabled = False
            End If

            FG.Enabled = True
            FG.Rows = FG.FixedRows
            FG.Rows = 2
            '        FG.RowPosition(FG.Rows - 2) = FG.Rows - 1
            '        FG.TextMatrix(FG.Rows - 1, 2) = "«÷€ÿ Â‰«"
            Me.DBCboClientName.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPChkGoOut.value = Unchecked
            XPDtbGoInDtae.value = Date
            XPDtbGoOutDtae.value = Date
            CboPaymentType.locked = False
            CboMaintenanceType.locked = False
            CboPaymentType.ListIndex = 0
            CboMaintenanceType.ListIndex = 0
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            XPDTPDueDate.value = Date
            CboMaintenanceType_Change
            Ele(5).Enabled = True

        Case "E"
            Me.Caption = "»Ì«‰«  ⁄„·Ì«  «·’Ì«‰…(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            TxtTransID.locked = False
            TxtTransSerial.locked = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPDtbGoInDtae.Enabled = True
            XPDtbGoOutDtae.Enabled = True
            DBCboClientName.locked = False
            XPBtnNewClients.Enabled = True
            '        XPMTxtRemarks.Locked = False
            XPChkGoOut.Enabled = True
        
            If XPTab301.CurrTab = 0 Then
                XPBtnAdd.Enabled = True
                XPBtnRemove.Enabled = True
            Else
                XPBtnAdd.Enabled = False
                XPBtnRemove.Enabled = False
            End If

            XPBtnRemove.Enabled = True
            FG.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DcboBankName.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPDTPDueDate.Enabled = True
            DtpDelayDate.Enabled = True

            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPaymentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            CboMaintenanceType.locked = False
            FG.Editable = flexEDKbdMouse
            CboPaymentType.locked = False
            DBCboClientName_Change
            CboMaintenanceType_Change
            Ele(5).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPrice_KeyDown(KeyCode As Integer, _
                             Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        cmdAdd_Click
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.text = "R" Then
        If Trim(Me.TxtTransID.text) = "" Then
            Me.TxtTransSerial.text = ""
        Else
            Me.TxtTransSerial.text = GetTransIDSerial(1, val(Me.TxtTransID.text), , 2)
        End If
    End If

End Sub

Private Sub TxtTransSerial_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Trim(Me.TxtTransSerial.text) = "" Then
            Me.TxtTransID.text = ""
        Else
            Me.TxtTransID.text = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)
        End If
    End If

End Sub

Private Sub XPBtnAdd_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Code")) <> "" Then
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
        FG.Col = FG.ColIndex("Code")
        FG.ShowCell FG.Rows - 1, FG.ColIndex("Code")
        FG.SetFocus
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
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

Private Sub Cmd_Click(Index As Integer)
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim AskOption As Boolean
    Dim intDef As Integer
    BolPrint = True
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            clear_all Me
            TxtModFlg.text = "N"
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            Me.DcboBox.BoundText = 1
            XPTxtMaintanenceID.text = CStr(new_id("TblMaintenece", "MaintananceID", "", True))
            XPTab301.CurrTab = 0
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1

        Case 1
            '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
            StrSQL = "select * From  Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« «Â«"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
            SaveData

        Case 3
            Call Undo

        Case 4
            Del_TransAction

        Case 5
            FrmMaintanenceSearch.Show vbModal

        Case 7
            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.Show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If

            PrintingData

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnNewClients_Click()
    On Error GoTo ErrTrap

    With FrmAddNewCustemer
        .DealingForm = Maintenance
        .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
        .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
        .lbl(0).Caption = "«”„ «·⁄„Ì·"
        .Show vbModal
        cSearchDcbo(0).Refresh
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.Rows = 2 Then
        FG.Clear flexClearScrollable, flexClearEverything
    Else

        If FG.Rows > 1 Then
            If FG.Row <> FG.FixedRows - 1 Then
                FG.RemoveItem (FG.Row)
            End If
        End If
    End If

    XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    XPTxtSerial(1).text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
                '            XPTxtSerial(1).Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(2).text = ""
                    XPTxtChqueNum.text = ""
                    XPDTPDueDate.value = Date
                    DcboBankName.BoundText = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(2).Enabled = True
                    XPTxtChqueNum.Enabled = True
                    XPDTPDueDate.Enabled = True
                    XPTxtValue(2).locked = False
                    XPTxtChqueNum.locked = False
                    DcboBankName.locked = False
                    DcboBankName.Enabled = True
                End If

            Else
                XPTxtValue(2).text = ""
                XPTxtValue(2).Enabled = False
                XPTxtChqueNum.Enabled = False
                XPDTPDueDate.Enabled = False
                DcboBankName.Enabled = False
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RsNotes As New ADODB.Recordset
    Dim RsDetails As New ADODB.Recordset
    Dim RsSerial As New ADODB.Recordset
    Dim RsCheckSerial As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim RsReplaceDetails As ADODB.Recordset
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim ReplaceID As Integer
    Dim Msg As String
    Dim BeginTrans As Boolean

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If DBCboClientName.text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If XPChkGoOut.value = Checked Then
            If XPChkPayType(0).value = Unchecked And XPChkPayType(1).value = Unchecked And XPChkPayType(2).value = Unchecked Then
                Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄" & Chr(13)
                Msg = Msg + "Ê«·ﬁÌ„… «·„«·Ì… «·„Õ’·…" & Chr(13)
                Msg = Msg + "„ﬁ«»· ⁄„·Ì… «·’Ì«‰…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTab301.CurrTab = 1
                Exit Sub
            End If
        End If

        If CboMaintenanceType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·’Ì«‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboMaintenanceType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If CboMaintenanceType.ListIndex = 1 Then
            If TxtTransSerial.text = "" Then
                Msg = "·ﬁœ ﬁ„  »«Œ Ì«— ’Ì«‰…  »⁄ «·÷„«‰ " & Chr(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Sub
            End If
        End If

        For RowNum = 1 To FG.Rows - 1

            If IsNumeric(Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))) Then
                StrSQL = "select * From TblItems where ItemID=" & Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsSerial.EOF Or RsSerial.BOF Then
                    Msg = "·« ÊÃœ √Ì »Ì«‰«  ⁄‰  «·’‰›" & Chr(13)
                    Msg = Msg + Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code"))) & Chr(13)
                    Msg = Msg + "≈–« ﬂ«‰ ·„ Ì „  ”ÃÌ·Â" & Chr(13)
                    Msg = Msg + "≈÷€ÿ Œ«‰… ’‰› ÃœÌœ" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTab301.CurrTab = 0
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("Name")
                    FG.ShowCell RowNum, FG.ColIndex("Name")
                    FG.SetFocus
                    Exit Sub
                End If

                If CboMaintenanceType.ListIndex <> 1 Then
                    If val(FG.TextMatrix(RowNum, FG.ColIndex("cost"))) = 0 Then
                        Msg = "·„ Ì „  ÕœÌœ «· ﬂ·›… «·Œ«’ »«·’‰›" & Chr(13)
                        Msg = Msg + RsSerial("ItemName").value
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("cost")
                        FG.ShowCell RowNum, FG.ColIndex("cost")
                        FG.SetFocus
                        Exit Sub
                    End If
                End If

                RsSerial.Close
            Else
                Msg = "·« ÊÃœ √Ì »Ì«‰«  ⁄‰  «·’‰›" & Chr(13)
                Msg = Msg + Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code"))) & Chr(13)
                Msg = Msg + "≈–« ﬂ«‰ ·„ Ì „  ”ÃÌ·Â" & Chr(13)
                Msg = Msg + "≈÷€ÿ Œ«‰… ’‰› ÃœÌœ" & Chr(13)
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTab301.CurrTab = 0
                FG.Row = RowNum
                FG.Col = FG.ColIndex("Name")
                FG.ShowCell RowNum, FG.ColIndex("Name")
                FG.SetFocus
                Exit Sub
            End If

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If RsSerial("HaveSerial").value = True Then
                        Msg = "«·’‰› " & Chr(13)
                        Msg = Msg + RsSerial("ItemName").value & Chr(13)
                        Msg = Msg + "·Â ”Ì—Ì«·" & Chr(13)
                        Msg = Msg + "„‰ ›÷·ﬂ √œŒ· «·”Ì—Ì«· «·Œ«’ »Â"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("Serial")
                        FG.ShowCell RowNum, FG.ColIndex("Serial")
                        FG.SetFocus
                        Exit Sub
                    End If

                    RsSerial.Close
                End If
            End If

            '«· √ﬂœ √‰ «·ﬁÿ⁄… ·Ì”  „ÊÃÊœ… ›Ì ⁄„·Ì… ’Ì«‰… √Œ—Ï
            Select Case Me.TxtModFlg.text

                Case "N"

                    If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                        If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                            StrSQL = "select * From QryMaintananceReport where ItemID=" & Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                            StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                            StrSQL = StrSQL + " and GoOut=False"
                            Set RsTemp = New ADODB.Recordset
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                                Msg = " „ «” ·«„ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                                Msg = Msg + "–«  «·”Ì—Ì«·        : " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                Msg = Msg + "„‰ «·⁄„Ì·            :  " & RsTemp("CusName").value & Chr(13)
                                Msg = Msg + "» «—ÌŒ                   :  " & RsTemp("DateGoIN").value & Chr(13)
                                Msg = Msg + "·Ì „ ≈Ã—«¡ ⁄„·Ì… ’Ì«‰… ·Â« Ê·„ Ì „  ”·Ì„Â« »⁄œ" & Chr(13)
                                Msg = Msg + " —ﬁ„ «·⁄„·Ì…         : " & RsTemp("MaintananceID").value
                                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                XPTab301.CurrTab = 0
                                FG.Row = RowNum
                                FG.Col = FG.ColIndex("Name")
                                FG.ShowCell RowNum, FG.ColIndex("Name")
                                FG.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If

                Case "E"

                    If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                        If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                            StrSQL = "select * From QryMaintananceReport where ItemID=" & Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                            StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                            StrSQL = StrSQL + " and GoOut=False"
                            Set RsTemp = New ADODB.Recordset
                            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                                If RsTemp("MaintananceID").value <> XPTxtMaintanenceID.text Then
                                    Msg = " „ «” ·«„ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                                    Msg = Msg + "–«  «·”Ì—Ì«·        : " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                    Msg = Msg + "„‰ «·⁄„Ì·            :  " & RsTemp("CusName").value & Chr(13)
                                    Msg = Msg + "» «—ÌŒ                   :  " & RsTemp("DateGoIN").value & Chr(13)
                                    Msg = Msg + "·Ì „ ≈Ã—«¡ ⁄„·Ì… ’Ì«‰… ·Â« Ê·„ Ì „  ”·Ì„Â« »⁄œ" & Chr(13)
                                    Msg = Msg + " —ﬁ„ «·⁄„·Ì…         : " & RsTemp("MaintananceID").value
                                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    XPTab301.CurrTab = 0
                                    FG.Row = RowNum
                                    FG.Col = FG.ColIndex("Name")
                                    FG.ShowCell RowNum, FG.ColIndex("Name")
                                    FG.SetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

            End Select

            '«· √ﬂœ √‰ «·ﬁÿ⁄… ﬁœ  „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ… ›Ì Õ«·… «·’Ì«‰…  »⁄ «·÷„«‰
            If CboMaintenanceType.ListIndex = 1 Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                    If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                        If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                            StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                            StrSQL = StrSQL + " AND Transaction_Serial='" & Trim(TxtTransSerial.text) & "'"
                            StrSQL = StrSQL + " AND Transaction_Type=2"
                            StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
                            RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                            If RsSerial.EOF Or RsSerial.BOF Then
                                Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                                Msg = Msg + FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
                                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
                                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            
                                '»Ì«‰«  «·›« Ê—… «· Ì  „ »Ì⁄ «·ﬁÿ⁄Â ›ÌÂ«
                                StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                                Set RsTemp = New ADODB.Recordset
                                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                                    Msg = "·ﬁœ  „ »Ì⁄ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                                    Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                    Msg = Msg + "≈·Ï «·⁄„Ì· : " & RsTemp("CusName").value & Chr(13)
                                    Msg = Msg + "›Ì «·›« Ê—… —ﬁ„ : " & RsTemp("Transaction_Serial").value
                                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                End If
                            
                                XPTab301.CurrTab = 0
                                FG.Row = RowNum
                                FG.Col = FG.ColIndex("Name")
                                FG.ShowCell RowNum, FG.ColIndex("Name")
                                FG.SetFocus
                                Exit Sub
                            End If

                            If IsNull(RsSerial("guaranteeTime").value) Then
                                Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
                                Msg = Msg + FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                Msg = Msg + "·Ì” ·Â« ÷„«‰"
                                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                XPTab301.CurrTab = 0
                                FG.Row = RowNum
                                FG.Col = FG.ColIndex("Name")
                                FG.ShowCell RowNum, FG.ColIndex("Name")
                                FG.SetFocus
                                Exit Sub
                            End If

                            If (DateDiff("d", XPDtbGoInDtae.value, DateAdd("m", RsSerial("guaranteeTime").value, RsSerial("Transaction_Date").value))) < 0 Then
                                Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
                                Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").value & Chr(13)
                                Msg = Msg + "–«  «·”Ì—Ì«·  " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
                                Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").value, "yyyy/m/d") & Chr(13)
                                Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").value & Chr(13)
                                Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").value & "  ‘Â—" & Chr(13)
                                Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"

                                If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
                                    XPTab301.CurrTab = 0
                                    FG.Row = RowNum
                                    FG.Col = FG.ColIndex("Name")
                                    FG.ShowCell RowNum, FG.ColIndex("Name")
                                    FG.SetFocus
                                    Exit Sub
                                End If
                            End If

                            RsSerial.Close
                        End If
                    End If
                End If
            End If

            '«·√’‰«› «· Ì ·Ì” ·Â« ”Ì—Ì«· ›Ì Õ«·… «·’Ì«‰…  »⁄ «·÷„«‰
            If CboMaintenanceType.ListIndex = 1 Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexUnchecked Then
                    StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                    StrSQL = StrSQL + " AND Transaction_Serial='" & val(TxtTransSerial.text) & "'"
                    StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
                    Set RsSerial = New ADODB.Recordset
                    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If RsSerial.EOF Or RsSerial.BOF Then
                        Msg = "«·’‰› : "
                        Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                        Msg = Msg + "·„ Ì „ »Ì⁄ √Ì ﬂ„Ì… „‰Â ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("Name")
                        FG.ShowCell RowNum, FG.ColIndex("Name")
                        FG.SetFocus
                        Exit Sub
                    Else
                        Msg = "«·’‰›" & Chr(13)
                        Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                        Msg = Msg + "·Ì” ·Â ”Ì—Ì«·" & Chr(13)
                        Msg = Msg + "«·√’‰«› «· Ì ·Ì” ·Â« ”Ì—Ì«· ·Ì” ·Â« ÷„«‰" & Chr(13)
                        Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"

                        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
                            XPTab301.CurrTab = 0
                            FG.Row = RowNum
                            FG.Col = FG.ColIndex("Name")
                            FG.ShowCell RowNum, FG.ColIndex("Name")
                            FG.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Next RowNum

        If CboPaymentType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPaymentType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.XPChkPayType(0).value = vbChecked Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If XPChkPayType(2).value = vbChecked Then
            If DcboBankName.BoundText = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If Trim(Me.XPTxtChqueNum.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!!"
                Screen.MousePointer = vbDefault
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If Check_CheckNum(Me.XPTxtChqueNum.text, val(Me.XPTxtMaintanenceID.text), Me.TxtModFlg.text, 1) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If val(XPTxtValue(0).text) + val(XPTxtValue(1).text) + val(XPTxtValue(2).text) > val(XPTxtSum.text) Then
            Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & Chr(13)
            Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTab301.CurrTab = 1
            Exit Sub
        End If

        If val(XPTxtValue(0).text) + val(XPTxtValue(1).text) + val(XPTxtValue(2).text) < val(XPTxtSum.text) Then
            Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & Chr(13)
            Msg = Msg + "√ﬁ· „‰ ≈Ã„«·Ì «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTab301.CurrTab = 1
            Exit Sub
        End If
    
        If Me.TxtModFlg.text = "E" Then
            StrSQL = "delete From TblMainteneceDetails where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            StrSQL = "delete From Notes where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            StrSQL = "delete From MaintenanceJuncTransaction where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            StrSQL = "delete From Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        RsDetails.Open "[TblMainteneceDetails]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        If Me.TxtModFlg.text = "N" Then
            rs.AddNew
        End If

        Cn.BeginTrans
        BeginTrans = True
        rs("MaintananceID").value = val(XPTxtMaintanenceID.text)
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
        rs("DateGoIN").value = XPDtbGoInDtae.value
        rs("DateGoOUT").value = XPDtbGoOutDtae.value
        rs("GoOut").value = XPChkGoOut.value
        rs("UserID").value = user_id

        If CboPaymentType.ListIndex = -1 Then
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPaymentType.ListIndex)
        End If

        If CboMaintenanceType.ListIndex = -1 Then
            rs("MType").value = 0
        Else
            rs("MType").value = val(CboMaintenanceType.ListIndex)
        End If

        If CboMaintenanceType.ListIndex = 1 Then
            rs("Transaction_ID").value = IIf(Me.TxtTransID.text = "", Null, val(Me.TxtTransID.text))
        Else
            rs("Transaction_ID").value = Null
        End If

        rs.update

        If CboMaintenanceType.ListIndex = 1 Then
            '«·⁄·«ﬁ… »Ì‰ ⁄„·Ì«  «·’Ì«‰… Ê›Ê« Ì— «·»Ì⁄
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open "MaintenanceJuncTransaction", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            RsTemp.AddNew
            RsTemp("JuncID").value = CStr(new_id("MaintenanceJuncTransaction", "JuncID", "", True))
            RsTemp("Transaction_ID").value = IIf(Me.TxtTransID.text = "", Null, val(Me.TxtTransID.text))
            RsTemp("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsTemp.update
        End If

        For RowNum = 1 To FG.Rows - 1
            RsDetails.AddNew
            RsDetails("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsDetails("ItemCode").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RsDetails("ItemID").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Name"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))

            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                RsCheckSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsCheckSerial.EOF Or RsCheckSerial.BOF) Then
                    If RsCheckSerial("HaveSerial").value = True Then
                        RsDetails("ItemSerial").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                    End If
                End If

                RsCheckSerial.Close
            End If

            RsDetails("Cost").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))), "", val(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))))
            RsDetails.update
        Next RowNum

        If Me.XPChkPayType(0).value = Checked Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("MaintananceID").value = val(XPTxtMaintanenceID.text)

            If Me.TxtModFlg.text = "N" Then
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(0).text = RsNotes("NoteSerial").value
            ElseIf Trim(XPTxtSerial(0).text) <> "" Then
                RsNotes("NoteSerial").value = Trim(XPTxtSerial(0).text)
            Else
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(0).text = RsNotes("NoteSerial").value
            End If

            RsNotes("NoteType").value = 0
            RsNotes("NoteDate").value = XPDtbGoInDtae.value
            RsNotes("Note_Value").value = IIf(XPTxtValue(0).text = "", Null, (XPTxtValue(0).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = Null
            RsNotes("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
            RsNotes("CusID").value = Null
            RsNotes.update
        End If

        If Me.XPChkPayType(1).value = Checked Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("MaintananceID").value = val(XPTxtMaintanenceID.text)

            If Me.TxtModFlg.text = "N" Then
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(1).text = RsNotes("NoteSerial").value
            ElseIf Trim(XPTxtSerial(1).text) <> "" Then
                RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
            Else
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(1).text = RsNotes("NoteSerial").value
            End If

            RsNotes("NoteType").value = 1
            RsNotes("NoteDate").value = XPDtbGoInDtae.value
            RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = Null
            RsNotes("CusID").value = Null
            RsNotes("DueDate").value = DtpDelayDate.value
            RsNotes.update
        End If

        If Me.XPChkPayType(2).value = Checked Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteDate").value = XPDtbGoInDtae.value
            RsNotes("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsNotes("NoteType").value = 2
            RsNotes("Note_Value").value = IIf(XPTxtValue(2).text = "", Null, val(XPTxtValue(2).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = IIf(DcboBankName.BoundText = "", Null, val(DcboBankName.BoundText))
            RsNotes("ChqueNum").value = IIf(XPTxtChqueNum.text = "", "", Trim(XPTxtChqueNum.text))
            RsNotes("DueDate").value = XPDTPDueDate.value
            RsNotes("CusID").value = Me.DBCboClientName.BoundText
            RsNotes.update
        End If

        'Õ›Ÿ »Ì«‰«  «·ﬁÿ⁄… «· Ì  „ «” »œ«·Â«
        If Me.Tag = "XX" Then

            ' «·‹ «ﬂœ √‰ »Ì«‰«  «·«” »œ«· ÂÌ ‰›” »Ì«‰«  ⁄„·Ì… «·’Ì«‰…
            If FrmReplace.TxtTransID.text <> TxtTransID.text Or FrmReplace.XPTxtMaintanenceID.text <> XPTxtMaintanenceID.text Or FrmReplace.DCboItemsName.BoundText <> FG.TextMatrix(FG.Tag, FG.ColIndex("Code")) Or FrmReplace.TxtItemSerial.text <> FG.TextMatrix(FG.Tag, FG.ColIndex("Serial")) Then
                Msg = "·ﬁœ  „  €ÌÌ— »⁄÷ «·»Ì«‰«  «· Ì  „  ”ÃÌ·Â« ›Ì ⁄„·Ì… «·«” »œ«·" & Chr(13)
                Msg = Msg + "·‰ Ì „  ”ÃÌ· ⁄„·Ì… «·«” »œ«·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                GoTo CompleteSaving
            End If
        
            Set RsReplace = New ADODB.Recordset
            RsReplace.Open "Transactions", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            '«·ﬁÿ⁄… «·„” »œ·…
            RsReplace.AddNew
            RsReplace("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
            ReplaceID = RsReplace("Transaction_ID").value
            RsReplace("StoreID").value = FrmReplace.DCboStoreName.BoundText
            RsReplace("CusID").value = IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
            RsReplace("Transaction_Type").value = 12
            RsReplace("MaintenanceID").value = Trim(XPTxtMaintanenceID.text)
            RsReplace("ReturnID").value = Trim(TxtTransID.text)
            RsReplace("Transaction_Date").value = FrmReplace.DtbReplaceDate.value
            
            RsReplace.update
        
            Set RsReplaceDetails = New ADODB.Recordset
            RsReplaceDetails.Open "Transaction_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            RsReplaceDetails.AddNew
            RsReplaceDetails("Transaction_ID").value = RsReplace("Transaction_ID").value
            RsReplaceDetails("Item_ID").value = FrmReplace.DCboItemsName.BoundText
            RsReplaceDetails("ItemSerial").value = Trim(FrmReplace.TxtItemSerial.text)
            RsReplaceDetails("Quantity").value = 1
            RsReplaceDetails.update
        
            '«·⁄·«ﬁ… »Ì‰ ⁄„·Ì«  «·’Ì«‰… Ê»Ì«‰«  «·ﬁÿ⁄ «·„” »œ·…
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open "MaintenanceJuncTransaction", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            RsTemp.AddNew
            RsTemp("JuncID").value = CStr(new_id("MaintenanceJuncTransaction", "JuncID", "", True))
            RsTemp("Transaction_ID").value = ReplaceID
            RsTemp("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsTemp.update
        
            '«·ﬁÿ⁄… «·ÃœÌœ…
            RsReplace.AddNew
            RsReplace("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
            RsReplace("StoreID").value = FrmReplace.DCboStoreName.BoundText
            RsReplace("CusID").value = IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
            RsReplace("Transaction_Type").value = 13
            '            RsReplace("MaintenanceID").Value = Trim(XPTxtMaintanenceID.Text)
            RsReplace("ReturnID").value = ReplaceID
            RsReplace("Transaction_Date").value = FrmReplace.DtbReplaceDate.value
            RsReplace.update
        
            '        Set RsReplaceDetails = New ADODB.Recordset
            '        RsReplaceDetails.Open "Transaction_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            RsReplaceDetails.AddNew
            RsReplaceDetails("Transaction_ID").value = RsReplace("Transaction_ID").value
            RsReplaceDetails("Item_ID").value = FrmReplace.DCboItemsName.BoundText
            RsReplaceDetails("ItemSerial").value = Trim(FrmReplace.TxtNewSerial.text)
            RsReplaceDetails("Quantity").value = 1
            RsReplaceDetails.update
        End If

CompleteSaving:
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
    End If

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Del_TransAction()
    Dim RsTemp As ADODB.Recordset
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If XPTxtMaintanenceID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtMaintanenceID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
            StrSQL = "select * From  Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "ÊÕ–› Â–Â «·⁄„·Ì… ”ÌƒœÌ ≈·Ï Õ–› »Ì«‰«  ⁄„·Ì… «·«” »œ«·" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…"

                If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
                    If Not rs.RecordCount < 1 Then
                        rs.delete
                        StrSQL = "delete From Transactions where MaintenanceID=" & val(XPTxtMaintanenceID.text)
                        Cn.Execute StrSQL, , adExecuteNoRecords
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

                If Not rs.RecordCount < 1 Then
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

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "MaintananceID='" & val(XPTxtMaintanenceID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ’Ì«‰… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·’Ì«‰…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ’Ì«‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdReplace, "«” »œ«· ..." & Wrap & "·«” »œ«· ﬁÿ⁄…  »⁄ «·÷„«‰" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsNotes As New ADODB.Recordset
    Dim RsDetails As New ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "MaintananceID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtMaintanenceID.text = IIf(IsNull(rs("MaintananceID").value), "", (rs("MaintananceID").value))
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    XPDtbGoInDtae.value = IIf(IsNull(rs("DateGoIN").value), Date, rs("DateGoIN").value)
    XPDtbGoOutDtae.value = IIf(IsNull(rs("DateGoOUT").value), Date, rs("DateGoOUT").value)
    XPChkGoOut.value = IIf(rs("GoOut").value = True, vbChecked, vbUnchecked)
    CboPaymentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    CboMaintenanceType.ListIndex = IIf(IsNull(rs("MType").value), 0, rs("MType").value)

    If Not IsNull(rs("Transaction_ID").value) Then
        Me.TxtTransID.text = rs("Transaction_ID").value
        'TxtTransSerial.Text = GetTransIDSerial(1, Val(Rs("Transaction_ID").Value))
    Else
        Me.TxtTransID.text = ""
    End If

    StrSQL = "select * From Notes where MaintananceID=" & val(XPTxtMaintanenceID.text)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtValue(2).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    XPTxtChqueNum.text = ""
    DcboBankName.BoundText = ""
    XPDTPDueDate.value = Date
    DtpDelayDate.value = Date

    If Not RsNotes.EOF Or RsNotes.BOF Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", Trim(RsNotes("BoxID").value))
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 2 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
                XPTxtValue(2).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtChqueNum.text = IIf(IsNull(RsNotes("ChqueNum").value), "", Trim(RsNotes("ChqueNum").value))
                Me.DcboBankName.BoundText = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                XPDTPDueDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            RsNotes.MoveNext

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT TblItems.HaveSerial,* FROM TblItems INNER JOIN TblMainteneceDetails " & "ON TblItems.ItemID = TblMainteneceDetails.ItemID"
    StrSQL = StrSQL + "  where MaintananceID=" & val(rs("MaintananceID").value)
    'StrSql = "select * From TblMainteneceDetails where MaintananceID=" & Val(Rs("MaintananceID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 0 To RsDetails.RecordCount - 1
            FG.Cell(flexcpPicture, Num + 1, FG.ColIndex("Replace")) = ""
            FG.Cell(flexcpData, Num + 1, FG.ColIndex("Replace")) = ""
            FG.TextMatrix(Num + 1, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("TblMainteneceDetails.ItemCode")), "", Trim(RsDetails("TblMainteneceDetails.ItemCode").value))
            FG.TextMatrix(Num + 1, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("TblMainteneceDetails.ItemID")), "", Trim(RsDetails("TblMainteneceDetails.ItemID").value))
            FG.TextMatrix(Num + 1, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial").value = True Then
                FG.TextMatrix(Num + 1, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «· Ì  „  ›Ì Â–Â «·⁄„·Ì…
                If (RsDetails("TblMainteneceDetails.ItemID")) <> "" And RsDetails("ItemSerial") <> "" Then
                    StrSQL = "select * From ReplacedItems where MaintenanceID=" & XPTxtMaintanenceID.text
                    StrSQL = StrSQL + " and ItemID=" & RsDetails("TblMainteneceDetails.ItemID")
                    StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        FG.Cell(flexcpPicture, Num + 1, FG.ColIndex("Replace")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                        FG.Cell(flexcpData, Num + 1, FG.ColIndex("Replace")) = "X"
                    End If
                End If
            End If

            FG.TextMatrix(Num + 1, FG.ColIndex("Cost")) = IIf(IsNull(RsDetails("Cost")), "", Trim(RsDetails("Cost").value))
            RsDetails.MoveNext
        Next Num

    End If

    XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTab301.CurrTab = 0 Then
            XPBtnAdd.Enabled = True
            XPBtnRemove.Enabled = True
        Else
            XPBtnAdd.Enabled = False
            XPBtnRemove.Enabled = False
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintingData()
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If XPTxtMaintanenceID.text <> "" Then
            Set MaintenReport = New ClsMaintananceReport
            MaintenReport.MaintenanceDataShort XPTxtMaintanenceID.text
        End If

    Else

        If XPTxtMaintanenceID.text <> "" Then
            Set MaintenReport = New ClsMaintananceReport
            MaintenReport.MaintenanceData XPTxtMaintanenceID.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
                StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)

            Case "E"
                StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
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
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPaymentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtSum_Change()

    If CboPaymentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

End Sub

Private Sub DBCboClientName_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                CboPaymentType.locked = True
                CboPaymentType.ListIndex = 0
            Else
                CboPaymentType.locked = False
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub FillItemData()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›
    StrSQL = "SELECT * From TblItems"
    fill_combo Me.DCboItemsCode, StrSQL
    '«”„ «·’‰›
    StrSQL = "SELECT * From TblItems"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then

        With DCboItemsName
            Set .RowSource = rs
            .BoundColumn = rs(0).name
            .ListField = rs(2).name
            .BoundText = ""
            .text = ""
        End With

    End If

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboItemsCode
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DCboItemsName
    ''Õ«·… «·’‰›
    'With CboItemCase
    '    .AddItem "ÃœÌœ"
    '    .AddItem "„” ⁄„·"
    'End With
    Exit Sub
ErrTrap:
End Sub

