VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManGoBack 
   Caption         =   "—ÃÊ⁄ «·÷„«‰ „‰ «·„Ê—œ"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   Icon            =   "FrmManGoBack0.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7905
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12765
      _cx             =   22516
      _cy             =   13944
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
      _GridInfo       =   $"FrmManGoBack0.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6900
         Width           =   12735
         _cx             =   22463
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
            Left            =   5940
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   945
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3780
            TabIndex        =   3
            Top             =   45
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·›« Ê—…"
            Height          =   255
            Index           =   0
            Left            =   6915
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   75
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   1
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   765
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   885
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1755
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   105
            Width           =   675
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   135
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   330
            Index           =   4
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   75
            Width           =   855
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1575
         Index           =   0
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   645
         Width           =   12735
         _cx             =   22463
         _cy             =   2778
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
         Begin VB.TextBox TxtOrgManSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   420
            Width           =   1560
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   195
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1170
            Width           =   4140
         End
         Begin VB.TextBox TxtOrgManID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   8040
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   60
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6630
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1170
            Width           =   4635
         End
         Begin VB.TextBox XPTxtMaintanenceID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9720
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   60
            Width           =   1560
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6105
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   90
            Visible         =   0   'False
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   6615
            TabIndex        =   13
            Top             =   810
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoInDtae 
            Height          =   315
            Left            =   195
            TabIndex        =   14
            Top             =   795
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            Format          =   95420417
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   195
            TabIndex        =   15
            Top             =   60
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   195
            TabIndex        =   16
            Top             =   435
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdSearchTrans 
            Height          =   315
            Left            =   8895
            TabIndex        =   62
            Top             =   450
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "..."
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
            ButtonImage     =   "FrmManGoBack0.frx":0623
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdOpenTrans 
            Height          =   345
            Left            =   8040
            TabIndex        =   63
            Top             =   450
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "..."
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
            ButtonImage     =   "FrmManGoBack0.frx":09BD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   315
            Index           =   18
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1200
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ≈Ì’«· «·œŒÊ·"
            Height          =   405
            Index           =   10
            Left            =   11325
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   300
            Index           =   16
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1230
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”·Ì„"
            Height          =   255
            Index           =   3
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   810
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   450
            Index           =   6
            Left            =   11430
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   870
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   285
            Index           =   8
            Left            =   11370
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   90
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   375
            Index           =   25
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   90
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   270
            Index           =   24
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   435
            Width           =   1260
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4605
         Left            =   15
         TabIndex        =   22
         Top             =   2280
         Width           =   12735
         _cx             =   22463
         _cy             =   8123
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
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…"
         Align           =   0
         CurrTab         =   1
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
         Picture(0)      =   "FrmManGoBack0.frx":0D57
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4140
            Index           =   2
            Left            =   -13290
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   12645
            _cx             =   22304
            _cy             =   7303
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
            GridRows        =   3
            GridCols        =   2
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmManGoBack0.frx":10F1
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1530
               Index           =   5
               Left            =   0
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   0
               Width           =   12645
               _cx             =   22304
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
               Begin MSDataListLib.DataCombo DCboSupDeci 
                  Height          =   315
                  Left            =   6555
                  TabIndex        =   87
                  Top             =   630
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtNewSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   1200
                  Width           =   3030
               End
               Begin VB.TextBox TxtReItemQty 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.TextBox TxtTicketNo 
                  Height          =   315
                  Left            =   9780
                  TabIndex        =   28
                  Top             =   285
                  Width           =   1620
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2010
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   285
                  Width           =   2730
               End
               Begin VB.TextBox TxtCost 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2010
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   630
                  Width           =   2730
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   285
                  Width           =   1740
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   4785
                  TabIndex        =   29
                  Top             =   285
                  Width           =   3195
                  _ExtentX        =   5636
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   8025
                  TabIndex        =   30
                  Top             =   285
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdShowTransItems 
                  Height          =   345
                  Left            =   11610
                  TabIndex        =   31
                  Top             =   210
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "...."
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
                  ButtonImage     =   "FrmManGoBack0.frx":1145
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   405
                  Left            =   60
                  TabIndex        =   75
                  Top             =   660
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmManGoBack0.frx":14DF
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   4210752
               End
               Begin MSDataListLib.DataCombo DcboReItemName 
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   79
                  Top             =   1200
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboReItemCode 
                  Height          =   315
                  Left            =   10560
                  TabIndex        =   80
                  Top             =   1200
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboReStore 
                  Height          =   315
                  Left            =   285
                  TabIndex        =   81
                  Top             =   1200
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«· «·ÃœÌœ"
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
                  Height          =   225
                  Index           =   27
                  Left            =   4785
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   960
                  Width           =   1860
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
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
                  Height          =   225
                  Index           =   26
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   960
                  Width           =   1470
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
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
                  Height          =   225
                  Index           =   19
                  Left            =   7395
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   960
                  Width           =   2865
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Œ“‰"
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
                  Height          =   225
                  Index           =   15
                  Left            =   870
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   960
                  Width           =   1830
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
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
                  Height          =   195
                  Index           =   5
                  Left            =   2835
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   990
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· ﬂ "
                  Height          =   195
                  Index           =   11
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   60
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   225
                  Index           =   31
                  Left            =   7935
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   30
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   225
                  Index           =   30
                  Left            =   4935
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   30
                  Width           =   2880
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   225
                  Index           =   28
                  Left            =   2205
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   30
                  Width           =   2310
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ê÷⁄"
                  Height          =   270
                  Index           =   23
                  Left            =   11445
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   660
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—ﬁ «Ê «· Œ’Ì„ "
                  Height          =   270
                  Index           =   7
                  Left            =   4785
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   690
                  Width           =   1665
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   195
                  Index           =   9
                  Left            =   195
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   30
                  Width           =   1410
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2205
               Left            =   0
               TabIndex        =   39
               Top             =   1530
               Width           =   12645
               _cx             =   22304
               _cy             =   3889
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmManGoBack0.frx":1DB9
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
               Left            =   465
               TabIndex        =   40
               Top             =   3735
               Width           =   12180
               _ExtentX        =   21484
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4140
            Index           =   6
            Left            =   45
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   45
            Width           =   12645
            _cx             =   22304
            _cy             =   7303
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
            GridRows        =   7
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmManGoBack0.frx":1F86
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄ «·√Ã·"
               Height          =   645
               Index           =   1
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   1665
               Width           =   6840
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   3120
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   5070
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   210
                  Width           =   1035
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   69
                  Top             =   150
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   635
                  _Version        =   393216
                  Format          =   95420417
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   210
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   17
                  Left            =   6240
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   270
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   525
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄ «·‰ﬁœÏ"
               Height          =   675
               Index           =   0
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   630
               Width           =   6840
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3090
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5040
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   210
                  Width           =   1155
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   56
                  Top             =   180
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
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   210
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6180
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   240
                  Width           =   465
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   630
               Index           =   20
               Left            =   11430
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   7350
         Width           =   12735
         _cx             =   22463
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
            Left            =   11325
            TabIndex        =   42
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9975
            TabIndex        =   43
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   8520
            TabIndex        =   44
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   7125
            TabIndex        =   45
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   5670
            TabIndex        =   46
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   4290
            TabIndex        =   47
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   30
            TabIndex        =   48
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   2790
            TabIndex        =   49
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   375
            Left            =   1455
            TabIndex        =   50
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   15
         Width           =   12735
         _cx             =   22463
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
         Caption         =   "—ÃÊ⁄ «·÷„«‰ „‰ «·„Ê—œ"
         Align           =   0
         AutoSizeChildren=   7
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1635
            TabIndex        =   89
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "FrmManGoBack0.frx":202C
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
            Left            =   945
            TabIndex        =   90
            Top             =   120
            Width           =   690
            _ExtentX        =   1217
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
            ButtonImage     =   "FrmManGoBack0.frx":23C6
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
            Left            =   2550
            TabIndex        =   91
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
            ButtonImage     =   "FrmManGoBack0.frx":2760
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
            Left            =   105
            TabIndex        =   92
            Top             =   120
            Width           =   780
            _ExtentX        =   1376
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
            ButtonImage     =   "FrmManGoBack0.frx":2AFA
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
   End
End
Attribute VB_Name = "FrmManGoBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim MaintenReport As ClsMaintananceReport
Dim cSearchDcbo(4) As clsDCboSearch

Public BolPrint As Boolean

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    XPTab301.TabCaption(1) = "Notes"
    Fram(0).Caption = "Cash"
    Fram(1).Caption = "Credit"
    lbl(13).Caption = "Value"
    lbl(17).Caption = "Value"
    lbl(22).Caption = "Box"
    lbl(21).Caption = "Due Date"
    lbl(1).Caption = "Curr. Rec"

    With Fg
        .TextMatrix(0, .ColIndex("Code")) = "code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("Count")) = "Qty"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "HaveSerial"
        .TextMatrix(0, .ColIndex("Serial")) = "serial"
        .TextMatrix(0, .ColIndex("TicketNO")) = "TicketNO"
        ' .TextMatrix(0, .ColIndex("CusNotes")) = "Cus. Complaint"
        .TextMatrix(0, .ColIndex("TicketNO")) = "TicketNO"
        .TextMatrix(0, .ColIndex("SupDeci")) = "Supplier Decision"
        .TextMatrix(0, .ColIndex("NewSerial")) = "New Serial"
        .TextMatrix(0, .ColIndex("cost")) = "Cost"

    End With

    'Me.Caption = "Maintenance Delivery"
    Me.Caption = "Returned Parts From Supplier"
    C1Elastic6.Caption = Me.Caption
    lbl(8).Caption = "Opr#"
    lbl(10).Caption = "Recive No."

    lbl(6).Caption = "Customer"
    lbl(16).Caption = "Cash Customer Name"

    lbl(25).Caption = " Employee"
    lbl(24).Caption = "Store Name"
 
    lbl(3).Caption = "Delivery Date"
    lbl(18).Caption = "Payment Type"
    lbl(11).Caption = "Ticket NO"
 
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(28).Caption = "Item Serial"
    lbl(9).Caption = "Quantity"

    lbl(23).Caption = "Status"
    lbl(7).Caption = "Different"
 
    'ChkFastReplace.Caption = "Immediate replacement"
    lbl(26).Caption = "Item Code"
    lbl(19).Caption = "Item Name"
    lbl(27).Caption = "New Serial"
    lbl(5).Caption = "Quantity"
    lbl(15).Caption = "Store Name"
    lbl(0).Caption = "Totals"

    lbl(4).Caption = "BY Employee"
  
    lbl(2).Caption = "Curr. Rec."
    'lbl(2).Caption = "Rec. Count:"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.XPTab301.TabCaption(0) = "Items"
    
End Sub

Private Sub DCboSupDeci_Change()
    'If Me.DCboSupDeci.BoundText = -1 Then
    '    Exit Sub
    'ElseIf Me.DCboSupDeci.BoundText = 0 Then
    '    Me.lbl(5).Enabled = False
    '    Me.TxtNewSerial.Enabled = False
    '    Me.lbl(7).Enabled = False
    '    Me.TxtCost.Enabled = False
    'ElseIf Me.DCboSupDeci.BoundText = 1 Then
    '    Me.lbl(5).Enabled = False
    '    Me.TxtNewSerial.Enabled = False
    '    Me.lbl(7).Enabled = False
    '    Me.TxtCost.Enabled = False
    'ElseIf Me.DCboSupDeci.BoundText = 2 Then
    '    Me.lbl(5).Enabled = True
    '    Me.TxtNewSerial.Enabled = True
    '    Me.lbl(7).Enabled = False
    '    Me.TxtCost.Enabled = False
    'ElseIf Me.DCboSupDeci.BoundText = 3 Then
    '    Me.lbl(5).Enabled = True
    '    Me.TxtNewSerial.Enabled = True
    '    Me.lbl(7).Enabled = True
    '    Me.TxtCost.Enabled = True
    'ElseIf Me.DCboSupDeci.BoundText = 4 Then
    '    Me.lbl(5).Enabled = True
    '    Me.TxtNewSerial.Enabled = True
    '    Me.lbl(7).Enabled = True
    '    Me.TxtCost.Enabled = True
    'ElseIf Me.DCboSupDeci.BoundText = 5 Then
    '    Me.lbl(5).Enabled = False
    '    Me.TxtNewSerial.Enabled = False
    '    Me.lbl(7).Enabled = True
    '    Me.TxtCost.Enabled = True
    'ElseIf Me.DCboSupDeci.BoundText = 6 Then
    '    Me.lbl(5).Enabled = False
    '    Me.TxtNewSerial.Enabled = False
    '    Me.lbl(7).Enabled = False
    '    Me.TxtCost.Enabled = False
    'ElseIf Me.DCboSupDeci.BoundText = 7 Then
    '    Me.lbl(5).Enabled = False
    '    Me.TxtNewSerial.Enabled = False
    '    Me.lbl(7).Enabled = True
    '    Me.TxtCost.Enabled = True
    'End If
End Sub

Private Sub CmdAdd_Click()
    '“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·

    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngFindRow As Long
    Dim LngRow As Long
    Dim LngTransID As Long

    On Error GoTo ErrTrap

    If DCboItemsCode.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemsCode.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If val(TxtQuantity.text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtQuantity.SetFocus
        Exit Sub
    End If

    If Me.DCboSupDeci.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— ﬁ—«— «·„Ê—œ »‘√‰ Â–Â «·ﬁÿ⁄… „‰ «·’‰› ....!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboSupDeci.SetFocus
        Exit Sub
    ElseIf Me.DCboSupDeci.BoundText = 0 Then
    
    ElseIf Me.DCboSupDeci.BoundText = 1 Then
    
    ElseIf Me.DCboSupDeci.BoundText = 2 Then
        '    Me.lbl(5).Enabled = True
        '    Me.TxtNewSerial.Enabled = True
        '    Me.lbl(7).Enabled = False
        '    Me.TxtCost.Enabled = False
    ElseIf Me.DCboSupDeci.BoundText = 3 Then
        '    Me.lbl(5).Enabled = True
        '    Me.TxtNewSerial.Enabled = True
        '    Me.lbl(7).Enabled = True
        '    Me.TxtCost.Enabled = True
    ElseIf Me.DCboSupDeci.BoundText = 4 Then
        '    Me.lbl(5).Enabled = False
        '    Me.TxtNewSerial.Enabled = False
        '    Me.lbl(7).Enabled = True
        '    Me.TxtCost.Enabled = True
    ElseIf Me.DCboSupDeci.BoundText = 5 Then
        '    Me.lbl(5).Enabled = False
        '    Me.TxtNewSerial.Enabled = False
        '    Me.lbl(7).Enabled = True
        '    Me.TxtCost.Enabled = True
    End If

    If Me.TxtNewSerial.Enabled = True And Trim(Me.TxtNewSerial.text) = "" Then
        Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·ﬁÿ⁄… «·ÃœÌœ… „‰ «·’‰› ...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerial.SetFocus
        Exit Sub
    End If

    If Me.TxtCost.Enabled = True And val(Me.TxtCost.text) = 0 Then
        Msg = "»—Ã«¡ ≈œŒ«· ›—ﬁ «·”⁄— «Ê «· Œ’Ì„...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerial.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "N" Then
        LngTransID = 0
    ElseIf Me.TxtModFlg.text = "E" Then
        LngTransID = val(Me.XPTxtMaintanenceID.text)
    End If

    StrSQL = "SELECT QryManStockComplete.* "
    StrSQL = StrSQL + " FROM dbo.QryManStockComplete(" & LngTransID & ") QryManStockComplete"
    StrSQL = StrSQL + " Where ItemID=" & Me.DCboItemsCode.BoundText

    If Trim$(Me.TxtSerial.text) <> "" Then
        StrSQL = StrSQL + " AND ItemSerial='" & Trim$(Me.TxtSerial.text) & "'"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Msg = "Â–Â «·ﬁÿ⁄… €Ì— „ÊÃÊœ… ·œÏ «·„Ê—œ «·„Õœœ,,,"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Trim$(Me.TxtSerial.text) <> "" Then
        LngFindRow = Fg.FindRow(Trim$(Me.TxtSerial.text), Fg.FixedRows, Fg.ColIndex("Serial"), False, True)
    End If

    With Fg

        If LngFindRow <= 0 Then
            If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                .Rows = .Rows + 1
            End If

            LngRow = .Rows - 1
        Else
            LngRow = LngFindRow
        End If

        .TextMatrix(LngRow, .ColIndex("Name")) = DCboItemsName.BoundText
        .TextMatrix(LngRow, .ColIndex("Code")) = DCboItemsName.BoundText
        .TextMatrix(LngRow, .ColIndex("Serial")) = TxtSerial.text
        .TextMatrix(LngRow, .ColIndex("Count")) = TxtQuantity.text
        .TextMatrix(LngRow, .ColIndex("TicketNO")) = TxtTicketNO.text
        .TextMatrix(LngRow, .ColIndex("SupDeci")) = Me.DCboSupDeci.BoundText + 1
        .TextMatrix(LngRow, .ColIndex("NewSerial")) = Trim$(Me.TxtNewSerial.text)
        .TextMatrix(LngRow, .ColIndex("Cost")) = val(Me.TxtCost.text)

        If TxtSerial.Tag = "T" Then
            .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexChecked
        ElseIf TxtSerial.Tag = "F" Then
            .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexUnchecked
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    DCboItemsCode.BoundText = ""
    DCboItemsName.BoundText = ""
    TxtSerial.text = ""

    With Me.Fg
        XPTxtSum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Cost"), .Rows - 1, .ColIndex("Cost"))
        Me.XPTxtValue(0).text = XPTxtSum.text
    End With

    'XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
    Fg.SetFocus
    '-------------------------------------------
    'StrSQL = ""
    'StrSQL = "select * From QryMaintananceReport where ItemID=" & _
    '        Trim(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
    'StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
    'StrSQL = StrSQL + " and GoOut=False"
    'Set RsTemp = New ADODB.Recordset
    'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Not (RsTemp.EOF Or RsTemp.BOF) Then
    '    If RsTemp("MaintananceID").Value <> XPTxtMaintanenceID.text Then
    '        Msg = " „ «” ·«„ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
    '        Msg = Msg + "–«  «·”Ì—Ì«·        : " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
    '        Msg = Msg + "„‰ «·⁄„Ì·            :  " & RsTemp("CusName").Value & Chr(13)
    '        Msg = Msg + "» «—ÌŒ                   :  " & RsTemp("DateGoIN").Value & Chr(13)
    '        Msg = Msg + "·Ì „ ≈Ã—«¡ ⁄„·Ì… ’Ì«‰… ·Â« Ê·„ Ì „  ”·Ì„Â« »⁄œ" & Chr(13)
    '        Msg = Msg + " —ﬁ„ «·⁄„·Ì…         : " & RsTemp("MaintananceID").Value
    '        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        XPTab301.CurrTab = 0
    '        FG.Row = RowNum
    '        FG.Col = FG.ColIndex("Name")
    '        FG.ShowCell RowNum, FG.ColIndex("Name")
    '        FG.SetFocus
    '        Exit Sub
    '    End If
    'End If
    'If (DateDiff("d", XPDtbGoInDtae.Value, DateAdd("m", RsSerial("guaranteeTime").Value, RsSerial("Transaction_Date").Value))) < 0 Then
    '    Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
    '    Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").Value & Chr(13)
    '    Msg = Msg + "–«  «·”Ì—Ì«·  " & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & Chr(13)
    '    Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").Value, "yyyy/m/d") & Chr(13)
    '    Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").Value & Chr(13)
    '    Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").Value & "  ‘Â—" & Chr(13)
    '    Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"
    '    If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
    '        XPTab301.CurrTab = 0
    '        FG.Row = RowNum
    '        FG.Col = FG.ColIndex("Name")
    '        FG.ShowCell RowNum, FG.ColIndex("Name")
    '        FG.SetFocus
    '        Exit Sub
    '    End If
    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdOpenTrans_Click()

    If val(Me.TxtOrgManID.text) <> 0 Then
        'Load FrmMaintenence
        'FrmMaintenence.Retrive val(Me.TxtOrgManID.text)
        'FrmMaintenence.show
        'FrmMaintenence.ZOrder 0
    End If

End Sub

Private Sub CmdSearchTrans_Click()
    Load FrmMaintanenceSearch
    FrmMaintanenceSearch.searchtype = 1
    Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtOrgManID
    FrmMaintanenceSearch.show vbModal
End Sub

Private Sub CmdShowTransItems_Click()
    Dim Msg As String

    If val(Me.TxtOrgManID.text) = 0 Then
        Msg = "ÌÃ» ≈Œ Ì«— ≈Ì’«· œŒÊ· «·’Ì«‰…...!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtOrgManSerial.SetFocus
        Exit Sub
    End If

    If Me.DCboStoreName.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «”„ «·„Œ“‰ √Ê·« ....!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Load FrmManChooseItems
    Set FrmManChooseItems.MyForm = Me
    FrmManChooseItems.ShowManTrans val(Me.TxtOrgManID.text)
    FrmManChooseItems.show
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
            TxtQuantity.Enabled = False
            TxtQuantity.text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtSerial.Enabled = False
            TxtQuantity.Enabled = True
            TxtQuantity.text = ""
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

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 7

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMaximized
            Else
                Me.WindowState = vbNormal
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    'On Error GoTo ErrTrap
    'Dim RsSerial As New ADODB.Recordset
    'Dim RsTemp As ADODB.Recordset
    'Dim Msg As String
    'Dim StrSQL As String
    'If XPDtbGoInDtae.Value = "" Then
    '    Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ⁄„·Ì… «·’Ì«‰…" & Chr(13)
    '    Msg = Msg + "ﬁ»· ≈œŒ«· »Ì«‰«  «·√’‰«›"
    '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    XPDtbGoInDtae.SetFocus
    '    Exit Sub
    'End If
    'If Col = Fg.ColIndex("Name") Then
    '    If Fg.TextMatrix(Row, Fg.ColIndex("Name")) <> "" Then
    '        Fg.TextMatrix(Row, Fg.ColIndex("Code")) = Fg.TextMatrix(Row, Fg.ColIndex("Name"))
    '        If IsNumeric(Fg.TextMatrix(Row, Fg.ColIndex("Code"))) Then
    '            StrSQL = "select * From TblItems where ItemID=" & _
    '            Fg.TextMatrix(Row, Fg.ColIndex("Code"))
    '            Set RsTemp = New ADODB.Recordset
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '            If RsTemp.EOF Or RsTemp.BOF Then
    '                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Sub
    '            Else
    '                If RsTemp("HaveSerial").Value = True Then
    '                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = True
    '                Else
    '                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = False
    '                End If
    '            End If
    '        Else
    '            Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    '    End If
    'End If
    'If Col = Fg.ColIndex("Code") Then
    '    If Fg.TextMatrix(Row, Fg.ColIndex("Code")) <> "" Then
    '        Fg.TextMatrix(Row, Fg.ColIndex("Name")) = Fg.TextMatrix(Row, Fg.ColIndex("Code"))
    '        StrSQL = "select * From TblItems where ItemID=" & _
    '        Fg.TextMatrix(Row, Fg.ColIndex("Code")) & ""
    '        If IsNumeric(Fg.TextMatrix(Row, Fg.ColIndex("Code"))) Then
    '            Set RsTemp = New ADODB.Recordset
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '            If RsTemp.EOF Or RsTemp.BOF Then
    '                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                Exit Sub
    '            Else
    '                If RsTemp("HaveSerial").Value = True Then
    '                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = True
    '                Else
    '                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = False
    '                End If
    '            End If
    '        Else
    '            Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    '    End If
    'End If
    'If CboMaintenanceType.ListIndex = 1 Then
    '    If Fg.TextMatrix(Row, Fg.ColIndex("Code")) <> "" Then
    '        If Fg.Cell(flexcpChecked, Row, Fg.ColIndex("HaveSerial")) = flexChecked Then
    '            If Fg.TextMatrix(Row, Fg.ColIndex("Serial")) <> "" Then
    '                StrSQL = "select * From QryGuarantee where Item_ID=" & _
    '                Fg.TextMatrix(Row, Fg.ColIndex("Code")) & _
    '                " and ItemSerial='" & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & "'"
    '                StrSQL = StrSQL + " AND Transaction_Serial='" & Val(TxtTransSerial.text) & "'"
    '                StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
    '                RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '                If RsSerial.EOF Or RsSerial.BOF Then
    '                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
    '                    Msg = Msg + Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
    '                    Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
    '                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
    '                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '
    '                    '»Ì«‰«  «·›« Ê—… «· Ì  „ »Ì⁄ «·ﬁÿ⁄Â ›ÌÂ«
    '                    StrSQL = "select * From QryGuarantee where Item_ID=" & _
    '                    Fg.TextMatrix(Row, Fg.ColIndex("Code")) & _
    '                    " and ItemSerial='" & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & "'"
    '                    Set RsTemp = New ADODB.Recordset
    '                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
    '                        Msg = "·ﬁœ  „ »Ì⁄ «·ﬁÿ⁄… : " & Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")) & Chr(13)
    '                        Msg = Msg + "–«  «·”Ì—Ì«· : " & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
    '                        Msg = Msg + "≈·Ï «·⁄„Ì· : " & RsTemp("CusName").Value & Chr(13)
    '                        Msg = Msg + "›Ì «·›« Ê—… —ﬁ„ : " & RsTemp("Transaction_ID").Value
    '                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                    End If
    '                    XPTab301.CurrTab = 0
    '                    Fg.Row = Row
    '                    Fg.Col = Fg.ColIndex("Name")
    '                    Fg.ShowCell Row, Fg.ColIndex("Name")
    '                    Fg.SetFocus
    '                    Exit Sub
    '                End If
    '                If IsNull(RsSerial("guaranteeTime").Value) Then
    '                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
    '                    Msg = Msg + Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
    '                    Msg = Msg + "·Ì” ·Â« ÷„«‰"
    '                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '                    XPTab301.CurrTab = 0
    '                    Fg.Row = Row
    '                    Fg.Col = Fg.ColIndex("Name")
    '                    Fg.ShowCell Row, Fg.ColIndex("Name")
    '                    Fg.SetFocus
    '                    Exit Sub
    '                End If
    '                If (DateDiff("d", XPDtbGoInDtae.Value, DateAdd("m", RsSerial("guaranteeTime").Value, RsSerial("Transaction_Date").Value))) < 0 Then
    '                    Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
    '                    Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").Value & Chr(13)
    '                    Msg = Msg + "–«  «·”Ì—Ì«·  " & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
    '                    Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").Value, "yyyy/m/d") & Chr(13)
    '                    Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").Value & Chr(13)
    '                    Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").Value & "  ‘Â—" & Chr(13)
    '                    Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"
    '                    If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
    '                        XPTab301.CurrTab = 0
    '                        Fg.Row = Row
    '                        Fg.Col = Fg.ColIndex("Name")
    '                        Fg.ShowCell Row, Fg.ColIndex("Name")
    '                        Fg.SetFocus
    '                        Exit Sub
    '                    End If
    '                End If
    '                RsSerial.Close
    '            End If
    '        End If
    '    End If
    'End If
    'XPTxtSum.text = Fg.Aggregate(flexSTSum, 1, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
    'Exit Sub
    'ErrTrap:
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)
    On Error GoTo ErrTrap

    If Col = Fg.ColIndex("HaveSerial") Then
        Cancel = True
    End If

    With Fg

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
'        FrmAddNewItem.DealingForm = Maintenance
'        FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
    '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«·
    Dim StrSQL As String
    Dim RsReplace As ADODB.Recordset

    With Fg

        If .Col = -1 Then Exit Sub
        If .Row <= 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("Name")) <> "" Then
            Me.DCboItemsCode.BoundText = .TextMatrix(.Row, .ColIndex("Name"))
            Me.DCboItemsName.BoundText = .TextMatrix(.Row, .ColIndex("Name"))
            Me.DCboSupDeci.BoundText = val(.TextMatrix(.Row, .ColIndex("SupDeci"))) - 1
            Me.TxtNewSerial.text = Trim$(.TextMatrix(.Row, .ColIndex("NewSerial")))
            Me.TxtCost.text = val(.TextMatrix(.Row, .ColIndex("Cost")))

            If .Cell(flexcpChecked, .Row, .ColIndex("HaveSerial")) = flexChecked Then
                Me.TxtSerial.Enabled = True
                Me.TxtQuantity.Enabled = False
                Me.TxtQuantity.text = 1
                Me.TxtSerial.text = .TextMatrix(.Row, .ColIndex("Serial"))
            Else
                Me.TxtSerial.Enabled = False
                Me.TxtQuantity.Enabled = True
                Me.TxtQuantity.text = .TextMatrix(.Row, .ColIndex("Count"))
                Me.TxtSerial.text = ""
            End If
        
            Me.TxtTicketNO.text = .TextMatrix(.Row, .ColIndex("TicketNO"))
        
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
    Dim RsTemp As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

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
    Me.Height = 8580
    Me.Width = 9700
    Resize_Form Me
    'AddTip
    SetDtpickerDate Me.XPDtbGoInDtae
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox

    If SystemOptions.UserInterface = ArabicInterface Then

        With CboPaymentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

    Else

        With CboPaymentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With

    End If

    Dcombos.GetManSupDecs Me.DCboSupDeci
    Fg.WallPaper = BGround.Picture

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboEmp

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    SetDtpickerDate DtpDelayDate
    LoadTBR
    Set rs = New ADODB.Recordset
    rs.Open "Select * From  TblMaintenece Where ManOperationTypeID=4", Cn, adOpenStatic, adLockOptimistic, adCmdText

    StrSQL = "Select * From TblItems"
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    StrList = Fg.BuildComboList(RsItems, "ItemName", "ItemID")

    If StrList <> "" Then
        Fg.ColComboList(Fg.ColIndex("Name")) = "|" & StrList
    End If

    StrList = Fg.BuildComboList(RsItems, "ItemCode", "ItemID")

    If StrList <> "" Then
        Fg.ColComboList(Fg.ColIndex("Code")) = "|" & StrList
    End If

    StrList = ""
    StrSQL = "SELECT SupDecID,SupDecName From TblManSupDecs "
    StrSQL = StrSQL + " Order By SupDecID"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsTemp.BOF Or RsTemp.EOF Then
        StrList = Fg.BuildComboList(RsTemp, "SupDecName", "SupDecID")
    End If

    If StrList <> "" Then
        Fg.ColComboList(Fg.ColIndex("SupDeci")) = "|" & StrList
    End If

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

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    With Button

        Select Case .key

            Case "RemoveRow"

                If Fg.Rows > 1 Then
                    If Fg.Rows = 2 Then
                        Me.Fg.Clear flexClearScrollable, flexClearEverything
                    Else

                        If Me.Fg.Rows > 1 Then
                            If Me.Fg.Row <> Me.Fg.FixedRows - 1 Then
                                Me.Fg.RemoveItem (Me.Fg.Row)
                            End If
                        End If
                    End If
                End If

        End Select

    End With

End Sub

Private Sub TxtModFlg_Change()

    Select Case Me.TxtModFlg.text

        Case "R"
    
            '     Me.Caption = " ”·Ì„ «·’Ì«‰… ··⁄„Ì·"
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
            DBCboClientName.locked = True
            Fg.Editable = flexEDNone

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
       
            Ele(5).Enabled = False
            Me.DcboEmp.locked = True
            Me.DCboStoreName.locked = True
        
            Me.DBCboClientName.Enabled = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdSearchTrans.Enabled = False

        Case "N"
            '     Me.Caption = " ”·Ì„ «·’Ì«‰… ··⁄„Ì·( ÃœÌœ )"
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
       
            DBCboClientName.locked = False

            Fg.Enabled = True
            Fg.Rows = Fg.FixedRows
            Fg.Rows = 2
            '        FG.RowPosition(FG.Rows - 2) = FG.Rows - 1
            '        FG.TextMatrix(FG.Rows - 1, 2) = "«÷€ÿ Â‰«"
            Me.DBCboClientName.locked = False
            Fg.Editable = flexEDNone
            XPDtbGoInDtae.value = Date '
        
            Ele(5).Enabled = True
            Me.DcboEmp.locked = False
            Me.DCboStoreName.locked = False
            Me.DBCboClientName.Enabled = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdSearchTrans.Enabled = True

        Case "E"
            '     Me.Caption = " ”·Ì„ «·’Ì«‰… ··⁄„Ì·(  ⁄œÌ· )"
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
        
            DBCboClientName.locked = False
        
            Fg.Enabled = True
            Me.DBCboClientName.locked = False
        
            Fg.Editable = flexEDNone
            DBCboClientName_Change
        
            Ele(5).Enabled = True
            Me.DcboEmp.locked = False
            Me.DCboStoreName.locked = False
            Me.DBCboClientName.Enabled = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdSearchTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtOrgManID_Change()
    Dim StrTemp As String
    Dim StrCashCusName As String
    Dim LngCusID As Long

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.TxtOrgManID.text <> "" Then
            StrTemp = GetManIDSerial(1, val(Me.TxtOrgManID.text), 1, , LngCusID, StrCashCusName)

            If StrTemp <> Me.TxtOrgManSerial.text Then
                Me.TxtOrgManSerial.text = StrTemp
            End If

            If val(Me.DBCboClientName.BoundText) <> LngCusID Then
                Me.DBCboClientName.BoundText = LngCusID
            End If

            If Trim$(Me.TxtCashCustomerName.text) <> StrCashCusName Then
                Me.TxtCashCustomerName.text = StrCashCusName
            End If
        End If
    End If

End Sub

Private Sub TxtOrgManSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOrgManSerial.text, 1)
End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.text, 1)
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
            XPTxtMaintanenceID.text = CStr(new_id("TblMaintenece", "MaintananceID", "", True))
            XPTab301.CurrTab = 0
            Fg.SetFocus
            Fg.Col = Fg.ColIndex("Code")
            Fg.Row = Fg.Rows - 1

        Case 1
            '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
            StrSQL = "select * From  Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« «Â«"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
            Load FrmMaintanenceSearch
            FrmMaintanenceSearch.searchtype = 4
            FrmMaintanenceSearch.show vbModal

        Case 7
            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
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

Private Sub SaveData()
    Dim RsNotes As ADODB.Recordset
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
    Dim note_id As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If DBCboClientName.text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.DcboEmp.BoundText = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„ÊŸ›...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmp.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    
        If Me.DCboStoreName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰....!!! " & Chr(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If CboPaymentType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPaymentType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If ItemsInGrid(Fg, Fg.ColIndex("Name")) = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— «·√’‰«› «· Ï ”Ê›  ”·„ ··⁄„·...!!! " & Chr(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        If val(Me.XPTxtValue(0).text) > 0 Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "»—Ã«¡  ÕœÌœ «”„ «·Œ“‰… ..!!! " & Chr(13)
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
        End If

        If (val(Me.XPTxtValue(0).text) + val(Me.XPTxtValue(1).text)) <> val(Me.XPTxtSum.text) Then
            Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ﬂ«›Ï ·ﬁÌ„… «·›« Ê—…" & Chr(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        If Me.TxtModFlg.text = "N" Then
            Me.XPTxtMaintanenceID.text = CStr(new_id("TblMaintenece", "MaintananceID", "", True))
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "delete From TblMainteneceDetails where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "delete From MaintenanceJuncTransaction where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "delete From Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "delete From Notes where MaintananceID=" & val(rs("MaintananceID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        RsDetails.Open "[TblMainteneceDetails]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        Cn.BeginTrans
        BeginTrans = True
        rs("MaintananceID").value = val(XPTxtMaintanenceID.text)
        rs("CusID").value = Me.DBCboClientName.BoundText ' IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
        rs("DateGoIN").value = XPDtbGoInDtae.value
        rs("DateGoOUT").value = Null
        rs("GoOut").value = 0
        rs("EmpID").value = Me.DcboEmp.BoundText
        rs("StoreID").value = Me.DCboStoreName.BoundText
        rs("UserID").value = user_id

        If CboPaymentType.ListIndex = -1 Then
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPaymentType.ListIndex)
        End If

        rs("MType").value = 0
        rs("Transaction_ID").value = Null
        rs("ManOperationTypeID").value = 4
    
        rs.update

        For RowNum = 1 To Fg.Rows - 1
            RsDetails.AddNew
            RsDetails("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsDetails("ItemID").value = IIf(IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))), "", Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))))

            If Not Fg.TextMatrix(RowNum, Fg.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))
                RsCheckSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsCheckSerial.EOF Or RsCheckSerial.BOF) Then
                    If RsCheckSerial("HaveSerial").value = True Then
                        RsDetails("ItemSerial").value = IIf(IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))), "", Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))))
                    End If
                End If

                RsCheckSerial.Close
            End If

            RsDetails("Quantity").value = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))
            RsDetails("TicketNO").value = Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("TicketNO")))
            RsDetails("CustomerNotes").value = Null
            RsDetails("EmpNotes").value = Null
            RsDetails("Cost").value = val(Me.Fg.TextMatrix(RowNum, Fg.ColIndex("Cost")))
            RsDetails("SupDeci").value = val(Me.Fg.TextMatrix(RowNum, Fg.ColIndex("SupDeci")))
            RsDetails("NewSerial").value = val(Me.Fg.TextMatrix(RowNum, Fg.ColIndex("NewSerial")))
            RsDetails.update
        Next RowNum

        Set RsNotes = New ADODB.Recordset
        StrSQL = "Notes"
        RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdTable

        If val(XPTxtValue(0).text) > 0 Then
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
            RsNotes("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes.update
        End If

        If val(XPTxtValue(1).text) > 0 Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
            note_id = RsNotes("NoteID").value
            RsNotes("NoteDate").value = XPDtbGoInDtae.value

            If Me.TxtModFlg.text = "N" Then
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(1).text = RsNotes("NoteSerial").value
            ElseIf Trim(XPTxtSerial(1).text) <> "" Then
                RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
            Else
                RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                XPTxtSerial(1).text = RsNotes("NoteSerial").value
            End If

            RsNotes("Transaction_ID").value = Null
            RsNotes("MaintananceID").value = val(XPTxtMaintanenceID.text)
            RsNotes("NoteType").value = 1
            RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = Null
            RsNotes("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("DueDate").value = DtpDelayDate.value
            RsNotes.update
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

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
            StrSQL = "select * From  Transactions where MaintenanceID=" & val(rs("MaintananceID").value)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "ÊÕ–› Â–Â «·⁄„·Ì… ”ÌƒœÌ ≈·Ï Õ–› »Ì«‰«  ⁄„·Ì… «·«” »œ«·" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…"

                If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
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

    'With TTP
    '   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl CmdReplace, _
    '    "«” »œ«· ..." & Wrap & _
    '    "·«” »œ«· ﬁÿ⁄…  »⁄ «·÷„«‰" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
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

    Dim RsDetails As New ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim Num As Integer

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
    Me.TxtCashCustomerName.text = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    XPDtbGoInDtae.value = IIf(IsNull(rs("DateGoIN").value), Date, rs("DateGoIN").value)
    CboPaymentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)

    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT TblItems.HaveSerial,* FROM TblItems INNER JOIN TblMainteneceDetails " & "ON TblItems.ItemID = TblMainteneceDetails.ItemID"
    StrSQL = StrSQL + "  where MaintananceID=" & val(rs("MaintananceID").value)
    'StrSql = "select * From TblMainteneceDetails where MaintananceID=" & Val(Rs("MaintananceID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = RsDetails.RecordCount + 1

        For Num = 0 To RsDetails.RecordCount - 1
            Fg.Cell(flexcpPicture, Num + 1, Fg.ColIndex("Replace")) = ""
            Fg.Cell(flexcpData, Num + 1, Fg.ColIndex("Replace")) = ""
            Fg.TextMatrix(Num + 1, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial").value = True Then
                Fg.TextMatrix(Num + 1, Fg.ColIndex("HaveSerial")) = True
                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «· Ì  „  ›Ì Â–Â «·⁄„·Ì…
                '            If (RsDetails("ItemID")) <> "" And RsDetails("ItemSerial") <> "" Then
                '                StrSQL = "select * From ReplacedItems where MaintenanceID=" & XPTxtMaintanenceID.text
                '                StrSQL = StrSQL + " and ItemID=" & RsDetails("TblMainteneceDetails.ItemID")
                '                StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                '                Set RsReplace = New ADODB.Recordset
                '                RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '                If Not (RsReplace.EOF Or RsReplace.BOF) Then
                '                    Fg.Cell(flexcpPicture, Num + 1, Fg.ColIndex("Replace")) = MDIFrmMain.ImgLstTree.ListImages("Request").Picture
                '                    Fg.Cell(flexcpData, Num + 1, Fg.ColIndex("Replace")) = "X"
                '                End If
                '            End If
            End If

            '
            Fg.TextMatrix(Num + 1, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", Trim(RsDetails("Quantity").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("SupDeci")) = IIf(IsNull(RsDetails("SupDeci")), "", val(RsDetails("SupDeci").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("NewSerial")) = IIf(IsNull(RsDetails("NewSerial")), "", val(RsDetails("NewSerial").value))
            Fg.TextMatrix(Num + 1, Fg.ColIndex("Cost")) = IIf(IsNull(RsDetails("Cost")), "", val(RsDetails("Cost").value))
            RsDetails.MoveNext
        Next Num

        Me.XPTxtSum.text = Fg.Aggregate(flexSTSum, Fg.FixedRows, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
        Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "Select * From Notes where MaintananceID=" & val(XPTxtMaintanenceID.text)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    DtpDelayDate.value = Date

    If Not (rs.BOF Or rs.EOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", Trim(RsNotes("BoxID").value))
            End If

            If RsNotes("NoteType").value = 1 Then
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            RsNotes.MoveNext

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
            End If

        Next Num

    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub

    'XPChkPayType(0).Value = Unchecked
    'XPChkPayType(1).Value = Unchecked
    'XPChkPayType(2).Value = Unchecked

    'XPTxtValue(2).text = ""

    'XPTxtChqueNum.text = ""
    'DCboBankName.BoundText = ""
    'XPDTPDueDate.Value = Date

    'If Not RsNotes.EOF Or RsNotes.BOF Then
    '    For Num = 1 To RsNotes.RecordCount
    '        If RsNotes("NoteType").Value = 0 Then
    '            XPChkPayType(0).Value = Checked
    '            XPChkPayType_Click (0)
    '            XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").Value), "", (RsNotes("Note_Value").Value))
    '            XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").Value), "", Trim(RsNotes("NoteSerial").Value))
    '            Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").Value), "", Trim(RsNotes("BoxID").Value))
    '        End If
    '        If RsNotes("NoteType").Value = 1 Then
    '            XPChkPayType(1).Value = Checked
    '            XPChkPayType_Click (1)
    '            XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").Value), "", (RsNotes("Note_Value").Value))
    '            XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").Value), "", Trim(RsNotes("NoteSerial").Value))
    '             DtpDelayDate.Value = IIf(IsNull(RsNotes("DueDate").Value), "", (RsNotes("DueDate").Value))
    '        End If
    '        If RsNotes("NoteType").Value = 2 Then
    '            XPChkPayType(2).Value = Checked
    '            XPChkPayType_Click (2)
    '            XPTxtValue(2).text = IIf(IsNull(RsNotes("Note_Value").Value), "", (RsNotes("Note_Value").Value))
    '            XPTxtChqueNum.text = IIf(IsNull(RsNotes("ChqueNum").Value), "", Trim(RsNotes("ChqueNum").Value))
    '            Me.DCboBankName.BoundText = IIf(IsNull(RsNotes("BankID").Value), "", RsNotes("BankID").Value)
    '            XPDTPDueDate.Value = IIf(IsNull(RsNotes("DueDate").Value), "", (RsNotes("DueDate").Value))
    '        End If
    '        RsNotes.MoveNext
    '        If FG.Rows > 10 Then
    '            If Num = 8 Then FG.Refresh
    '        End If
    '    Next Num
    'End If
ErrTrap:
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTab301.CurrTab = 0 Then
            'XPBtnAdd.Enabled = True
            'XPBtnRemove.Enabled = True
        Else
            'XPBtnAdd.Enabled = False
            'XPBtnRemove.Enabled = False
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
            MaintenReport.MaintenanceDataShort XPTxtMaintanenceID.text, 1
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
    Dim IntResult As Integer
    Dim StrMSG As String

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
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnNewClients_Click
        End If
    End If

    'If Shift = 2 Then
    '    XPTab301.SetFocus
    '    If KeyCode = vbKeyTab Then
    '        If XPTab301.CurrTab = 0 Then
    '            XPTab301.CurrTab = 1
    '            If XPChkPayType(0).Enabled = True Then
    '                XPChkPayType(0).SetFocus
    '            End If
    '        Else
    '            XPTab301.CurrTab = 0
    '            FG.SetFocus
    '        End If
    '    End If
    'End If

    Exit Sub
ErrTrap:
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

    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    On Error GoTo ErrTrap

    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsCodes Me.DCboItemsCode
    Dcombos.GetItemsNames Me.DCboItemsName
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

Private Sub LoadTBR()

    With Me.TBar
        .Buttons.Clear
        .AllowCustomize = False
        .Appearance = ccFlat
        .BorderStyle = ccNone
        .Style = tbrFlat
        .TextAlignment = tbrTextAlignBottom
        Set .ImageList = mdifrmmain.ImgLstTree
        .Buttons.add , "RemoveRow", , , "Minus"

    End With

End Sub

Private Function CheckItemInv(LngItemID As Long, _
                              StrItemSerial As String, _
                              LngTransID As Long) As Boolean
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    StrSQL = "select * From QryGuarantee where Item_ID=" & LngItemID
    StrSQL = StrSQL + " and ItemSerial='" & StrItemSerial & "'"
    StrSQL = StrSQL + " AND Transaction_ID='" & LngTransID & "'"
    StrSQL = StrSQL + " AND Transaction_Type=2"
    StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        CheckItemInv = False
    Else
        CheckItemInv = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Sub RetriveTicketNO(LngTicketID As Long)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngManID As Long
    Dim LngTempID As Long

    Me.TxtTicketNO.text = LngTicketID
    '---------------------------------
    Me.DcboReItemCode.BoundText = ""
    Me.DcboReItemName.BoundText = ""
    Me.TxtNewSerial.text = ""
    Me.TxtReItemQty.text = ""
    DcboReStore.BoundText = ""
    '---------------------------------
    LngManID = val(Me.TxtOrgManID.text)
    StrSQL = "Select * From TblMainteneceDetails Where TblMainteneceDetails.TicketNO=" & LngTicketID & ""
    StrSQL = StrSQL + " AND TblMainteneceDetails.MaintananceID=" & LngManID & ""

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.DCboItemsCode.BoundText = rs("ItemID").value
        Me.DCboItemsName.BoundText = rs("ItemID").value
        Me.TxtSerial.text = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
        Me.TxtQuantity.text = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
        Me.DCboSupDeci.BoundText = IIf(IsNull(rs("SupDeci").value), -1, rs("SupDeci").value)
        Me.TxtCost.text = IIf(IsNull(rs("Cost").value), "", rs("Cost").value)
        LngTempID = IIf(IsNull(rs("TableID").value), "", rs("TableID").value)
    End If

    If LngTempID <> 0 Then
        Set rs = New ADODB.Recordset
        StrSQL = "Select * From TblManDetailsReplacedItems Where ManDetID=" & LngTempID & ""
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        '--------------------------
        If Not (rs.BOF Or rs.EOF) Then
            Me.DcboReItemCode.BoundText = rs("ItemID").value
            Me.DcboReItemName.BoundText = rs("ItemID").value
            Me.TxtNewSerial.text = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            Me.TxtReItemQty.text = IIf(IsNull(rs("ItemQty").value), "", rs("ItemQty").value)
            DcboReStore.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
        End If

        '--------------------------
    End If

End Sub
