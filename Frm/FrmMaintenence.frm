VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintenence 
   Caption         =   "œŒÊ· ··’Ì«‰…"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   HelpContextID   =   80
   Icon            =   "FrmMaintenence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   8580
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   8070
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   8580
      _cx             =   15134
      _cy             =   14235
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
      _GridInfo       =   $"FrmMaintenence.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   7065
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   5310
            TabIndex        =   27
            Top             =   45
            Width           =   2130
            _ExtentX        =   3757
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
            Left            =   7530
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   75
            Width           =   945
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   135
            Width           =   675
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   120
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1530
         Index           =   0
         Left            =   15
         TabIndex        =   19
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
         Begin VB.TextBox TxtReciptNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   30
            Width           =   1005
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   810
            Width           =   3135
         End
         Begin VB.TextBox TxtTransID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   1200
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox CboMaintenanceType 
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   30
            Width           =   1545
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6150
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   1170
            Width           =   1185
         End
         Begin VB.TextBox XPTxtMaintanenceID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6780
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   60
            Width           =   795
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   375
            Left            =   4410
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   661
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
            ButtonImage     =   "FrmMaintenence.frx":0420
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   4770
            TabIndex        =   5
            Top             =   450
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoInDtae 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   390
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96337921
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   750
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96337921
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   30
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   1110
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdSearchTrans 
            Height          =   345
            Left            =   5610
            TabIndex        =   7
            Top             =   1170
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmMaintenence.frx":07BA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdOpenTrans 
            Height          =   345
            Left            =   5010
            TabIndex        =   59
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmMaintenence.frx":0B54
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.CheckBox ChkInv 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   8220
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·≈Ì’«·"
            Height          =   435
            Index           =   18
            Left            =   6180
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   30
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   12
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… «·»Ì⁄"
            Height          =   255
            Index           =   9
            Left            =   7290
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1230
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   24
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   375
            Index           =   25
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   -30
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’Ì«‰…"
            Height          =   375
            Index           =   10
            Left            =   4530
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   315
            Index           =   8
            Left            =   7590
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   90
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Œ—ÊÃ «·„ Êﬁ⁄"
            Height          =   375
            Index           =   7
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   690
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   6
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   480
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·œŒÊ·"
            Height          =   315
            Index           =   3
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   390
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4860
         Left            =   15
         TabIndex        =   22
         Top             =   2190
         Width           =   8550
         _cx             =   15081
         _cy             =   8572
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
         Caption         =   "«·√’‰«›"
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
         Picture(0)      =   "FrmMaintenence.frx":0EEE
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4395
            Index           =   2
            Left            =   45
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   45
            Width           =   8460
            _cx             =   14923
            _cy             =   7752
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
            GridRows        =   4
            GridCols        =   2
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmMaintenence.frx":1288
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   915
               Index           =   4
               Left            =   0
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1365
               Width           =   8460
               _cx             =   14923
               _cy             =   1614
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
               Appearance      =   5
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
               Begin MSDataListLib.DataCombo DcboReStore 
                  Height          =   315
                  Left            =   3480
                  TabIndex        =   77
                  Top             =   600
                  Width           =   3195
                  _ExtentX        =   5636
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.TextBox TxtReItemQty 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   270
                  Width           =   765
               End
               Begin VB.TextBox TxtReItemSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   270
                  Width           =   1605
               End
               Begin MSDataListLib.DataCombo DcboReItemName 
                  Height          =   315
                  Left            =   3480
                  TabIndex        =   69
                  Top             =   270
                  Width           =   3165
                  _ExtentX        =   5583
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboReItemCode 
                  Height          =   315
                  Left            =   6660
                  TabIndex        =   67
                  Top             =   270
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.CheckBox ChkFastReplace 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈” »œ«· ›Ê—Ï"
                  Height          =   285
                  Left            =   7110
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   30
                  Width           =   1305
               End
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   285
                  Index           =   1
                  Left            =   930
                  TabIndex        =   74
                  Top             =   270
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmMaintenence.frx":12E6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   285
                  Index           =   0
                  Left            =   3030
                  TabIndex        =   75
                  Top             =   270
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmMaintenence.frx":1680
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Œ“‰"
                  Height          =   255
                  Index           =   17
                  Left            =   6660
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   630
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   255
                  Index           =   16
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   30
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   255
                  Index           =   15
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   60
                  Width           =   1395
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·’‰›"
                  Height          =   255
                  Index           =   14
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   60
                  Width           =   2685
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   255
                  Index           =   13
                  Left            =   7590
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   360
                  Width           =   915
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1365
               Index           =   5
               Left            =   0
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   0
               Width           =   8460
               _cx             =   14923
               _cy             =   2408
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
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   285
                  Width           =   885
               End
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Index           =   1
                  Left            =   990
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   900
                  Width           =   3735
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2340
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   285
                  Width           =   1575
               End
               Begin VB.TextBox Txt 
                  Alignment       =   1  'Right Justify
                  Height          =   450
                  Index           =   0
                  Left            =   4950
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   900
                  Width           =   3465
               End
               Begin VB.TextBox TxtTicketNO 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   13
                  Top             =   285
                  Width           =   1395
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   3990
                  TabIndex        =   10
                  Top             =   285
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   6180
                  TabIndex        =   9
                  Top             =   285
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   405
                  Left            =   90
                  TabIndex        =   16
                  Top             =   810
                  Width           =   540
                  _ExtentX        =   953
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
                  ButtonImage     =   "FrmMaintenence.frx":1A1A
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
               Begin ImpulseButton.ISButton CmdShowTransItems 
                  Height          =   345
                  Left            =   7590
                  TabIndex        =   58
                  Top             =   240
                  Width           =   735
                  _ExtentX        =   1296
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
                  ButtonImage     =   "FrmMaintenence.frx":22F4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   180
                  Index           =   0
                  Left            =   1410
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   30
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„·«ÕŸ… «·„»œ∆Ì… ··„Œ ’"
                  Height          =   210
                  Index           =   23
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   645
                  Width           =   2205
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   225
                  Index           =   28
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   30
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   210
                  Index           =   30
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   0
                  Width           =   1920
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   225
                  Index           =   31
                  Left            =   6090
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   30
                  Width           =   1470
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· ﬂ "
                  Height          =   180
                  Index           =   11
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   30
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘ﬂÊÏ «·⁄„Ì·"
                  Height          =   210
                  Index           =   5
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   645
                  Width           =   1515
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   1740
               Left            =   0
               TabIndex        =   17
               Top             =   2280
               Width           =   8460
               _cx             =   14922
               _cy             =   3069
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmMaintenence.frx":268E
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin MSComctlLib.Toolbar TBar 
               Height          =   390
               Left            =   465
               TabIndex        =   56
               Top             =   4020
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   688
               ButtonWidth     =   609
               ButtonHeight    =   582
               Appearance      =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   1530
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   645
         Visible         =   0   'False
         Width           =   750
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   7515
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
            TabIndex        =   34
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   6675
            TabIndex        =   35
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   5700
            TabIndex        =   36
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   4770
            TabIndex        =   37
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   3825
            TabIndex        =   38
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   2880
            TabIndex        =   39
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            TabIndex        =   40
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   1890
            TabIndex        =   41
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   960
            TabIndex        =   42
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
         TabIndex        =   80
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
         Caption         =   "œŒÊ· «·’Ì«‰…"
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
            Left            =   1110
            TabIndex        =   81
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
            ButtonImage     =   "FrmMaintenence.frx":289F
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
            Left            =   630
            TabIndex        =   82
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmMaintenence.frx":2C39
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
            Left            =   1710
            TabIndex        =   83
            Top             =   120
            Width           =   375
            _ExtentX        =   661
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
            ButtonImage     =   "FrmMaintenence.frx":2FD3
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
            Left            =   90
            TabIndex        =   84
            Top             =   120
            Width           =   510
            _ExtentX        =   900
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
            ButtonImage     =   "FrmMaintenence.frx":336D
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
Attribute VB_Name = "FrmMaintenence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim TTD As clstooltipdemand
Dim MaintenReport As ClsMaintananceReport
Dim cSearchDcbo(6) As clsDCboSearch

Public BolPrint As Boolean

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Me.Caption = "Maintenance Entry"
    C1Elastic6.Caption = Me.Caption
    lbl(8).Caption = " ID"
    lbl(18).Caption = "Recive No."
    lbl(10).Caption = "Gran. Type"

    With FG
        .TextMatrix(0, .ColIndex("code")) = "code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("count")) = "Qty"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "HaveSerial"
        .TextMatrix(0, .ColIndex("serial")) = "serial"
        .TextMatrix(0, .ColIndex("TicketNO")) = "TicketNO"
        .TextMatrix(0, .ColIndex("CusNotes")) = "Cus. Complaint"
        .TextMatrix(0, .ColIndex("TicketNO")) = "TicketNO"
        .TextMatrix(0, .ColIndex("EmpNotes")) = "Initial Notes"
    End With

    lbl(6).Caption = "Customer"
    lbl(12).Caption = "Cash Customer Name"
    lbl(3).Caption = "Date In"
    lbl(7).Caption = "Expect. Out Date"
    lbl(24).Caption = "Store Name"
    lbl(9).Caption = "Bill NO#"
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(28).Caption = "Item Serial"
    lbl(0).Caption = "Quantity"
    lbl(11).Caption = "Ticket NO"
    lbl(5).Caption = "Cust. Complaint"
    lbl(23).Caption = "Initial Notes"
    ChkFastReplace.Caption = "Immediate replacement"
    lbl(13).Caption = "Item Code"
    lbl(14).Caption = "Item Name"
    lbl(15).Caption = "Item Serial"
    lbl(16).Caption = "Quantity"
    lbl(17).Caption = "Store Name"

    lbl(4).Caption = "BY Employee"
    lbl(25).Caption = " Employee"
  
    lbl(1).Caption = "Curr. Rec."
    lbl(2).Caption = "Rec. Count:"

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

Private Sub CboMaintenanceType_Change()
    On Error GoTo ErrTrap

    Exit Sub
ErrTrap:
End Sub

Private Sub CboMaintenanceType_Click()
    CboMaintenanceType_Change
End Sub

Private Sub ChkFastReplace_Click()

    If Me.ChkFastReplace.value = vbChecked Then
        Me.lbl(13).Enabled = True
        Me.lbl(14).Enabled = True
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboReItemCode.Enabled = True
        Me.DcboReItemName.Enabled = True
        Me.CmdSearch(0).Enabled = True
        Me.CmdSearch(1).Enabled = True
        Me.TxtReItemSerial.Enabled = True
        Me.TxtReItemQty.Enabled = True
        Me.DcboReStore.Enabled = True

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            Me.DcboReItemCode.BoundText = Me.DCboItemsCode.BoundText
            Me.DcboReItemName.BoundText = Me.DCboItemsName.BoundText
            Me.TxtReItemQty.text = val(Me.TxtQuantity.text)
        End If

    Else
        Me.lbl(13).Enabled = False
        Me.lbl(14).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        Me.CmdSearch(0).Enabled = False
        Me.CmdSearch(1).Enabled = False
        Me.TxtReItemSerial.Enabled = False
        Me.TxtReItemQty.Enabled = False
        Me.DcboReStore.Enabled = False
    End If

End Sub

Private Sub ChkInv_Click()
    On Error GoTo ErrTrap

    If ChkInv.value = vbUnchecked Then
        TxtTransSerial.Enabled = False
        lbl(9).Enabled = False
        'CmdSearch.Enabled = False
        CmdSearchTrans.Enabled = False
        CmdOpenTrans.Enabled = False
    ElseIf ChkInv.value = vbChecked Then
        TxtTransSerial.Enabled = True
        lbl(9).Enabled = True
        'CmdSearch.Enabled = True
        CmdSearchTrans.Enabled = True
        CmdOpenTrans.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAdd_Click()
    '“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·

    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim LngTransID As Long
    Dim LngFindRow As Long
    Dim LngRow As Long
    Dim LngNewTicketID As Long
    Dim LngQty As Long
    Dim RsReItemStock As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset

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

    LngQty = val(Me.TxtQuantity.text)

    If Me.TxtSerial.Enabled = True And Trim(Me.TxtSerial.text) = "" Then
        Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·’‰›...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerial.SetFocus
        Exit Sub
    End If

    If Me.CboMaintenanceType.ListIndex = -1 Then
        Msg = "»—Ã«¡  ÕœÌœ ‰Ê⁄ «·’Ì«‰… ﬁ»· ≈÷«›… «·’‰› ..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboMaintenanceType.SetFocus
        Exit Sub
    End If

    If Me.ChkInv.value = vbChecked Then
        If Trim(Me.TxtTransSerial.text) = "" Then
            Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ ›« Ê—… «·»Ì⁄...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Exit Sub
        End If

        If val(Me.TxtTransID.text) = 0 Then
            TxtTransSerial_KeyPress 13
        End If

        'If Trim(Me.TxtSerial.text) <> "" Then
        If CheckItemInv(Me.DCboItemsCode.BoundText, Trim(Me.TxtSerial.text), val(Me.TxtTransID.text)) = False Then
            Msg = "«·ﬁÿ⁄… „‰ «·’‰› : " & Me.DCboItemsName.text
            Msg = Msg & Chr(13) & "—ﬁ„ : " & Me.TxtSerial.text
            Msg = Msg & Chr(13) & "€Ì— „”Ã·… ›Ï «·›« Ê—… —ﬁ„ : " & Me.TxtTransSerial.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        'End If
    End If

    If Trim$(Me.TxtSerial.text) <> "" Then
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

        If Not (rs.BOF Or rs.EOF) Then
            Msg = "Â–Â «·ﬁÿ⁄… „ÊÃÊœ… ›Ï «·„Œ“‰ ›⁄·«,,,"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    If Me.ChkFastReplace.value = vbChecked Then

        If DcboReItemCode.text = "" Then
            Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰› «·„” »œ·..!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboReItemCode.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.DcboReItemName.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰› «·„” »œ·...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboReItemName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Me.TxtReItemSerial.Enabled = True And Trim(Me.TxtReItemSerial.text) = "" Then
            Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·’‰› «·„” »œ·...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReItemSerial.SetFocus
            Exit Sub
        End If

        If val(Me.TxtReItemQty.text) = 0 Then
            Msg = "»—Ã«¡ ≈œŒ«· ﬂ„Ì… «·’‰› «·„” »œ·...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReItemQty.SetFocus
            Exit Sub
        End If

        If Me.DcboReStore.BoundText = "" Then
            Msg = "»—Ã«¡  ÕœÌœ «·„Œ“‰ «·Œ«—Ã „‰Â «·’‰› «·„” »œ·...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboReStore.SetFocus
            Exit Sub
        End If

        Set RsReItemStock = GetItemQuantityStock(val(Me.DcboReItemName.BoundText), val(Me.DcboReStore.BoundText), XPDtbGoInDtae.value, , , Me.DcboReItemName.text, Trim$(Me.TxtReItemSerial.text))

        If RsReItemStock.BOF Or RsReItemStock.EOF Then
            Msg = "⁄›Ê« ﬂ„Ì… «·’‰› «·„” »œ· ... «Ê «·”Ì—Ì«· «·„œŒ· €Ì— „ÊÃÊœ ›Ï «·„Œ“‰...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReItemQty.SetFocus
            Exit Sub
        End If

        If RsReItemStock("Qty").value < val(Me.TxtReItemQty.text) Then
            Msg = "⁄›Ê« ﬂ„Ì… «·’‰› «·„” »œ·...€Ì— „ «Õ… ›Ï «·„Œ“‰ !"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReItemQty.SetFocus
            Exit Sub
        End If
    End If

    If val(Me.TxtTicketNO.text) <> 0 Then
        LngFindRow = FG.FindRow(val(Me.TxtTicketNO.text), FG.FixedRows, FG.ColIndex("TicketNO"), False, True)
    End If

    If Me.TxtModFlg.text = "N" Then
        If LngFindRow <= 0 Then
            LngNewTicketID = new_id("TblMainteneceDetails", "TicketNO", "")
            Me.TxtTicketNO.text = LngNewTicketID
        Else
            LngNewTicketID = Me.TxtTicketNO.text
        End If

        Do While ItemCount < LngQty

            With FG

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
            
                .TextMatrix(LngRow, .ColIndex("TicketNO")) = LngNewTicketID
                .TextMatrix(LngRow, .ColIndex("CusNotes")) = Txt(0).text
                .TextMatrix(LngRow, .ColIndex("EmpNotes")) = Txt(1).text
            
                If TxtSerial.Tag = "T" Then
                    .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexChecked
                ElseIf TxtSerial.Tag = "F" Then
                    .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                .AutoSize 0, .Cols - 1, False
            End With

            ItemCount = ItemCount + 1
            LngNewTicketID = LngNewTicketID + 1
        Loop

        With Me.FG

            If Me.ChkFastReplace.value = vbChecked Then
                LngFindRow = FG.FindRow("Rep-" & LngRow, FG.FixedRows, FG.ColIndex("RowFlag"), False, True)

                If LngFindRow = -1 Then
                    .AddItem "", LngRow + 1
                Else
                
                End If

                .TextMatrix(LngRow + 1, .ColIndex("Name")) = Me.DcboReItemName.BoundText
                .TextMatrix(LngRow + 1, .ColIndex("Code")) = Me.DcboReItemName.BoundText
                .TextMatrix(LngRow + 1, .ColIndex("Serial")) = Trim$(Me.TxtReItemSerial.text)

                If TxtReItemSerial.Enabled = True Then
                    .Cell(flexcpChecked, LngRow + 1, .ColIndex("HaveSerial")) = flexChecked
                ElseIf TxtReItemSerial.Enabled = False Then
                    .Cell(flexcpChecked, LngRow + 1, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                .TextMatrix(LngRow + 1, .ColIndex("Count")) = Trim(Me.TxtReItemQty.text)
                .TextMatrix(LngRow + 1, .ColIndex("TicketNO")) = .TextMatrix(LngRow, .ColIndex("TicketNO"))
                .TextMatrix(LngRow + 1, .ColIndex("EmpNotes")) = "≈” »œ«· „‰ " & Me.DcboReStore.text
                .Cell(flexcpData, LngRow + 1, .ColIndex("EmpNotes")) = Me.DcboReStore.BoundText
                .TextMatrix(LngRow + 1, .ColIndex("ManDetID")) = LngRow
                .TextMatrix(LngRow + 1, .ColIndex("RowFlag")) = "Rep-" & LngRow
                .Cell(flexcpBackColor, LngRow + 1, 1, LngRow + 1, .Cols - 1) = vbGreen
            End If

            .AutoSize 0, .Cols - 1, False
        End With

    ElseIf Me.TxtModFlg.text = "E" Then

        With FG

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
        
            .TextMatrix(LngRow, .ColIndex("TicketNO")) = Me.TxtTicketNO.text
            .TextMatrix(LngRow, .ColIndex("CusNotes")) = Txt(0).text
            .TextMatrix(LngRow, .ColIndex("EmpNotes")) = Txt(1).text
        
            If TxtSerial.Tag = "T" Then
                .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexChecked
            ElseIf TxtSerial.Tag = "F" Then
                .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexUnchecked
            End If

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    DCboItemsCode.BoundText = ""
    DCboItemsName.BoundText = ""
    TxtSerial.text = ""
    TxtTicketNO.text = ""
    Me.Txt(0).text = ""
    Me.Txt(1).text = ""
    '----------------
    Me.DcboReItemCode.BoundText = ""
    Me.DcboReItemName.BoundText = ""
    Me.TxtReItemSerial.text = ""
    Me.TxtReItemQty.text = ""
    Me.DcboReStore.BoundText = ""
    Me.ChkFastReplace.value = vbUnchecked
    '----------------
    FG.SetFocus
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdReplace_Click()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsSerial As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·’‰› «·–Ì  —€» ›Ì «” »œ«·Â "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If DBCboClientName.text = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·" & Chr(13)
        Msg = Msg + "«·–Ì ﬁ«„ »‘—«¡ Â–Â «·ﬁÿ⁄…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If CboMaintenanceType.ListIndex = 1 Then
        If TxtTransSerial.text = "" Then
            Msg = Msg + "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ " & Chr(13)
            Msg = Msg + "«· Ì  „ »Ì⁄ Â–« «·’‰› ›ÌÂ«"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

                        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbNo Then
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
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

            Else
                Msg = "Â–Â «·⁄„·Ì… Œ«’… »«·√’‰«› «· Ì   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdOpenTrans_Click()

    If val(Me.TxtTransID.text) <> 0 Then
        OpenScreen InvoiceScreen, val(Me.TxtTransID.text)
    End If

End Sub

Private Sub CmdSearch_Click(Index As Integer)
    Dim StrMSG As String

    Dim LngItemID As Long, LngStoreID As Long

    Select Case Index

        Case 0
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DcboReItemName
            FrmItemSearch.show vbModal

        Case 1
    
            LngItemID = val(Me.DcboReItemName.BoundText)
            LngStoreID = val(Me.DcboReStore.BoundText)

            If LngItemID = 0 Then
                Set TTD = New clstooltipdemand
                Set TTD.m_From = Me
                TTD.Style = TTBalloon
                TTD.Icon = TTIconError
                TTD.Centered = True
                TTD.RightToLeft = True
                TTD.CreateToolTip DcboReItemName.hwnd
                TTD.DelayTime = 250
                TTD.VisibleTime = 5000
                StrMSG = "Œÿ« ›Ï ≈Œ Ì«— «·’‰›...!!!"
                TTD.title = StrMSG
                StrMSG = "ÌÃ» «‰  ﬁÊ„ »≈Œ ÌÌ«— «·’‰› ﬁ»· ⁄—÷ "
                StrMSG = StrMSG & "«·”Ì—Ì«· ‰„»— «·„ «Õ… „‰Â"
                TTD.TipText = StrMSG
                TTD.PopupOnDemand = True
                TTD.show (DcboReItemName.Width / Screen.TwipsPerPixelY), (DcboReItemName.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
                Exit Sub
            Else

                If Not TTD Is Nothing Then
                    TTD.Destroy
                End If
            End If

            If LngStoreID = 0 Then
                Set TTD = New clstooltipdemand
                Set TTD.m_From = Me
                TTD.Style = TTBalloon
                TTD.Icon = TTIconError
                TTD.Centered = True
                TTD.RightToLeft = True
                TTD.CreateToolTip DcboReStore.hwnd
                TTD.DelayTime = 250
                TTD.VisibleTime = 5000
                StrMSG = "Œÿ« ›Ï ≈Œ Ì«— «·’‰›...!!!"
                TTD.title = StrMSG
                StrMSG = "ÌÃ» «‰  ﬁÊ„ »≈Œ ÌÌ«— «·„Œ“‰ ﬁ»· ⁄—÷ "
                StrMSG = StrMSG & "«·”Ì—Ì«· ‰„»— «·„ «Õ… „‰ «·’‰›"
                TTD.TipText = StrMSG
                TTD.PopupOnDemand = True
                TTD.show (DcboReStore.Width / Screen.TwipsPerPixelY), (DcboReStore.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
                Exit Sub
            Else

                If Not TTD Is Nothing Then
                    TTD.Destroy
                End If
            End If

            If LngItemID = 0 Or LngStoreID = 0 Then
                Exit Sub
            End If

            'Load FrmSerialList
            'FrmSerialList.RetrunType = 1
            'Set FrmSerialList.m_TextBox = TxtReItemSerial
            'FrmSerialList.GetData LngItemID, LngStoreID
            'FrmSerialList.show vbModal
    End Select

End Sub

Private Sub CmdSearchTrans_Click()
    ' ›« Ê—… „»Ì⁄« 
    Load FrmBuySearch
    FrmBuySearch.DealingForm = InvoiceTransaction
    Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
    FrmBuySearch.CboPaymentType.Enabled = False
    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
    FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    FrmBuySearch.show vbModal
End Sub

Private Sub CmdShowTransItems_Click()

    If val(Me.TxtTransID.text) = 0 Then
        Exit Sub
    End If

    Load FrmManChooseItems
    Set FrmManChooseItems.MyForm = Me
    FrmManChooseItems.LoadTrans val(Me.TxtTransID.text), InvoiceTransaction
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

Private Sub DcboReItemCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DcboReItemCode.BoundText <> "" Then
        DcboReItemName.BoundText = DcboReItemCode.BoundText
    Else
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsCode.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtReItemSerial.Enabled = True
            Me.TxtReItemQty.Enabled = False
            Me.TxtReItemQty.text = 1
            TxtReItemSerial.Tag = "T"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtReItemSerial.Enabled = False
            Me.TxtReItemQty.Enabled = True
            Me.TxtReItemQty.text = 1
            TxtReItemSerial.Tag = "F"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DcboReItemName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DcboReItemName.BoundText <> "" Then
        DcboReItemCode.BoundText = DcboReItemName.BoundText
    Else
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtReItemSerial.Enabled = True
            TxtReItemQty.Enabled = False
            TxtReItemQty.text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtReItemSerial.Enabled = False
            TxtReItemQty.Enabled = True
            TxtReItemQty.text = ""
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
    On Error GoTo ErrTrap
    Dim RsSerial As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim Msg As String
    Dim StrSQL As String

    If XPDtbGoInDtae.value = "" Then
        Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ⁄„·Ì… «·’Ì«‰…" & Chr(13)
        Msg = Msg + "ﬁ»· ≈œŒ«· »Ì«‰«  «·√’‰«›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    
                        '»Ì«‰«  «·›« Ê—… «· Ì  „ »Ì⁄ «·ﬁÿ⁄Â ›ÌÂ«
                        StrSQL = "select * From QryGuarantee where Item_ID=" & FG.TextMatrix(Row, FG.ColIndex("Code")) & " and ItemSerial='" & FG.TextMatrix(Row, FG.ColIndex("Serial")) & "'"
                        Set RsTemp = New ADODB.Recordset
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                            Msg = "·ﬁœ  „ »Ì⁄ «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")) & Chr(13)
                            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(Row, FG.ColIndex("Serial")) & Chr(13)
                            Msg = Msg + "≈·Ï «·⁄„Ì· : " & RsTemp("CusName").value & Chr(13)
                            Msg = Msg + "›Ì «·›« Ê—… —ﬁ„ : " & RsTemp("Transaction_ID").value
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

                        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbNo Then
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
      '  FrmAddNewItem.DealingForm = Maintenance
      '  FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub Fg_Click()
    Dim i As Long
    Dim LngLoadRow As Long
    Dim StrTemp  As String
    Dim LngTemp  As Long

    On Error GoTo ErrTrap

    With FG

        If .Col = -1 Then Exit Sub
        If .Row <= 0 Then Exit Sub
        If Trim(.TextMatrix(.Row, .ColIndex("RowFlag"))) = "" Then
            If .TextMatrix(.Row, .ColIndex("Name")) <> "" Then
                Me.DCboItemsCode.BoundText = .TextMatrix(.Row, .ColIndex("Name"))
                Me.DCboItemsName.BoundText = .TextMatrix(.Row, .ColIndex("Name"))

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
                Me.Txt(0).text = .TextMatrix(.Row, .ColIndex("CusNotes"))
                Me.Txt(1).text = .TextMatrix(.Row, .ColIndex("EmpNotes"))
                LngLoadRow = .Row

                For i = .FixedRows To .Rows - 1

                    If Trim$(.TextMatrix(i, .ColIndex("RowFlag"))) <> "" Then
                        StrTemp = Trim$(.TextMatrix(i, .ColIndex("RowFlag")))
                        LngTemp = val(Mid$(StrTemp, InStr(1, StrTemp, "-", vbTextCompare) + 1))

                        If LngTemp = LngLoadRow Then
                            Me.ChkFastReplace.value = vbChecked
                            Me.DcboReItemCode.BoundText = .TextMatrix(i, .ColIndex("Name"))
                            Me.DcboReItemName.BoundText = .TextMatrix(i, .ColIndex("Name"))
                            Me.TxtReItemSerial.text = .TextMatrix(i, .ColIndex("Serial"))
                            Me.TxtReItemQty.text = .TextMatrix(i, .ColIndex("Count"))
                            Me.DcboReStore.BoundText = val(.Cell(flexcpData, i, .ColIndex("EmpNotes")))
                            Exit For
                        End If
                    End If

                Next i

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
    '
    'On Error GoTo ErrTrap
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

    Resize_Form Me, TransactionSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'AddTip
    SetDtpickerDate Me.XPDtbGoInDtae
    SetDtpickerDate XPDtbGoOutDtae
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName

    With CboMaintenanceType

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "Œ«—Ã «·÷„«‰"
            .AddItem "œ«Œ· «·÷„«‰"
        Else
            .AddItem "OutSide"
            .AddItem "Inside"
        End If

    End With

    FG.WallPaper = BGround.Picture

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboEmp

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    LoadTBR
    Set rs = New ADODB.Recordset
    rs.Open "Select * From  TblMaintenece Where ManOperationTypeID=1", Cn, adOpenStatic, adLockOptimistic, adCmdText

    StrSQL = "Select * From TblItems"
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
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
    Me.ChkFastReplace.value = vbUnchecked
    ChkFastReplace_Click
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
    Set TTD = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    With Button

        Select Case .key

            Case "RemoveRow"

                If FG.Rows > 1 Then
                    If FG.Rows = 2 Then
                        Me.FG.Clear flexClearScrollable, flexClearEverything
                    Else

                        If Me.FG.Rows > 1 Then
                            If Me.FG.Row <> Me.FG.FixedRows - 1 Then
                                Me.FG.RemoveItem (Me.FG.Row)
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
       
            TxtTransSerial.locked = True
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

            CboMaintenanceType.locked = True
            Ele(5).Enabled = False
            Ele(4).Enabled = False
            Me.TBar.Buttons("RemoveRow").Enabled = False
            Me.TxtTicketNO.Enabled = False

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
        
            TxtTransSerial.locked = False
            XPBtnNewClients.Enabled = True

            If XPTab301.CurrTab = 0 Then
                'XPBtnAdd.Enabled = True
                'XPBtnRemove.Enabled = True
            Else
                'XPBtnAdd.Enabled = False
                'XPBtnRemove.Enabled = False
            End If

            FG.Enabled = True
            FG.Rows = FG.FixedRows
            FG.Rows = 2
            '        FG.RowPosition(FG.Rows - 2) = FG.Rows - 1
            '        FG.TextMatrix(FG.Rows - 1, 2) = "«÷€ÿ Â‰«"
            Me.DBCboClientName.locked = False
            FG.Editable = flexEDNone
        
            XPDtbGoInDtae.value = Date '
            XPDtbGoOutDtae.value = Date
            CboMaintenanceType.locked = False
            CboMaintenanceType.ListIndex = 0
            CboMaintenanceType_Change
            Ele(5).Enabled = True
            Ele(4).Enabled = True
            Me.TBar.Buttons("RemoveRow").Enabled = True
            Me.TxtTicketNO.Enabled = False

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
        
            If XPTab301.CurrTab = 0 Then
                'XPBtnAdd.Enabled = True
                'XPBtnRemove.Enabled = True
            Else
                'XPBtnAdd.Enabled = False
                'XPBtnRemove.Enabled = False
            End If

            'XPBtnRemove.Enabled = True
            FG.Enabled = True
            Me.DBCboClientName.locked = False
            CboMaintenanceType.locked = False
            FG.Editable = flexEDNone
            DBCboClientName_Change
            CboMaintenanceType_Change
            Ele(5).Enabled = True
            Ele(4).Enabled = True
            Me.TBar.Buttons("RemoveRow").Enabled = True
            Me.TxtTicketNO.Enabled = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.text, 1)
End Sub

Private Sub TxtTransID_Change()
    Dim StrTemp As String
    Dim LngCuID As Long

    If Trim(Me.TxtTransID.text) = "" Then
        Me.TxtTransSerial.text = ""
    Else
        StrTemp = GetTransIDSerial(1, val(Me.TxtTransID.text), , 2, LngCuID)
        
        If Trim$(Me.TxtTransSerial.text) <> StrTemp Then
            Me.TxtTransSerial.text = StrTemp
        End If

        If val(Me.DBCboClientName.BoundText) <> LngCuID Then
            Me.DBCboClientName.BoundText = LngCuID
        End If
    End If

End Sub

Private Sub TxtTransSerial_Change()
    'Dim StrTemp As String
    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    '    If Trim(Me.TxtTransSerial.text) = "" Then
    '        Me.TxtTransID.text = ""
    '    Else
    '        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)
    '        If Trim$(Me.TxtTransID.text) <> StrTemp Then
    '            Me.TxtTransID.text = StrTemp
    '        End If
    '    End If
    'End If
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

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            PutTrans
        End If
    End If

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
            FrmMaintanenceSearch.searchtype = 1
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

Private Sub XPBtnNewClients_Click()
    On Error GoTo ErrTrap

    'With FrmAddNewCustemer
    '    .DealingForm = Maintenance
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    '    .show vbModal
    '    cSearchDcbo(0).Refresh
    'End With

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
    Dim RsManDetailReplace As ADODB.Recordset

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If val(Me.TxtReciptNumber.text) = 0 Then
            Msg = "ÌÃ» ≈Ì’«· «·œŒÊ· ··’Ì«‰…...!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReciptNumber.SetFocus
            Exit Sub
        End If

        If CboMaintenanceType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·’Ì«‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboMaintenanceType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If DBCboClientName.text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
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

        If ChkInv.value = vbChecked Then
            If TxtTransSerial.text = "" Then
                Msg = "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Sub
            ElseIf PutTrans = False Then
                Exit Sub
            End If
        End If

        If Me.DCboStoreName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰....!!! " & Chr(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
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
        End If

        RsDetails.Open "[TblMainteneceDetails]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Set RsManDetailReplace = New ADODB.Recordset
        RsManDetailReplace.Open "TblManDetailsReplacedItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        Cn.BeginTrans
        BeginTrans = True
        rs("MaintananceID").value = val(XPTxtMaintanenceID.text)
        rs("ReciptNumber").value = val(Me.TxtReciptNumber.text)
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)

        If Me.DBCboClientName.BoundText = 2 Then
            rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
        Else
            rs("CashCustomerName").value = Null
        End If

        rs("DateGoIN").value = XPDtbGoInDtae.value
        rs("DateGoOUT").value = XPDtbGoOutDtae.value
        rs("EmpID").value = Me.DcboEmp.BoundText
        rs("StoreID").value = Me.DCboStoreName.BoundText
        rs("UserID").value = user_id

        If CboMaintenanceType.ListIndex = -1 Then
            rs("MType").value = 0
        Else
            rs("MType").value = val(CboMaintenanceType.ListIndex)
        End If

        If ChkInv.value = vbChecked Then
            rs("Transaction_ID").value = IIf(Me.TxtTransID.text = "", Null, val(Me.TxtTransID.text))
        Else
            rs("Transaction_ID").value = Null
        End If

        rs("GoOut").value = 0
        rs("ManOperationTypeID").value = 1
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

            If FG.TextMatrix(RowNum, FG.ColIndex("RowFlag")) = "" Then
                RsDetails.AddNew
                RsDetails("MaintananceID").value = val(XPTxtMaintanenceID.text)
                RsDetails("ItemID").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Name"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))

                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                    RsDetails("ItemSerial").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                Else
                    RsDetails("ItemSerial").value = Null
                End If

                RsDetails("Quantity").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))
                RsDetails("TicketNO").value = Trim$(FG.TextMatrix(RowNum, FG.ColIndex("TicketNO")))
                RsDetails("CustomerNotes").value = Trim$(FG.TextMatrix(RowNum, FG.ColIndex("CusNotes")))
                RsDetails("EmpNotes").value = Trim$(FG.TextMatrix(RowNum, FG.ColIndex("EmpNotes")))
                RsDetails.update
                FG.TextMatrix(RowNum, FG.ColIndex("TableID")) = RsDetails("TableID").value
            ElseIf FG.TextMatrix(RowNum, FG.ColIndex("RowFlag")) Like "Rep-" & "*" Then
                Dim StrTemp  As String
                Dim X As Integer, LngRow As Integer
            
                StrTemp = FG.TextMatrix(RowNum, FG.ColIndex("RowFlag"))
                X = Len("Rep-")
                LngRow = val(Mid$(StrTemp, X + 1))
            
                RsManDetailReplace.AddNew
                RsManDetailReplace("ManDetID").value = FG.TextMatrix(LngRow, FG.ColIndex("TableID"))
                RsManDetailReplace("ItemID").value = FG.TextMatrix(RowNum, FG.ColIndex("Name"))

                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                    RsManDetailReplace("ItemSerial").value = IIf(IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))), "", Trim(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                Else
                    RsManDetailReplace("ItemSerial").value = Null
                End If

                RsManDetailReplace("ItemQty").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))
                RsManDetailReplace("StoreID").value = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("EmpNotes")))
                RsManDetailReplace("ReplaceType").value = -1
                RsManDetailReplace.update
                FG.TextMatrix(RowNum, FG.ColIndex("TableID")) = RsManDetailReplace("ID").value
            End If

        Next RowNum
    
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
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ’Ì«‰… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·’Ì«‰…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ’Ì«‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
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
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
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
    Dim i As Integer, IntRow As Integer
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
    Me.TxtReciptNumber.text = IIf(IsNull(rs("ReciptNumber").value), "", (rs("ReciptNumber").value))
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)

    If DBCboClientName.BoundText = 2 Then
        Me.TxtCashCustomerName.text = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
    Else
        Me.TxtCashCustomerName.text = ""
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    XPDtbGoInDtae.value = IIf(IsNull(rs("DateGoIN").value), Date, rs("DateGoIN").value)
    XPDtbGoOutDtae.value = IIf(IsNull(rs("DateGoOUT").value), Date, rs("DateGoOUT").value)

    CboMaintenanceType.ListIndex = IIf(IsNull(rs("MType").value), 0, rs("MType").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)

    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkInv.value = vbChecked
        Me.TxtTransID.text = rs("Transaction_ID").value
    Else
        Me.ChkInv.value = vbUnchecked
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    End If

    ChkInv_Click
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT  dbo.TblMainteneceDetails.MaintananceID, dbo.TblMainteneceDetails.ItemID," & "dbo.TblItems.HaveSerial, dbo.TblMainteneceDetails.ItemSerial,dbo.TblMainteneceDetails.Quantity," & "dbo.TblMainteneceDetails.CustomerNotes, dbo.TblMainteneceDetails.TicketNO," & "dbo.TblMainteneceDetails.EmpNotes , dbo.TblMainteneceDetails.TableId "
    StrSQL = StrSQL + " FROM dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + " dbo.TblMainteneceDetails ON dbo.TblItems.ItemID = dbo.TblMainteneceDetails.ItemID"
    StrSQL = StrSQL + "  where MaintananceID=" & val(rs("MaintananceID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = FG.FixedRows

        For i = 1 To RsDetails.RecordCount
            FG.AddItem ""
            IntRow = FG.Rows - 1
            FG.TextMatrix(IntRow, FG.ColIndex("TableID")) = RsDetails("TableID").value
            FG.Cell(flexcpPicture, IntRow, FG.ColIndex("Replace")) = ""
            FG.Cell(flexcpData, IntRow, FG.ColIndex("Replace")) = ""
            FG.TextMatrix(IntRow, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            FG.TextMatrix(IntRow, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            FG.TextMatrix(IntRow, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial").value = True Then
                FG.Cell(flexcpChecked, IntRow, FG.ColIndex("HaveSerial")) = flexChecked
            Else
                FG.Cell(flexcpChecked, IntRow, FG.ColIndex("HaveSerial")) = flexUnchecked
            End If

            FG.TextMatrix(IntRow, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", Trim(RsDetails("Quantity").value))
            FG.TextMatrix(IntRow, FG.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").value))
            FG.TextMatrix(IntRow, FG.ColIndex("CusNotes")) = IIf(IsNull(RsDetails("CustomerNotes")), "", Trim(RsDetails("CustomerNotes").value))
            FG.TextMatrix(IntRow, FG.ColIndex("EmpNotes")) = IIf(IsNull(RsDetails("EmpNotes")), "", Trim(RsDetails("EmpNotes").value))
        
            '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «· Ì  „  ›Ì Â–Â «·⁄„·Ì…
            Set RsReplace = New ADODB.Recordset
            StrSQL = "SELECT     dbo.TblManDetailsReplacedItems.ID, dbo.TblManDetailsReplacedItems.Man" & "DetID, dbo.TblManDetailsReplacedItems.ItemID,dbo.TblManD" & "etailsReplacedItems.ItemSerial, dbo.TblManDetailsReplacedItems.ItemQty, dbo.TblM" & "anDetailsReplacedItems.StoreID,dbo.TblItems.HaveSerial, " & "dbo.TblStore.StoreName FROM         dbo.TblItems INNER JOIN                     " & "  dbo.TblManDetailsReplacedItems ON dbo.TblItems.ItemID = dbo.TblManDetailsRepla" & "cedItems.ItemID INNER JOIN dbo.TblStore ON dbo.TblManDetai" & "lsReplacedItems.StoreID = dbo.TblStore.StoreID"
        
            StrSQL = StrSQL + " Where ManDetID=" & RsDetails("TableID").value
            RsReplace.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsReplace.BOF Or RsReplace.EOF) Then
                FG.AddItem ""
                IntRow = FG.Rows - 1
                FG.TextMatrix(IntRow, FG.ColIndex("TableID")) = RsReplace("ID").value
                FG.Cell(flexcpPicture, IntRow, FG.ColIndex("Replace")) = ""
                FG.Cell(flexcpData, IntRow, FG.ColIndex("Replace")) = ""
                FG.TextMatrix(IntRow, FG.ColIndex("Code")) = IIf(IsNull(RsReplace("ItemID")), "", Trim(RsReplace("ItemID").value))
                FG.TextMatrix(IntRow, FG.ColIndex("Name")) = IIf(IsNull(RsReplace("ItemID")), "", Trim(RsReplace("ItemID").value))
                FG.TextMatrix(IntRow, FG.ColIndex("Serial")) = IIf(IsNull(RsReplace("ItemSerial")), "", Trim(RsReplace("ItemSerial").value))

                If RsReplace("HaveSerial").value = True Then
                    FG.Cell(flexcpChecked, IntRow, FG.ColIndex("HaveSerial")) = flexChecked
                Else
                    FG.Cell(flexcpChecked, IntRow, FG.ColIndex("HaveSerial")) = flexUnchecked
                End If

                FG.TextMatrix(IntRow, FG.ColIndex("Count")) = IIf(IsNull(RsReplace("ItemQty")), "", Trim(RsReplace("ItemQty").value))
                FG.TextMatrix(IntRow, FG.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").value))
                'Fg.TextMatrix(IntRow, Fg.ColIndex("CusNotes")) = IIf(IsNull(RsReplace("CustomerNotes")), "", Trim(RsReplace("CustomerNotes").Value))
                FG.TextMatrix(IntRow, FG.ColIndex("EmpNotes")) = "≈” »œ«· „‰ " & IIf(IsNull(RsReplace("StoreName")), "", Trim(RsReplace("StoreName").value))
                FG.Cell(flexcpData, IntRow, FG.ColIndex("EmpNotes")) = IIf(IsNull(RsReplace("StoreID")), "", Trim(RsReplace("StoreID").value))
                FG.TextMatrix(IntRow, FG.ColIndex("ManDetID")) = IIf(IsNull(RsReplace("ManDetID")), "", Trim(RsReplace("ManDetID").value))
                FG.TextMatrix(IntRow, FG.ColIndex("RowFlag")) = "Rep-" & IntRow - 1
                FG.Cell(flexcpBackColor, IntRow, 1, IntRow, FG.Cols - 1) = vbGreen
            End If

            RsDetails.MoveNext
        Next i

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
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
            XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
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
    Dim Msg As String

    On Error GoTo ErrTrap

    If val(Me.DBCboClientName.BoundText) = 2 Then
        Me.TxtCashCustomerName.Enabled = True
        Me.lbl(12).Enabled = True
    Else
        Me.TxtCashCustomerName.Enabled = False
        Me.lbl(12).Enabled = False
    End If

    Exit Sub
ErrTrap:

    If Err.Number = 7 Then
        Msg = "Ì⁄«‰Ï «·»—‰«„Ã „‰ ‰ﬁ’ ›Ï –«ﬂ—… «·ÃÂ«“"
        Msg = Msg & Chr(13) & "ÌÃ» €·ﬁ «·»—‰«„Ã Ê≈⁄«œ…  ‘€Ì· «·ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub FillItemData()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos

    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›
    Set Dcombos = New ClsDataCombos

    Dcombos.GetItemsCodes Me.DCboItemsCode
    '«”„ «·’‰›
    Dcombos.GetItemsNames Me.DCboItemsName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboItemsCode
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DCboItemsName

    Dcombos.GetItemsCodes Me.DcboReItemCode
    Set cSearchDcbo(4) = New clsDCboSearch
    Set cSearchDcbo(4).Client = Me.DcboReItemCode
    Dcombos.GetItemsNames Me.DcboReItemName
    Set cSearchDcbo(5) = New clsDCboSearch
    Set cSearchDcbo(5).Client = Me.DcboReItemName

    Dcombos.GetStores Me.DcboReStore
    Set cSearchDcbo(6) = New clsDCboSearch
    Set cSearchDcbo(6).Client = Me.DcboReStore
    Set Dcombos = Nothing
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

Private Function PutTrans() As Boolean
    Dim StrTemp As String
    Dim Msg As String

    If Trim(Me.TxtTransSerial.text) = "" Then
        Me.TxtTransID.text = ""
    Else
        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)

        If StrTemp = "" Then
            Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ... ≈” Œœ„ ‘«‘… «·»ÕÀ.!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            PutTrans = False
        Else

            If Trim$(Me.TxtTransID.text) <> StrTemp Then
                Me.TxtTransID.text = StrTemp
            End If

            PutTrans = True
        End If
    End If

End Function

