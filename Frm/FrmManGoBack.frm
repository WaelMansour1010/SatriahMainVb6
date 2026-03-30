VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManGoBack1 
   Caption         =   "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "FrmManGoBack.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   9000
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7980
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      _cx             =   15875
      _cy             =   14076
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
      _GridInfo       =   $"FrmManGoBack.frx":0CCA
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6975
         Width           =   8970
         _cx             =   15822
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
            Left            =   6630
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   30
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4095
            TabIndex        =   17
            Top             =   45
            Width           =   1440
            _ExtentX        =   2540
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
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   75
            Width           =   900
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   135
            Width           =   585
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   105
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   750
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   1
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·›« Ê—…"
            Height          =   255
            Index           =   0
            Left            =   7740
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   75
            Width           =   1215
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1215
         Index           =   0
         Left            =   15
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   645
         Width           =   8970
         _cx             =   15822
         _cy             =   2143
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5220
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   60
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox XPTxtMaintanenceID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6945
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   60
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   4770
            TabIndex        =   1
            Top             =   450
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoInDtae 
            Height          =   315
            Left            =   4770
            TabIndex        =   2
            Top             =   780
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            Format          =   95420417
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   90
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   420
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   240
            Index           =   24
            Left            =   3090
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   420
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   240
            Index           =   25
            Left            =   3090
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   315
            Index           =   8
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ"
            Height          =   300
            Index           =   6
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   465
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·—ÃÊ⁄"
            Height          =   315
            Index           =   3
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   795
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5085
         Left            =   15
         TabIndex        =   30
         Top             =   1875
         Width           =   8970
         _cx             =   15822
         _cy             =   8969
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
         Picture(0)      =   "FrmManGoBack.frx":0D60
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4620
            Index           =   2
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   8880
            _cx             =   15663
            _cy             =   8149
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
            _GridInfo       =   $"FrmManGoBack.frx":10FA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1530
               Index           =   5
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   0
               Width           =   8880
               _cx             =   15663
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
               Begin VB.TextBox TxtReItemQty 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1200
                  Width           =   825
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1230
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   285
                  Width           =   1275
               End
               Begin VB.TextBox TxtCost 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   630
                  Width           =   915
               End
               Begin VB.TextBox TxtNewSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2565
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   1200
                  Width           =   2070
               End
               Begin VB.ComboBox CboSupDeci 
                  Height          =   315
                  Left            =   4650
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   660
                  Width           =   3585
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2565
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   285
                  Width           =   2055
               End
               Begin VB.TextBox TxtTicketNo 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   10
                  Top             =   285
                  Width           =   1125
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   4665
                  TabIndex        =   7
                  Top             =   285
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   7065
                  TabIndex        =   6
                  Top             =   285
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   525
                  Left            =   60
                  TabIndex        =   33
                  Top             =   600
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   926
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
                  ButtonImage     =   "FrmManGoBack.frx":114E
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   0
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   0
               End
               Begin ImpulseButton.ISButton CmdShowTransItems 
                  Height          =   345
                  Left            =   8265
                  TabIndex        =   5
                  Top             =   270
                  Width           =   540
                  _ExtentX        =   953
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
                  ButtonImage     =   "FrmManGoBack.frx":1A28
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcboReItemName 
                  Height          =   315
                  Left            =   4665
                  TabIndex        =   76
                  Top             =   1200
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboReItemCode 
                  Height          =   315
                  Left            =   7065
                  TabIndex        =   77
                  Top             =   1200
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboReStore 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   78
                  Top             =   1200
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   81
                  Top             =   1200
                  Width           =   540
                  _ExtentX        =   953
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
                  ButtonImage     =   "FrmManGoBack.frx":1DC2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
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
                  Index           =   18
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   990
                  Width           =   780
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
                  Index           =   16
                  Left            =   390
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   960
                  Width           =   1290
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
                  Index           =   15
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   960
                  Width           =   1980
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
                  Index           =   10
                  Left            =   7125
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   195
                  Index           =   9
                  Left            =   1395
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   30
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—ﬁ «Ê «· Œ’Ì„ "
                  Height          =   270
                  Index           =   7
                  Left            =   3420
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   660
                  Width           =   1185
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
                  Index           =   5
                  Left            =   3090
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   960
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ—«— «·„Ê—œ"
                  Height          =   420
                  Index           =   23
                  Left            =   8250
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   660
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   225
                  Index           =   28
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   30
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   225
                  Index           =   30
                  Left            =   4845
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   30
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   225
                  Index           =   31
                  Left            =   7110
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   30
                  Width           =   1050
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· ﬂ "
                  Height          =   195
                  Index           =   11
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   30
                  Width           =   1050
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2685
               Left            =   0
               TabIndex        =   39
               Top             =   1530
               Width           =   8880
               _cx             =   15663
               _cy             =   4736
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
               FormatString    =   $"FrmManGoBack.frx":215C
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
               Height          =   600
               Left            =   465
               TabIndex        =   40
               Top             =   4215
               Width           =   8415
               _ExtentX        =   14843
               _ExtentY        =   1058
               ButtonWidth     =   609
               ButtonHeight    =   953
               Appearance      =   1
               _Version        =   393216
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4620
            Index           =   6
            Left            =   9615
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   45
            Width           =   8880
            _cx             =   15663
            _cy             =   8149
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
            _GridInfo       =   $"FrmManGoBack.frx":2365
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄ «·‰ﬁœÏ"
               Height          =   675
               Index           =   0
               Left            =   1605
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   630
               Width           =   6840
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5040
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   210
                  Width           =   1155
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3090
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   180
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   71
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
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6180
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   240
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   210
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   270
                  Index           =   22
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   210
                  Width           =   765
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄ «·√Ã·"
               Height          =   645
               Index           =   1
               Left            =   1605
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1665
               Width           =   6840
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   5070
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   210
                  Width           =   1035
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   3120
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   180
                  Width           =   1185
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   360
                  Left            =   120
                  TabIndex        =   64
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
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   210
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   17
                  Left            =   6240
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   270
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   210
                  Width           =   1155
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   630
               Index           =   20
               Left            =   7665
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   7
         Left            =   15
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   15
         Width           =   8970
         _cx             =   15822
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
         Caption         =   "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ"
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
         _GridInfo       =   $"FrmManGoBack.frx":240B
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1335
            TabIndex        =   42
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
            ButtonImage     =   "FrmManGoBack.frx":24B3
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
            TabIndex        =   43
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
            ButtonImage     =   "FrmManGoBack.frx":284D
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
            TabIndex        =   44
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
            ButtonImage     =   "FrmManGoBack.frx":2BE7
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
            TabIndex        =   45
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
            ButtonImage     =   "FrmManGoBack.frx":2F81
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   7425
         Width           =   8970
         _cx             =   15822
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
            Left            =   8085
            TabIndex        =   47
            Top             =   90
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   7035
            TabIndex        =   48
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
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
            Left            =   6015
            TabIndex        =   49
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   5100
            TabIndex        =   50
            Top             =   90
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   3915
            TabIndex        =   51
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   2985
            TabIndex        =   52
            Top             =   90
            Width           =   825
            _ExtentX        =   1455
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
            TabIndex        =   53
            Top             =   120
            Width           =   810
            _ExtentX        =   1429
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
            Left            =   1935
            TabIndex        =   54
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
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
            Left            =   915
            TabIndex        =   55
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
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
   End
End
Attribute VB_Name = "FrmManGoBack1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim MaintenReport As ClsMaintananceReport
Dim cSearchDcbo(6) As clsDCboSearch

Public BolPrint As Boolean

Private Sub CboSupDeci_Change()

    If Me.CboSupDeci.ListIndex = -1 Then
        Exit Sub
    ElseIf Me.CboSupDeci.ListIndex = 0 Then
        '—›÷ Œ—ÊÃ „‰ «·÷„«‰
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = False
        Me.TxtCost.Enabled = False
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    ElseIf Me.CboSupDeci.ListIndex = 1 Then
        '—›÷ ≈‰ Â  › —… «·÷„«‰
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = False
        Me.TxtCost.Enabled = False
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    
    ElseIf Me.CboSupDeci.ListIndex = 2 Then
        '≈” »œ· »ﬁÿ⁄… «Œ—Ï
        Me.lbl(5).Enabled = True
        Me.TxtNewSerial.Enabled = True
        Me.lbl(7).Enabled = False
        Me.TxtCost.Enabled = False
    
        Me.CmdSearch.Enabled = True
        DcboReItemCode.Enabled = True
        Me.DcboReItemName.Enabled = True
        TxtNewSerial.Enabled = True
        TxtReItemQty.Enabled = True
        DcboReStore.Enabled = True
        Me.lbl(10).Enabled = True
        Me.lbl(15).Enabled = True
        Me.lbl(5).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(16).Enabled = True
    
    ElseIf Me.CboSupDeci.ListIndex = 3 Then
        '≈” »œ«· „⁄ œ›⁄ ›—ﬁ ”⁄—
        Me.lbl(5).Enabled = True
        Me.TxtNewSerial.Enabled = True
        Me.lbl(7).Enabled = True
        Me.TxtCost.Enabled = True
    
        Me.CmdSearch.Enabled = True
        DcboReItemCode.Enabled = True
        Me.DcboReItemName.Enabled = True
        TxtNewSerial.Enabled = True
        TxtReItemQty.Enabled = True
        DcboReStore.Enabled = True
        Me.lbl(10).Enabled = True
        Me.lbl(15).Enabled = True
        Me.lbl(5).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(16).Enabled = True
    
    ElseIf Me.CboSupDeci.ListIndex = 4 Then
        ' Œ’Ì„ ⁄·Ï «·„Ê—œ
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = True
        Me.TxtCost.Enabled = True
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    
    ElseIf Me.CboSupDeci.ListIndex = 5 Then
        ' „ «· ’·ÌÕ œ«Œ· «·÷„«‰
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = False
        Me.TxtCost.Enabled = False
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    
    ElseIf Me.CboSupDeci.ListIndex = 6 Then
        ' „ «· ’·ÌÕ » ﬂ·›…
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = True
        Me.TxtCost.Enabled = True
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    ElseIf Me.CboSupDeci.ListIndex = 7 Then
        '≈” »œ«· „⁄ «Œ–  ›—ﬁ ”⁄—
        Me.lbl(5).Enabled = True
        Me.TxtNewSerial.Enabled = True
        Me.lbl(7).Enabled = True
        Me.TxtCost.Enabled = True
    
        Me.CmdSearch.Enabled = True
        DcboReItemCode.Enabled = True
        Me.DcboReItemName.Enabled = True
        TxtNewSerial.Enabled = True
        TxtReItemQty.Enabled = True
        DcboReStore.Enabled = True
        Me.lbl(10).Enabled = True
        Me.lbl(15).Enabled = True
        Me.lbl(5).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(16).Enabled = True
    
    Else
        Me.lbl(5).Enabled = False
        Me.TxtNewSerial.Enabled = False
        Me.lbl(7).Enabled = False
        Me.TxtCost.Enabled = False
    
        Me.CmdSearch.Enabled = False
        DcboReItemCode.Enabled = False
        Me.DcboReItemName.Enabled = False
        TxtNewSerial.Enabled = False
        TxtReItemQty.Enabled = False
        DcboReStore.Enabled = False
        Me.lbl(10).Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(5).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(16).Enabled = False
    
    End If

End Sub

Private Sub CboSupDeci_Click()
    CboSupDeci_Change
End Sub

Private Sub CmdAdd_Click()
    '“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·
    Dim LngTransID As Long
    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngFindRow As Long
    Dim LngRow As Long

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

    If Me.CboSupDeci.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ﬁ—«— «·„Ê—œ »‘√‰ Â–Â «·ﬁÿ⁄… „‰ «·’‰› ....!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboSupDeci.SetFocus
        Exit Sub
    ElseIf Me.CboSupDeci.ListIndex = 2 Or Me.CboSupDeci.ListIndex = 3 Or Me.CboSupDeci.ListIndex = 7 Then

        'Õ«·… ≈” »œ«· ·«»œ „‰ ÊÃÊœ «·’‰› «·„” »œ·
        If Me.DcboReItemCode.BoundText = "" Then
            Msg = "·«»œ „‰  ÕœÌœ ﬂÊœ «·’‰› «·„” »œ·....!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboReItemCode.SetFocus
            Exit Sub
        End If

        If Me.DcboReItemName.BoundText = "" Then
            Msg = "·«»œ „‰  ÕœÌœ «”„ «·’‰› «·„” »œ·....!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboReItemName.SetFocus
            Exit Sub
        End If

        If Me.TxtNewSerial.Enabled = True And Me.TxtNewSerial.text = "" Then
            Msg = "·«»œ „‰ ≈œŒ«· ”Ì—»«· «·’‰› «·„” »œ· ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNewSerial.SetFocus
            Exit Sub
        End If

        If val(Me.TxtReItemQty.text) = 0 Then
            Msg = "·«»œ „‰ ≈œŒ«· ﬂ„Ì… «·’‰› «·„” »œ· ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtReItemQty.SetFocus
            Exit Sub
        End If

    ElseIf Me.CboSupDeci.ListIndex = 1 Then
    ElseIf Me.CboSupDeci.ListIndex = 2 Then
    ElseIf Me.CboSupDeci.ListIndex = 3 Then
    ElseIf Me.CboSupDeci.ListIndex = 4 Then
    ElseIf Me.CboSupDeci.ListIndex = 5 Then
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

    StrSQL = "SELECT QryManSupStockComplete.* "
    StrSQL = StrSQL + " FROM dbo.QryManSupStockComplete(" & LngTransID & ") QryManSupStockComplete"
    StrSQL = StrSQL + " Where ItemID=" & Me.DCboItemsCode.BoundText
    StrSQL = StrSQL + " AND CusID=" & Me.DBCboClientName.BoundText

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

    '----------------------------------------------------
    If Trim$(Me.TxtTicketNO.text) <> "" Then
        LngFindRow = Fg.FindRow(Trim$(Me.TxtTicketNO.text), Fg.FixedRows, Fg.ColIndex("TicketNO"), False, True)
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
        .TextMatrix(LngRow, .ColIndex("SupDeci")) = Me.CboSupDeci.ListIndex + 1
        .TextMatrix(LngRow, .ColIndex("Cost")) = val(Me.TxtCost.text)

        If TxtSerial.Tag = "T" Then
            .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexChecked
        ElseIf TxtSerial.Tag = "F" Then
            .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexUnchecked
        End If

        .AutoSize 0, .Cols - 1, False

        If Me.DcboReItemName.BoundText <> "" Then
            LngFindRow = Fg.FindRow("Rep-" & LngRow, Fg.FixedRows, Fg.ColIndex("RowFlag"), False, True)

            If LngFindRow = -1 Then
                .AddItem "", LngRow + 1
            Else
            End If

            .TextMatrix(LngRow + 1, .ColIndex("Name")) = Me.DcboReItemName.BoundText
            .TextMatrix(LngRow + 1, .ColIndex("Code")) = Me.DcboReItemName.BoundText
            .TextMatrix(LngRow + 1, .ColIndex("Serial")) = Trim$(Me.TxtNewSerial.text)

            If TxtNewSerial.Enabled = True Then
                .Cell(flexcpChecked, LngRow + 1, .ColIndex("HaveSerial")) = flexChecked
            ElseIf TxtNewSerial.Enabled = False Then
                .Cell(flexcpChecked, LngRow + 1, .ColIndex("HaveSerial")) = flexUnchecked
            End If

            .TextMatrix(LngRow + 1, .ColIndex("Count")) = Trim(Me.TxtReItemQty.text)
            .TextMatrix(LngRow + 1, .ColIndex("TicketNO")) = .TextMatrix(LngRow, .ColIndex("TicketNO"))
            .TextMatrix(LngRow + 1, .ColIndex("SupDeci")) = "≈” »œ«· „‰ " & Me.DcboReStore.text
            .Cell(flexcpData, LngRow + 1, .ColIndex("SupDeci")) = Me.DcboReStore.BoundText
            .TextMatrix(LngRow + 1, .ColIndex("ManDetID")) = LngRow
            .TextMatrix(LngRow + 1, .ColIndex("RowFlag")) = "Rep-" & LngRow
            .Cell(flexcpBackColor, LngRow + 1, 1, LngRow + 1, .Cols - 1) = vbGreen
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    '--------------------------------------------------------
    Me.XPTxtSum.text = Fg.Aggregate(flexSTSum, Fg.FixedRows, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
    DCboItemsCode.BoundText = ""
    DCboItemsName.BoundText = ""
    TxtSerial.text = ""
    Me.TxtCost.text = ""
    Me.DcboReItemCode.BoundText = ""
    Me.DcboReItemName.BoundText = ""
    Me.TxtNewSerial.text = ""
    Me.TxtReItemQty.text = ""
    Me.DcboReStore.BoundText = ""

    Fg.SetFocus
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DcboReItemCode
    FrmItemSearch.show vbModal
End Sub

Private Sub CmdShowTransItems_Click()

    If Me.DBCboClientName.BoundText = "" Then
        Exit Sub
    End If

    Load FrmManChooseItems
    Set FrmManChooseItems.MyForm = Me
    FrmManChooseItems.ShowManSupStock Me.DBCboClientName.BoundText, Me.DBCboClientName.text
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

    StrSQL = "select * From TblItems where ItemID=" & DcboReItemCode.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtNewSerial.Enabled = True
            Me.TxtReItemQty.Enabled = False
            Me.TxtReItemQty.text = 1
            TxtNewSerial.Tag = "T"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtNewSerial.Enabled = False
            Me.TxtReItemQty.Enabled = True
            Me.TxtReItemQty.text = 1
            TxtNewSerial.Tag = "F"
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

    StrSQL = "select * From TblItems where ItemID=" & DcboReItemName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            TxtNewSerial.Enabled = True
            TxtReItemQty.Enabled = False
            TxtReItemQty.text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtNewSerial.Enabled = False
            TxtReItemQty.Enabled = True
            TxtReItemQty.text = ""
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

'    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'        FrmAddNewItem.DealingForm = Maintenance
'        FrmAddNewItem.show vbModal
'    End If

End Sub

Private Sub Fg_Click()
    Dim i As Long
    Dim LngLoadRow As Long
    Dim StrTemp  As String
    Dim LngTemp  As Long

    On Error GoTo ErrTrap

    With Fg

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
                Me.CboSupDeci.ListIndex = val(Me.Fg.TextMatrix(.Row, .ColIndex("SupDeci"))) - 1
                Me.TxtCost.text = val(Me.Fg.TextMatrix(.Row, .ColIndex("Cost")))
                LngLoadRow = .Row

                For i = .FixedRows To .Rows - 1

                    If Trim$(.TextMatrix(i, .ColIndex("RowFlag"))) <> "" Then
                        StrTemp = Trim$(.TextMatrix(i, .ColIndex("RowFlag")))
                        LngTemp = val(Mid$(StrTemp, InStr(1, StrTemp, "-", vbTextCompare) + 1))

                        If LngTemp = LngLoadRow Then
                            Me.DcboReItemCode.BoundText = .TextMatrix(i, .ColIndex("Name"))
                            Me.DcboReItemName.BoundText = .TextMatrix(i, .ColIndex("Name"))
                            Me.TxtNewSerial.text = .TextMatrix(i, .ColIndex("Serial"))
                            Me.TxtReItemQty.text = .TextMatrix(i, .ColIndex("Count"))
                            Me.DcboReStore.BoundText = val(.Cell(flexcpData, i, .ColIndex("SupDeci")))
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
    Dim RsTemp As ADODB.Recordset

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
    Me.Height = 8580
    Me.Width = 9700
    Resize_Form Me
    'AddTip
    SetDtpickerDate Me.XPDtbGoInDtae
    SetDtpickerDate Me.DtpDelayDate
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboEmp

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    '-------------------------------
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open "Select * From TblManSupDecs Where DecType=1 Order By SupDecID", Cn, adOpenStatic, adLockReadOnly, adCmdText
    '-------------------------------
    LoadTBR

    '-------------------------------
    With Me.CboSupDeci
        .Clear
        RsTemp.MoveFirst

        Do While Not RsTemp.EOF
            .AddItem RsTemp("SupDecName").value
            .ItemData(.NewIndex) = RsTemp("SupDecID").value
            RsTemp.MoveNext
        Loop

    End With

    Fg.WallPaper = BGround.Picture

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

    Fg.ColComboList(Fg.ColIndex("MType")) = "#1;»«· ﬂ·›…|#2; »⁄ «·÷„«‰"
    RsTemp.MoveFirst
    StrList = Fg.BuildComboList(RsTemp, "SupDecName", "SupDecID")

    If Trim$(StrList) <> "" Then
        StrList = "|" & StrList
    End If

    Fg.ColComboList(Fg.ColIndex("SupDeci")) = StrList

    Set rs = New ADODB.Recordset
    rs.Open "Select * From  TblMaintenece Where ManOperationTypeID=3", Cn, adOpenStatic, adLockOptimistic, adCmdText
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
            Me.Caption = "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ"
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

        Case "N"
            Me.Caption = "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ( ÃœÌœ )"
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

        Case "E"
            Me.Caption = "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ(  ⁄œÌ· )"
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
    End Select

    Exit Sub
ErrTrap:
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
            FrmMaintanenceSearch.searchtype = 3
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

''    With FrmAddNewCustemer
 '       .DealingForm = Maintenance
 '       .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
 '       .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
 '       .lbl(0).Caption = "«”„ «·⁄„Ì·"
 '       .show vbModal
 '       cSearchDcbo(0).Refresh
 '   End With

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
    Dim RsManDetailReplace As ADODB.Recordset
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim ReplaceID As Integer
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim note_id As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If DBCboClientName.text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ...!!!"
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

        If ItemsInGrid(Fg, Fg.ColIndex("Name")) = -1 Then
            Msg = "ÌÃ» ≈Œ Ì·— «·√’‰«› «·„— Ã⁄… „‰ «·’„«‰...!!! " & Chr(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
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
        End If

        RsDetails.Open "[TblMainteneceDetails]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Set RsManDetailReplace = New ADODB.Recordset
        RsManDetailReplace.Open "TblManDetailsReplacedItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        Cn.BeginTrans
        BeginTrans = True
        rs("MaintananceID").value = val(XPTxtMaintanenceID.text)
        rs("CusID").value = Me.DBCboClientName.BoundText ' IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
        rs("DateGoIN").value = XPDtbGoInDtae.value
        rs("DateGoOUT").value = Null
        rs("GoOut").value = 0
        rs("EmpID").value = Me.DcboEmp.BoundText
        rs("StoreID").value = Me.DCboStoreName.BoundText
        rs("UserID").value = user_id
        rs("MType").value = 0
        rs("Transaction_ID").value = Null
        rs("ManOperationTypeID").value = 3
        rs.update

        For RowNum = 1 To Fg.Rows - 1

            If Fg.TextMatrix(RowNum, Fg.ColIndex("RowFlag")) = "" Then
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
                RsDetails("NewSerial").value = Null
                RsDetails.update
                Fg.TextMatrix(RowNum, Fg.ColIndex("TableID")) = RsDetails("TableID").value
            ElseIf Fg.TextMatrix(RowNum, Fg.ColIndex("RowFlag")) Like "Rep-" & "*" Then
                Dim StrTemp  As String
                Dim X As Integer, LngRow As Integer
            
                StrTemp = Fg.TextMatrix(RowNum, Fg.ColIndex("RowFlag"))
                X = Len("Rep-")
                LngRow = val(Mid$(StrTemp, X + 1))
            
                RsManDetailReplace.AddNew
                RsManDetailReplace("ManDetID").value = Fg.TextMatrix(LngRow, Fg.ColIndex("TableID"))
                RsManDetailReplace("ItemID").value = Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))

                If Fg.Cell(flexcpChecked, RowNum, Fg.ColIndex("HaveSerial")) = flexChecked Then
                    RsManDetailReplace("ItemSerial").value = IIf(IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))), "", Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))))
                Else
                    RsManDetailReplace("ItemSerial").value = Null
                End If

                RsManDetailReplace("ItemQty").value = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))
                RsManDetailReplace("StoreID").value = val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("SupDeci")))
                RsManDetailReplace("ReplaceType").value = 1
                RsManDetailReplace.update
                Fg.TextMatrix(RowNum, Fg.ColIndex("TableID")) = RsManDetailReplace("ID").value
        
            End If

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
    Dim RsNotes As New ADODB.Recordset
    Dim RsDetails As New ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Long
    Dim i As Integer
    Dim IntRow As Integer

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

    Me.DcboEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)

    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT TblItems.HaveSerial,* FROM TblItems INNER JOIN TblMainteneceDetails " & "ON TblItems.ItemID = TblMainteneceDetails.ItemID"
    StrSQL = StrSQL + "  where MaintananceID=" & val(rs("MaintananceID").value)
    'StrSql = "select * From TblMainteneceDetails where MaintananceID=" & Val(Rs("MaintananceID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = Fg.FixedRows
        Fg.Rows = Fg.FixedRows

        For i = 1 To RsDetails.RecordCount
            Fg.AddItem ""
            IntRow = Fg.Rows - 1
            Fg.Cell(flexcpPicture, IntRow, Fg.ColIndex("Replace")) = ""
            Fg.Cell(flexcpData, IntRow, Fg.ColIndex("Replace")) = ""
            Fg.TextMatrix(IntRow, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            Fg.TextMatrix(IntRow, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            Fg.TextMatrix(IntRow, Fg.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial").value = True Then
                Fg.TextMatrix(IntRow, Fg.ColIndex("HaveSerial")) = True
            Else
                Fg.TextMatrix(IntRow, Fg.ColIndex("HaveSerial")) = False
            End If        '

            Fg.TextMatrix(IntRow, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", Trim(RsDetails("Quantity").value))
            Fg.TextMatrix(IntRow, Fg.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").value))
            Fg.TextMatrix(IntRow, Fg.ColIndex("SupDeci")) = IIf(IsNull(RsDetails("SupDeci")), "", val(RsDetails("SupDeci").value))
            Fg.TextMatrix(IntRow, Fg.ColIndex("Cost")) = IIf(IsNull(RsDetails("Cost")), "", val(RsDetails("Cost").value))
            '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «· Ì  „  ›Ì Â–Â «·⁄„·Ì…
            Set RsReplace = New ADODB.Recordset
            StrSQL = "SELECT     dbo.TblManDetailsReplacedItems.ID, dbo.TblManDetailsReplacedItems.Man" & "DetID, dbo.TblManDetailsReplacedItems.ItemID,dbo.TblManD" & "etailsReplacedItems.ItemSerial, dbo.TblManDetailsReplacedItems.ItemQty, dbo.TblM" & "anDetailsReplacedItems.StoreID,dbo.TblItems.HaveSerial, " & "dbo.TblStore.StoreName FROM         dbo.TblItems INNER JOIN                     " & "  dbo.TblManDetailsReplacedItems ON dbo.TblItems.ItemID = dbo.TblManDetailsRepla" & "cedItems.ItemID INNER JOIN dbo.TblStore ON dbo.TblManDetai" & "lsReplacedItems.StoreID = dbo.TblStore.StoreID"
        
            StrSQL = StrSQL + " Where ManDetID=" & RsDetails("TableID").value
            RsReplace.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsReplace.BOF Or RsReplace.EOF) Then
                Fg.AddItem ""
                IntRow = Fg.Rows - 1
                Fg.TextMatrix(IntRow, Fg.ColIndex("TableID")) = RsReplace("ID").value
                Fg.Cell(flexcpPicture, IntRow, Fg.ColIndex("Replace")) = ""
                Fg.Cell(flexcpData, IntRow, Fg.ColIndex("Replace")) = ""
                Fg.TextMatrix(IntRow, Fg.ColIndex("Code")) = IIf(IsNull(RsReplace("ItemID")), "", Trim(RsReplace("ItemID").value))
                Fg.TextMatrix(IntRow, Fg.ColIndex("Name")) = IIf(IsNull(RsReplace("ItemID")), "", Trim(RsReplace("ItemID").value))
                Fg.TextMatrix(IntRow, Fg.ColIndex("Serial")) = IIf(IsNull(RsReplace("ItemSerial")), "", Trim(RsReplace("ItemSerial").value))

                If RsReplace("HaveSerial").value = True Then
                    Fg.Cell(flexcpChecked, IntRow, Fg.ColIndex("HaveSerial")) = flexChecked
                Else
                    Fg.Cell(flexcpChecked, IntRow, Fg.ColIndex("HaveSerial")) = flexUnchecked
                End If

                Fg.TextMatrix(IntRow, Fg.ColIndex("Count")) = IIf(IsNull(RsReplace("ItemQty")), "", Trim(RsReplace("ItemQty").value))
                Fg.TextMatrix(IntRow, Fg.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").value))
            
                Fg.TextMatrix(IntRow, Fg.ColIndex("SupDeci")) = "≈” »œ·  ›Ï " & IIf(IsNull(RsReplace("StoreName")), "", Trim(RsReplace("StoreName").value))
                Fg.Cell(flexcpData, IntRow, Fg.ColIndex("SupDeci")) = IIf(IsNull(RsReplace("StoreID")), "", Trim(RsReplace("StoreID").value))
                Fg.TextMatrix(IntRow, Fg.ColIndex("ManDetID")) = IIf(IsNull(RsReplace("ManDetID")), "", Trim(RsReplace("ManDetID").value))
                Fg.TextMatrix(IntRow, Fg.ColIndex("RowFlag")) = "Rep-" & IntRow - 1
                Fg.Cell(flexcpBackColor, IntRow, 1, IntRow, Fg.Cols - 1) = vbGreen
            End If

            RsDetails.MoveNext
        Next i

        Fg.AutoSize 0, Fg.Cols - 1, False
        Me.XPTxtSum.text = Fg.Aggregate(flexSTSum, Fg.FixedRows, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
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
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
        
            Else
        
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
    Dim Dcombos As ClsDataCombos

    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsCodes Me.DCboItemsCode
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

